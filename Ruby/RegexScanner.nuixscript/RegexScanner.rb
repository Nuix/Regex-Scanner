script_directory = File.dirname(__FILE__)

# Load up Nx.jar for settings dialog and progress dialog
require File.join(script_directory,"Nx.jar")
java_import "com.nuix.nx.NuixConnection"
java_import "com.nuix.nx.LookAndFeelHelper"
java_import "com.nuix.nx.dialogs.ChoiceDialog"
java_import "com.nuix.nx.dialogs.CustomDialog"
java_import "com.nuix.nx.dialogs.TabbedCustomDialog"
java_import "com.nuix.nx.dialogs.CommonDialogs"
java_import "com.nuix.nx.dialogs.ProgressDialog"
java_import "com.nuix.nx.dialogs.ProcessingStatusDialog"
java_import "com.nuix.nx.digest.DigestHelper"
java_import "com.nuix.nx.controls.models.Choice"

LookAndFeelHelper.setWindowsIfMetal
NuixConnection.setUtilities($utilities)
NuixConnection.setCurrentNuixVersion(NUIX_VERSION)

# Load super utilities for RegexScanner
require File.join(script_directory,"SuperUtilities.jar")
java_import "com.nuix.superutilities.SuperUtilities"
$su = SuperUtilities.init($utilities,NUIX_VERSION)
# Load excel helper class for reporting
load File.join(script_directory,"Xlsx.rb")
# For thread safety
require 'thread'

# Important Note: With this set to true, items are scanned using a Java Parallel Stream, this likely means faster
# scan times with the trade off of increased memory pressure because now you have more than 1 thread in memory
# scanning more than one item at a time.  The result of this means that the Java heap may grow faster than GC
# can keep up with, which means the potential for memory related errors.  Setting this to false means slower scan
# times but lower memory pressures.
$scan_in_parallel = true

# When $scan_in_parallel = true, this will determine how many scans are performed concurrently
# A higher number will likely yield faster results but also more concurrently resource usage, which means a potentially
# increased chance of memory shortage while running.  A higher value is not always better!
$scan_concurrency = 4

# To get the text of an item from the API you call Item.getTextObject, this in returns a Nuix object named Text
# To work with this text you have 2 choices:
# - call Text.toString to convert to a String
# - Use the Text object as a CharSequence (which it extends)
# 
# It appears that running a Java Pattern/Matcher (regex classes) over a String performs better, presumably because the
# entirety of the text is in memory, but this also means increased memory usage.  Using the Text object as a CharSequence
# seems to exhibit reduced Garbage Collection pressure, but is also slower, presumably because the text is being read into memory
# in segments as needed.  This setting allows you to control the threshold of when the Text object is first converted
# to a string vs used as a CharSequence.
com.nuix.superutilities.regex.RegexScanner.setMaxToStringLength(1024 * 1024 * 5)

# Are we going to run this against a selection of items or
# all items?
items = nil
if !$current_selected_items.nil? && $current_selected_items.size > 0
	items = $current_selected_items
else
	items = $current_case.searchUnsorted("")
end

# CSV reports include a timestemp so we need to capture one
filename_timestamp = $su.getFormatUtility.getFilenameTimestamp

property_choices = $current_case.getMetadataItems.map{|mi|mi.getName}.uniq.sort.map{|name| Choice.new(name,name,"Metadata property: #{name}",true)}

# Setup the settings dialog
dialog = TabbedCustomDialog.new("Regex Scanner 2")
dialog.setHelpFile(File.join(script_directory,"Help.html"))

expressions_tab = dialog.addTab("expressions_tab","Regular Expressions")
expressions_tab.appendHeader("Items: #{items.size}")
expressions_tab.appendCsvTable("regex_expressions",["Title","Regex"])

scan_settings_tab = dialog.addTab("scan_settings_tab","Scan Settings")
scan_settings_tab.appendCheckBox("skip_excluded_items","Skip Excluded Items",true)
scan_settings_tab.appendCheckBox("case_sensitive","Expressions are Case Sensitive",false)
scan_settings_tab.appendCheckBox("capture_context","Capture Match Value Context",false)
scan_settings_tab.appendSpinner("context_size_chars","Context Size in Characters",100,1,1000,5)
scan_settings_tab.enabledOnlyWhenChecked("context_size_chars","capture_context")
scan_settings_tab.appendCheckBox("scan_content","Scan Item Content",true)
scan_settings_tab.appendCheckBox("scan_properties","Scan Item Properties",true)
scan_settings_tab.appendHeader("Properties to Scan")
scan_settings_tab.appendChoiceTable("properties_to_scan","Properties to Scan",property_choices)
scan_settings_tab.enabledOnlyWhenChecked("properties_to_scan","scan_properties")

reporting_tab = dialog.addTab("reporting_tab","Reporting")
reporting_tab.appendCheckableTextField("apply_tags",false,"tag_template","RegexScannerMatch|{location}|{title}","Apply Tags")
reporting_tab.appendCheckableTextField("apply_custom_metadata",false,"field_name_template","RegexScannerMatch_{location}_{title}","Apply Custom Metadata")

reporting_tab.appendCheckBox("generate_report_csv","Generate CSV Report",false)
reporting_tab.appendDirectoryChooser("report_csv_directory","CSV Report Directory")
reporting_tab.enabledOnlyWhenChecked("report_csv_directory","generate_report_csv")

reporting_tab.appendCheckBox("generate_report_xlsx","Generate XLSX Report",false)
reporting_tab.appendSaveFileChooser("report_xlsx_file","XLSX Report File","Excel XLSX","xlsx")
reporting_tab.enabledOnlyWhenChecked("report_xlsx_file","generate_report_xlsx")
reporting_tab.appendCheckBox("include_item_path","Include Item Path",false)
reporting_tab.appendCheckBox("include_physical_path","Include Physical Ancestors's Path",false)

# Define settings validation performed by the settings dialog
dialog.validateBeforeClosing do |values|
	# Make sure at least one Regex was provided
	if values["regex_expressions"].size < 1
		CommonDialogs.showWarning("Please provide at least one regular expression.")
		next false
	end

	# Make sure all entries have a regular expression and a title
	all_entries_populated = true
	values["regex_expressions"].each_with_index do |expression_entry,index|
		if expression_entry["Title"].strip.empty?
			CommonDialogs.showWarning("Please provide a title for entry #{index+1}")
			all_entries_populated = false
			break
		end

		if expression_entry["Regex"].strip.empty?
			CommonDialogs.showWarning("Please provide an expression for entry #{index+1}")
			all_entries_populated = false
			break
		end
	end
	next false if !all_entries_populated

	# Make sure all regular expressions compile (test them for errors)
	all_compiled = true
	values["regex_expressions"].each_with_index do |expression,index|
		begin
			java.util.regex.Pattern.compile(expression["Regex"])
		rescue Exception => exc
			all_compiled = false
			CommonDialogs.showError("Error in expression #{index+1}:\n\n#{exc.message}")
		end
		break if !all_compiled
	end
	next false if !all_compiled

	# If we are scanning metadata properties, make sure some properties to be scanned are actually selected.
	if values["scan_properties"] && values["properties_to_scan"].size < 1
		CommonDialogs.showWarning("You have selected scan properties, but selected no properties to scan.  Please select"+
			" at least one property to scan.")
		next false
	end

	# Make sure a CSV directory was picked if we are reporting to CSV
	if values["generate_report_csv"] && values["report_csv_directory"].strip.empty?
		CommonDialogs.showWarning("Please provide a value for 'CSV Report Directory'")
		next false
	end

	# Notify user that since they are applying custom metadata we need to close all tabs
	if values["apply_custom_metadata"] && NuixConnection.getCurrentNuixVersion.isAtLeast("6.2")
		# Get user confirmation about closing all workbench tabs
		if CommonDialogs.getConfirmation("The script needs to close all workbench tabs, proceed?") == false
			next false
		end
	end

	next true
end

# Display the settings dialog
dialog.display
# If the user clicked Ok then lets proceed
if dialog.getDialogResult == true
	# Get settings dialog values
	values = dialog.toMap

	# Extract settings dialog values into variables for convenience
	regex_expressions = values["regex_expressions"]

	scan_properties = values["scan_properties"]
	properties_to_scan = values["properties_to_scan"]
	scan_content = values["scan_content"]
	case_sensitive = values["case_sensitive"]
	capture_context = values["capture_context"]
	context_size_chars = values["context_size_chars"]

	apply_tags = values["apply_tags"]
	tag_template = values["tag_template"]

	apply_custom_metadata = values["apply_custom_metadata"]
	field_name_template = values["field_name_template"]

	generate_report_csv = values["generate_report_csv"]
	report_csv_directory = values["report_csv_directory"]

	generate_report_xlsx = values["generate_report_xlsx"]
	report_xlsx_file = values["report_xlsx_file"]

	include_physical_path = values["include_physical_path"]
	include_item_path = values["include_item_path"]

	skip_excluded_items = values["skip_excluded_items"]

	# Close all tabs if we can and need to
	if NuixConnection.getCurrentNuixVersion.isAtLeast("6.2") && apply_custom_metadata
		$window.closeAllTabs
	end

	# Show a progress dialog as we are about to get to work
	ProgressDialog.forBlock do |pd|
		pd.setTitle("Regex Scanner 2")

		# When a message is logged to the progress dialog we would like
		# it written to standard output and logs
		pd.onMessageLogged do |message|
			puts message
		end

		pd.logMessage("Scanning in Parallel: #{$scan_in_parallel}")
		if $scan_in_parallel
			pd.logMessage("Concurrency: #{$scan_concurrency}")
		end

		if skip_excluded_items
			pd.logMessage("Removing any excluded items present in input items...")
			item_count_before = items.size
			pd.logMessage("Items Before: #{item_count_before}")
			items = items.reject{|item| item.isExcluded}
			item_count_after = items.size
			pd.logMessage("Items After: #{item_count_after}")
			pd.logMessage("Excluded Items Removed: #{item_count_before - item_count_after}")
		end

		# Configure our RegexScanner instance from SuperUtilities
		scanner = $su.createRegexScanner
		scanner.setScanProperties(scan_properties)
		scanner.setPropertiesToScan(properties_to_scan)
		scanner.setScanContent(scan_content)
		scanner.setCaseSensitive(case_sensitive)
		scanner.setCaptureContextualText(capture_context)
		scanner.setContextSize(context_size_chars)
		start_time = Time.now

		regex_expressions.each do |record|
			scanner.addPattern(record["Title"],record["Regex"])
		end

		# Setup some variables we will be using to track
		# various information while scanning
		matched_item_count = 0
		matched_value_count = 0
		error_count = 0
		last_progress = Time.now
		last_status_update = Time.now
		tag_grouped = Hash.new{|h,k| h[k] = [] }

		pd.setMainProgress(0,items.size)

		# Hookup a callback to regex scanner progress so that progress dialog is updated
		scanner.whenProgressUpdated do |value|
			elapsed_seconds = (Time.now - start_time).to_i
			elapsed = $su.getFormatUtility.secondsToElapsedString(elapsed_seconds)
			message = "#{elapsed} #{value}/#{items.size}, Matched Items: #{matched_item_count}, Matches: #{matched_value_count}, Errors: #{error_count}"

			# Update the main status 16x a second (save some cycles by throttling it a bit)
			# if (Time.now - last_status_update) > 2
				pd.setMainStatus("Scanning #{message}")
				last_status_update = Time.now
				pd.setMainProgress(value)
			# end

			# Record status every 5 seconds to log area
			if (Time.now - last_progress) > 5
				pd.logMessage("Scanned #{message}")
				last_progress = Time.now
			end
		end

		# Define headers for error reporting
		errors_headers = [
			"Title",
			"Expression",
			"Location",
			"Item GUID",
			"Error Message",
		]

		# Define headers for match reporting
		matches_headers = [
			"GUID",
			"Item Name",
			"ItemKind",
			"Expression Title",
			"Location",
			"Value",
			"ValueContext",
			"Match Start",
			"Match End",
		]

		matches_headers << "ItemPath" if include_item_path
		matches_headers << "PhysicalFilePath" if include_physical_path

		matches_csv = nil
		xlsx_report = nil
		matches_sheet = nil

		# Initialize CSV reporting
		if generate_report_csv
			require 'csv'
			java.io.File.new(report_csv_directory).mkdirs
			summary_csv_file = File.join(report_csv_directory,"#{filename_timestamp}_RegexSummary.csv").gsub("/","\\")
			matches_csv_file = File.join(report_csv_directory,"#{filename_timestamp}_RegexMatches.csv").gsub("/","\\")
			errors_csv_file = File.join(report_csv_directory,"#{filename_timestamp}_RegexScanErrors.csv").gsub("/","\\")

			CSV.open(summary_csv_file,"w:utf-8") do |summary_csv|
				summary_csv << [
					"Title",
					"Regular Exporession",
				]

				scanner.getPatterns.each do |pattern_info|
					summary_csv << [
						pattern_info.getTitle,
						pattern_info.getExpression,
					]
				end
			end

			matches_csv = CSV.open(matches_csv_file,"w:utf-8")
			matches_csv << matches_headers

			errors_csv = CSV.open(errors_csv_file,"w:utf-8")
			errors_csv << errors_headers
		end

		# Initialize XLSX reporting
		if generate_report_xlsx
			java.io.File.new(report_xlsx_file).getParentFile.mkdirs
			xlsx_report = Xlsx.new

			summary_sheet = xlsx_report.get_sheet("Summary")
			summary_sheet << [
				"Title",
				"Regular Exporession",
			]
			scanner.getPatterns.each do |pattern_info|
				summary_sheet << [
					pattern_info.getTitle,
					pattern_info.getExpression,
				]
			end

			matches_sheet = xlsx_report.get_sheet("Matches")
			matches_sheet << matches_headers

			errors_sheet = xlsx_report.get_sheet("Errors")
			errors_sheet << errors_headers
		end

		matches_sheet_row_count = 0
		matches_sheet_count = 1
		semaphore = Mutex.new

		# Configure callback on RegexScanner to capture when it reports
		# errors and record them to relevant places (reports, log area)
		scanner.whenErrorOccurs do |scan_error|
			error_count += 1

			title = ""
			expression = ""
			location = ""
			guid = scan_error.getItem.getGuid
			message = scan_error.getException.getMessage

			# Depending on where error occurred pattern info may not be available
			# so we should handle this to prevent error reporting from generating errors
			if !scan_error.getPatternInfo.nil?
				title = scan_error.getPatternInfo.getTitle
				expression = scan_error.getPatternInfo.getExpression
			end

			# Depending on where error occurred location may not be available
			# so we should handle this to prevent error reporting from generating errors
			if !scan_error.getLocation.nil?
				location = scan_error.getLocation
			end

			# Generate array representing error report row
			errors_values = [
				title,
				expression,
				location,
				guid,
				message,
			]

			# Record error to CSV
			if generate_report_csv
				errors_csv << errors_values
			end

			# Record error to Excel
			if generate_report_xlsx
				errors_sheet << errors_values
			end

			# Record error to log area
			pd.logMessage("Error while scanning:")
			pd.logMessage("\tTitle: #{title}")
			pd.logMessage("\tExpression: #{expression}")
			pd.logMessage("\tLocation: #{location}")
			pd.logMessage("\tGUID: #{expression}")
			pd.logMessage("\tMessage: #{message}")
		end

		# Actual scanning of items using a callback to record matches
		work_proc = Proc.new do |item_match_collection|
			# Abort scanning if user requested abort through the progress dialog
			if pd.abortWasRequested
				scanner.abortScan
			end

			pd.setSubStatus("")
			matched_item_count += 1
			item = item_match_collection.getItem
			# Create hash inside hash where inner hash has empty array as initial value
			# Used to group by hash[location][title] = [item1,item2,etc]
			location_title_grouped = Hash.new{|h,k| h[k] = Hash.new{|h2,k2| h2[k2] = [] } }

			# Record data for the matches against an item
			item_match_collection.getMatches.each do |match|

				matched_value_count += 1

				# We will be using these to resolve placeholders user may have
				# specified in the settings dialog for tags, custom metdata
				placeholders = {
					"location" => match.getLocation,
					"title" => match.getPatternInfo.getTitle,
				}

				# Record tag info if we are applying tags
				if apply_tags
					tag = $su.getFormatUtility.resolvePlaceholders(tag_template,placeholders)
					semaphore.synchronize {
						tag_grouped[tag] << item
					}
				end

				# Record custom metadata info if we are applying custom metadata
				if apply_custom_metadata
					location_title_grouped[match.location][match.getPatternInfo.getTitle] << match.getValue
				end

				# Get match strings (value matches and possible match context string)
				match_value = match.getValue
				match_value_context = match.getValueContext

				# Need to handle really long matches and truncate them if they
				# exceed Excel limits
				if match_value.size > 32_000 || match_value_context.size > 32_000
					pd.logMessage("GUID: #{item.getGuid}")
					pd.logMessage("Regex Title: #{match.getPatternInfo.getTitle}")
					pd.logMessage("Match Location: match.getLocation")

					# Does the matched value exceed the limit?
					if match_value.size > 32_000
						match_value = match_value[0..32_000]
						pd.logMessage("- Match value exceeds Excel 32K character cell limit and will be truncated in report")
					end

					# Does the context value exceed the limit?
					if match_value_context.size > 32_000
						pd.logMessage("- Match value context exceeds Excel 32K character cell limit and will be truncated in report")
						match_value_context = match_value_context[0..32_000]
					end
				end
				
				# Generate array that represents a row in the report
				report_row = [
					item.getGuid,
					item.getLocalisedName,
					item.getType.getKind.getName,
					match.getPatternInfo.getTitle,
					match.getLocation,
					match_value,
					match_value_context,
					match.getMatchStart,
					match.getMatchEnd,
				]

				if include_item_path
					report_row << item.getLocalisedPathNames.join("/")
				end

				if include_physical_path
					report_row << $su.getSuperItemUtility.getPhysicalAncestorPath(item)
				end

				# Record to CSV
				if generate_report_csv
					matches_csv << report_row
				end

				# Record to XLSX, also check if we overflow the approximately 1 million rows
				# per Excel sheet and start in on a new sheet if needed
				if generate_report_xlsx
					semaphore.synchronize {
						matches_sheet << report_row
						matches_sheet_row_count += 1
						if matches_sheet_row_count >= 1_000_000
							matches_sheet_row_count = 0
							matches_sheet_count += 1
							matches_sheet = xlsx_report.get_sheet("Matches #{matches_sheet_count}")
							matches_sheet << matches_headers
						end
					}
				end
			end

			# Apply custom metdata if settings specified to do so
			if apply_custom_metadata
				semaphore.synchronize {
					cm = item.getCustomMetadata
					location_title_grouped.each do |location,title_grouped|
						title_grouped.each do |title,values|
							placeholders = {
								"location" =>location,
								"title" => title,
							}
							field_name = $su.getFormatUtility.resolvePlaceholders(field_name_template,placeholders)
							cm[field_name] = values.uniq.join("; ")
						end
					end
				}
			end
		end

		if $scan_in_parallel == true
			scanner.scanItemsParallel(items,work_proc,$scan_concurrency)
		else
			scanner.scanItems(items,work_proc)
		end

		# Scanning completed so report one final status to the log area
		elapsed_seconds = (Time.now - start_time).to_i
		elapsed = $su.getFormatUtility.secondsToElapsedString(elapsed_seconds)
		message = "#{elapsed} Scanned #{items.size}/#{items.size}, Matched Items: #{matched_item_count}, Matches: #{matched_value_count}, Errors: #{error_count}"
		pd.setMainStatusAndLogIt(message)

		# Finalize CSV reports by closing them
		if generate_report_csv
			matches_csv.close
			errors_csv.close
		end

		# Finalize Excel report by saving it
		if generate_report_xlsx
			xlsx_report.save(report_xlsx_file)
		end

		# Apply tags based on information collected earlier, if settings specified to apply tags
		if apply_tags
			pd.setSubStatus("Applying tags...")
			annotater = $utilities.getBulkAnnotater
			tag_grouped.each do |tag,items|
				pd.setSubStatusAndLogIt("Applying tag '#{tag}' to #{items.size} items")
				annotater.addTag(tag, items)
			end
		end

		# Finalize the progress dialog status based on whether the user aborted early or
		# we completed as expected
		if pd.abortWasRequested
			pd.setMainStatusAndLogIt("User Aborted")
		else
			pd.setCompleted
		end

		# For user convenience report to the logging area the final total elapsed time
		elapsed_seconds = (Time.now - start_time).to_i
		elapsed = $su.getFormatUtility.secondsToElapsedString(elapsed_seconds)
		pd.logMessage("Total Time: #{elapsed}")

		# If we closed all the tabs we should open a new one for the user
		if apply_custom_metadata
			$window.openTab("workbench",{:search=>""})
		end
	end
end