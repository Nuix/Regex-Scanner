Regex Scanner
=============

![This script was last tested in Nuix 7.8](https://img.shields.io/badge/Script%20Tested%20in%20Nuix-7.8-green.svg)

View the GitHub project [here](https://github.com/Nuix/Regex-Scanner) or download the latest release [here](https://github.com/Nuix/Regex-Scanner/releases).

# Overview

**Written By:** Jason Wells

This script allows you report on when items are found to contain matches to regular expressions in their content text and/or the text equivalent (once converted to text) of their properties.  The script supports generating reports, tagging and custom metadata.

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

## Settings

**Tip:** Settings can be saved/loaded through the file menu.

### Scan Settings Tab

| Setting | Description |
|---------|-------------|
| **Skip Excluded Items** | When checked, excluded items are filtered from the input set of items.  If items are selected when the script is ran, the input set of items will be those selected.  If no items are selected when the script is ran, the input set of items will be all items in the case. |
| **Expressions are Case Sensitive** | When checked, matches are made in a case sensitive manner (capitalization must match that in a given expression). |
| **Capture Match Value Context** | When checked, reports will contain contextual text for each match (some of the text surround the match). |
| **Context Size in Characters** | Determines how many characters before and after a match are included in match context information when **Capture Match Value Context** is checked. |
| **Scan Item Content** | When checked, the content text of items will be checked for matches to the provided expressions. |
| **Scan Item Properties** | When checked, metadata property values of items will be checked for matches to the provided expressions.  Metadata property values are converted to strings before matching using the corresponding metadata profile fields. |
| **Properties to Scan** | When **Scan Item Properties** is checked, this table determines which metadata properties will be scanned.  Uncheck properties if you do not wish to get matches from particular metadata properties. |

### Reporting Tab

| Setting | Description |
|---------|-------------|
| **Apply Tags** | When checked, each item which has a match will have one or more tags applied, based on the provided tag template.  In the template the placeholder `{location}` will be replaced with the location the match was made.  This will be replaced with `Content`.  If a match is made against a metadata property, this will be replaced with the name of the metadata property in which the match was made.  The placeholder `{title}` will be replaced with the title value associated to the matching regular expression. |
| **Apply Custom Metadata** | When checked, each item which has a match will have one or more custom metadata fields applied.  The name of the field applied will be based upon the template provided.  In the template the placeholder `{location}` will be replaced with the location the match was made.  This will be replaced with `Content`.  If a match is made against a metadata property, this will be replaced with the name of the metadata property in which the match was made.  The placeholder `{title}` will be replaced with the title value associated to the matching regular expression.  The value of the custom metadata field applied will be a semicolon space (`; `) delimited list of the actual matches made. |
| **Generate CSV Report** | When checked, a series of CSVs will be generated in the directory specified reporting information about what items had a match, which expression matched, where each match occurred, what the text of the actual match was, any errors that occurred and context text if **Capture Match Value Context** is checked on the scan settings tab. |
| **Generate XLSX Report** | When checked, a Excel XLSX file will be generated in the directory specified reporting information about what items had a match, which expression matched, where each match occurred, what the text of the actual match was, any errors that occurred and context text if **Capture Match Value Context** is checked on the scan settings tab. |
| **Include Item Path** | When checked, report CSV/XLSX will contain a column with the item path of each matched item, exameple: `Evidence 1/BobSmith.pst/Inbox/RE: lunch today?`. |
| **Include Physical Ancestor Path** | When checked, report CSV/XLSX will contain a column with the file system path of the physical file ancestor of each matched item (if there is one).  So for example, scanning the contents of a PST file `C:\SourceData\BobSmith.pst`, this would then contain the PST file path for each email matched within that PST. |

## Cloning this Repository

This script relies on code from [Nx](https://github.com/Nuix/Nx) to present a settings dialog and progress dialog.  This JAR file is not included in the repository (although it is included in release downloads).  If you clone this repository, you will also want to obtain a copy of Nx.jar by either:
1. Building it from [the source](https://github.com/Nuix/Nx)
2. Downloading an already built JAR file from the [Nx releases](https://github.com/Nuix/Nx/releases)

Once you have a copy of Nx.jar, make sure to include it in the same directory as the scripts.

This script also relies on code from [SuperUtilities](https://github.com/Nuix/SuperUtilities), which contains the code to do the actual work.  This JAR file is not included in the repository (although it is included in release downloads).  If you clone this repository, you will also want to obtain a copy of SuperUtilities.jar by either:
1. Building it from [the source](https://github.com/Nuix/SuperUtilities)
2. Downloading an already built JAR file from the [Nx releases](https://github.com/Nuix/SuperUtilities/releases)

Once you also have a copy of SuperUtilities.jar, make sure to include it in the same directory as the scripts.

# License

```
Copyright 2018 Nuix

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
