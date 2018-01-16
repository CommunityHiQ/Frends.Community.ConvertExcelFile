- [Frends.Community.ConvertExcelFile](#Frends.Community.ConvertExcelFile)
   - [Installing](#installing)
   - [Building](#building)
   - [Contributing](#contributing)
   - [Documentation](#documentation)
      - [ConvertExcelFile](#convertExcelFile)
		 - [Input](#input)
		 - [Options](#options)
		 - [Result](#result)
   - [License](#license)
       
# Frends.Community.ConvertExcelFile
This repository contais FRENDS4 Community Excel Tasks

# Installing
You can install the task via Frends UI Task view or you can find the nuget package from the following nuget feed
https://www.myget.org/F/frends/api/v3/index.json

Tasks
=====

## ConvertExcelFile

Reads Excel file and converts it to XML, CSV and JSON according to the task input parameters.

### Task Properties

#### Input
| Property  | Type  | Description |Example|
|-----------|-------|-------------|-------|
| Path  | string | Path of the Excel file to be read. | C:\temp\ExcelFile.xlsx|

#### Options
| Property  | Type  | Description |Example|
|-----------|-------|-------------|-------|
| ReadOnlyWorkSheetWithName  | string | Excel work sheet name to be read. If empty, all work sheets are read. |Sheet1| 
| CsvSeparator| string | Csv Separator | ; |
| UseNumbersAsColumnHeaders| bool | If set to true, numbers will be used as column headers instead of letters (A = 1, B = 2...) | true |
| ThrowErrorOnfailure| bool | Throws an exception if conversion fails. |  true |

### Result
| Property  | Type  | Description |Example|
|-----------|-------|-------------|-------|
| ResultData | DataSet  | Conversion result as a DataSet| |
| Success | bool | Task execution result. | true |
| Message | string | Exception message | "File not found"|
| ToXml() | string| Converts result to XML| XML-string|
| ToCsv() | string | Converts result to CSV| CSV-string |
| ToJToken() | JToken |  Converts result to Json||

# License
This project is licensed under the MIT License - see the LICENSE file for details

# Building
Ensure that you have https://www.myget.org/F/frends/api/v3/index.json added to your nuget feeds

Clone a copy of the repo

git clone https://github.com/CommunityHiQ/Frends.Community.ConvertExcelFile.git

Restore dependencies

nuget restore Frends.Community.ConvertExcelFile

Rebuild the project

Run Tests with nunit3. Tests can be found under

Frends.Community.ConvertExcelFileTests\bin\Release\Frends.Community.ConvertExcelFileTests.dll

Create a nuget package

`nuget pack nuspec/Frends.Community.ConvertExcelFile.nuspec`

# Contributing
When contributing to this repository, please first discuss the change you wish to make via issue, email, or any other method with the owners of this repository before making a change.

1. Fork the repo on GitHub
2. Clone the project to your own machine
3. Commit changes to your own branch
4. Push your work back up to your fork
5. Submit a Pull request so that we can review your changes

NOTE: Be sure to merge the latest from "upstream" before making a pull request!

# Change Log

| Version             | Changes                 |
| ---------------------| ---------------------|
| 1.4.0 | Initial Frends.Community version of ExcelTask converted from old Frends.Common code base |
| 1.4.1 | Minor updates and changes to documentation and naming conventions |