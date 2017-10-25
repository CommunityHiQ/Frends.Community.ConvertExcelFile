- [FRENDS.Community.Excel.ConvertExcelFile](#FRENDS.Community.Excel.ConvertExcelFile)
   - [Installing](#installing)
   - [Building](#building)
   - [Contributing](#contributing)
   - [Documentation](#documentation)
      - [ConvertExcelFile](#convertExcelFile)
		 - [Input](#input)
		 - [Options](#options)
		 - [Result](#result)
   - [License](#license)
       
# FRENDS.Community.Excel.ConvertExcelFile
This repository contais FRENDS4 Community Excel Tasks

## Installing
You can install the task via FRENDS UI Task view or you can find the nuget package from the following nuget feed
https://www.myget.org/F/frends/api/v2

## Building
Ensure that you have https://www.myget.org/F/frends/api/v2 added to your nuget feeds

Clone a copy of the repo

git clone https://github.com/FrendsPlatform/Frends.Web.git

Restore dependencies

nuget restore frends.web`

Rebuild the project

Run Tests with nunit3. Tests can be found under

Frends.Web.Tests\bin\Release\Frends.Web.Tests.dll

Create a nuget package

nuget pack nuspec/Frends.Web.nuspec`

## Contributing
When contributing to this repository, please first discuss the change you wish to make via issue, email, or any other method with the owners of this repository before making a change.

1. Fork the repo on GitHub
2. Clone the project to your own machine
3. Commit changes to your own branch
4. Push your work back up to your fork
5. Submit a Pull request so that we can review your changes

NOTE: Be sure to merge the latest from "upstream" before making a pull request!

## Documentation

### ConvertExcelFile

Reads Excel file and converts it to XML or CSV according to the task input parameters.

#### Input
| Property  | Type  | Description |Example|
|-----------|-------|-------------|-------|
| Path  | string | Path of the Excel file to be read. | C:\temp\ExcelFile.xlsx|

#### Options
| Property  | Type  | Description |
|-----------|-------|-------------|
| ReadOnlyWorkSheetWithName  | string | Excel work sheet name to be read. If empty, all work sheets are read. | 
| outputFileType| Enum(typeOf(OutputFileType) | Choose format output string as XML or CSV. |
| CsvSeparator| string | Csv Separator |
| UseNumbersAsColumnHeaders| bool | If set to true, outputs column headers as numbers instead of letters. |

#### Result
| Property  | Type  | Description |
|-----------|-------|-------------|
| result| string  | Returns result object (string). |

## License
This project is licensed under the MIT License - see the LICENSE file for details