# convert-spreadsheets
The program converts any spreadsheet to .xlsx Strict conformance and to meet archival data quality specifications. It can be used in simple archival workflows. It receives any filepath, if it is a spreadsheet file format, it will convert, rename to 1.xlsx and finally delete the original file.

## Dependencies
:warning: **[Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)**

Excel is used in the background for conversion.

## How to use
Download the executable version [here](https://github.com/Asbjoedt/convert-spreadsheets/releases). There's no need to install. In your terminal change directory to the folder where convert-spreadsheets.exe is. Then, to execute the program input:
```
.\convert-spreadsheets.exe "[filepath]"
```

## Packages and software
The following packages and software are used under license.
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), Copyright (c) Microsoft Corporation
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
