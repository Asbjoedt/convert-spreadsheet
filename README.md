# convert-spreadsheets

The program converts any spreadsheet to .xlsx Strict conformance and to meet archival data quality specifications. It can be used in simple archival workflows. It receives any filepath, if it is a spreadsheet file format, it will convert, rename to 1.xlsx and finally delete the original file.

## Dependencies

:warning: **[Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)**
* If you want to convert legacy Excel and convert .xlsx conformance from Transitional to Strict (mandatory)

:warning: **[LibreOffice](https://www.libreoffice.org/)**
* If you want to convert OpenDocument spreadsheets
* You need to install program in its default directory, or create environment variable "LibreOffice" with path to your installation

## How to use
Download the executable version [here](https://github.com/Asbjoedt/convert-spreadsheets/releases). There's no need to install. In your terminal change directory to the folder where convert-spreadsheets.exe is. Then, to execute the program input:
```
.\convert-spreadsheets.exe "[filepath]"
```

## Packages and software

The following packages and software are used under license in CLISC. [Read more](https://github.com/Asbjoedt/CLISC/wiki/Dependencies).

* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), Copyright (c) Microsoft Corporation
* [LibreOffice](https://www.libreoffice.org/), Mozilla Public License v2.0
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
