# convert-spreadsheet
The program converts any spreadsheet to .xlsx Strict conformance and to meet [archival data quality specifications](https://github.com/Asbjoedt/CLISC/wiki/Archival-Data-Quality). It can be used in simple archival workflows. It receives any filepath, if it is a spreadsheet file format, it will convert, rename to 1.xlsx and finally delete the original file.

* For more information, see repository **[CLISC](https://github.com/Asbjoedt/CLISC)**

## Dependencies
:warning: **[Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)**

Excel can be used as conversion tool.

:warning: **[LibreOffice](https://libreoffice.org)**

LibreOffice can be used as conversion tool.

## How to use
Download the executable version [here](https://github.com/Asbjoedt/convert-spreadsheet/releases). There's no need to install. In your terminal change directory to the folder where convert-spreadsheet.exe is. Then, to execute the program input:
```
.\convert-spreadsheet.exe --inputfilepath="[filepath]" --outputfileformat="[extension]"
```

**Parameters**

Required
```
--inputfilepath="[filepath]" // path to the file you want to convert
--outputfileformat="[extension]" // your output file format
```
Optional
```
--outputfolder="[folder]" // your custom folder for output file, i.e. "C:\Users\%USERNAME%\Desktop"
--libreoffice // if you want to use LibreOffice as conversion tool instead of Excel
--policy // if you want to convert data to comply with archiving requirements
--delete // if original file should be deleted
--rename="[filename]" // your custom filename, i.e "1".
```

**Exit codes**

The program writes information to the console and it also returns an exit code to integrate in workflows.
```
0 = File failed conversion
1 = File completed conversion
```

## Packages and software
The following packages and software are used under license.
* [LibreOffice](https://www.libreoffice.org/), Mozilla Public License v2.0
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), Copyright (c) Microsoft Corporation
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
