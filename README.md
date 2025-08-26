# convert-spreadsheet
The program converts any input spreadsheet file format to any output spreadsheet file format using Excel or LibreOffice in the background. If output file format is xlsx, the program can convert to meet [archival data quality specifications](https://github.com/Asbjoedt/CLISC/wiki/Archival-Data-Quality). The program is intended for use in archival workflows.

* For more information, see repository **[CLISC](https://github.com/Asbjoedt/CLISC)**

## Dependencies
:warning: **[Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)**

Excel can be used as conversion tool.

:warning: **[LibreOffice](https://libreoffice.org)**

LibreOffice can be used as conversion tool.

:warning: **[ODS-ArchivalRequirements](https://github.com/Asbjoedt/ODS-ArchivalRequirements)**

ODS-ArchivalRequirements are used for policy compliance for ODS spreadsheets.

## How to use
Download the executable version [here](https://github.com/Asbjoedt/convert-spreadsheet/releases). There's no need to install. In your terminal change directory to the folder where convert-spreadsheet.exe is. Then, to execute the program input:
```
.\convert-spreadsheet.exe --inputfilepath="[filepath]" --outputfileformat="[extension]"
```

**Parameters**

Required
```
--inputfilepath="[filepath]" // path to the file you want to convert
--outputfileformat="[extension]" // your output file format e.g. "ods", "xlsx", "xlsx-strict", "csv"
```
Optional
```
--outputfolder="[folder]" // your custom folder for output file, i.e. "C:\Users\%USERNAME%\Desktop"
--libreoffice // if you want to use LibreOffice as conversion tool instead of Excel
--policy // if you want to convert data to comply with regular archiving requirements
--policy-strict // if you want to convert data to comply with strict archiving requirements
--delete // if original file should be deleted
--rename="[filename]" // your custom filename e.g. "1".
```

**Exit codes**

The program writes information to the console and it also returns an exit code to integrate in workflows.
```
0 = File failed conversion
1 = File completed conversion
```

## Packages and software
The following packages and software are used under license.
* [.Net 9](https://dotnet.microsoft.com/en-us/download/dotnet/9.0), copyright (c) Microsoft Corporation
* [LibreOffice](https://www.libreoffice.org/), Mozilla Public License v2.0
* [Magick.Net](https://github.com/dlemstra/Magick.NET), Apache-2.0 license, copyright (c) Dirk Lemstra
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), Copyright (c) Microsoft Corporation
* [ODS-ArchivalRequirements](https://github.com/Asbjoedt/ODS-ArchivalRequirements), MIT license, copyright (c) Asbjørn Skødt
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
