
# Cross-platform Open-Source VBA Compiler

This work-in-progress tool is designed to help developers create Excel macro-workbooks from source VBA `.bas`, `.cls` and `.frm` files. This tool does not require Excel to be installed on the system running it, and operates fully outside of the VBE.

This project is based off the amazing work by [EvilClippy](https://github.com/outflanknl/EvilClippy).

## Details

Excel macro workbooks store VBA code in 2 parts. `Optimised PCode` and `Compressed Source`. If the `Optimised PCode` part is missing (or incompatable with the running version of Excel) then the `Compressed Source` is used to create a new set of `Optimised PCode`. This tool intends to erase the Optimised PCode and Inject new compressed source code into the Excel workbook.

## Proposed VBA Project structure:

```
|- Libs                     //Hidden (background) modules
|  |- STD
|  |  |- IEnumVariant.bas
|  |  |- stdArray.cls
|  |  |- stdDictionary.cls
|  |  |- stdCallback.cls
|  |  |- ...
|  |
|  |- JSONBag
|  |  |- JSONBag.cls
|  |  |- License.txt
|  |
|  |- VBA-WEB
|  |  |- Request.cls
|  |  |- UtcConverter.bas
|
|- Src                      //Project files (not hidden) modules
|  |- ThisWorkbook.wkb
|  |- Dashboard.sht
|  |- ADMIN.sht
|  |- Settings.sht
|  |- HUDMain.bas
|  |- HUD
|  |  |- HUDCore.cls
|  |  |- HUDUI.cls
|  |  |- HUDConnect.cls
|  |  |- HUDGraph.cls
|  |  |- HUDGraphNode.cls
|  |  |- HUDGraphEdge.cls
|  |  |- HUDGDI.bas
|
|- Input.xlsm
|- Output.xlsm
|- Setup.json
```

with Setup format:

```json
{
  "hideLibs":true,
  "linker":[
    {"codename":"Dashboard","file":"Dashboard.sht"},      //Using sheet codename
    {"sheetIndex": 2,"file":"ADMIN.sht"},                 //Using sheet index
    {"sheetName":"SETTINGS", "file":"Settings.sht"}       //Using sheet name
  ]
}

```

## Internal VBACompiler structure:

```
root
|- main.cs - Command line application which performs the above task using the VBA.cs library.
|- VBA.cs  - Might rename to VBAProject.cs - Provides a base API for working with VBA workbooks. Can open existing workbooks or create from a template. Add modules to the project, remove them, hide them etc.
|- libs
|  |- Kabod.Vba.Compression.cs - Provides implementation of VBA compression routine (so we can decompress and compress vba source code out of VB project.)
|  |- options.cs - Library for working with commandline options?
|  |- utils.cs - Small helpful functions

```


## Authors

* openVBA Compiler - Sancarn 
* EvilClippy - Stan Hegt ([@StanHacked](https://twitter.com/StanHacked)) / [Outflank](https://www.outflank.nl)
* EvilClippy - Carrie Roberts ([@OrOneEqualsOne](https://twitter.com/OrOneEqualsOne) / Walmart).
* EvilClippy - Nick Landers ([@monoxgas](https://twitter.com/monoxgas) / Silent Break Security) for pointing me towards OpenMCDF.