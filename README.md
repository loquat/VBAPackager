
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

## Long term goals of this project / integrations with this project

This project was designed as a stepping stone for making VBA projects easier to write and build in any IDE of your choosing. 3 key issues stand in the way of moving VBA projects out of the IDE and into custom development environments:

1. A VBA code parser and language server - this is possibly going to be created by [rubberduck](https://github.com/rubberduck-vba/Rubberduck/issues/5176).
2. A VBA compiler / file reader - which this project aims to become.
3. A VBA interpreter and/or http debug server - for running and debugging code.

On its own this only makes an external editor as good as the VBE, but we can do much better than that if we: 

4. A VBA Babel-like compiler - this would allow us to parse custom code and translate it into official VBA code. Could have great benefits like line numbers on error, method property modification in code, optimised in-line functions/lambda expressions, inline classes, and much more...

## Authors

* openVBA Compiler - Sancarn 
* EvilClippy - Stan Hegt ([@StanHacked](https://twitter.com/StanHacked)) / [Outflank](https://www.outflank.nl)
* EvilClippy - Carrie Roberts ([@OrOneEqualsOne](https://twitter.com/OrOneEqualsOne) / Walmart).
* EvilClippy - Nick Landers ([@monoxgas](https://twitter.com/monoxgas) / Silent Break Security) for pointing me towards OpenMCDF.