
# Cross-platform Open-Source VBA Compiler

This work-in-progress tool is designed to help developers create Excel macro-workbooks from source VBA `.bas`, `.cls` and `.frm` files. This tool does not require Excel to be installed on the system running it, and operates fully outside of the VBE.

This project is based off the amazing work by [EvilClippy](https://github.com/outflanknl/EvilClippy).

## Details

Excel macro workbooks store VBA code in 2 parts. `Optimised PCode` and `Compressed Source`. If the `Optimised PCode` part is missing (or incompatable with the running version of Excel) then the `Compressed Source` is used to create a new set of `Optimised PCode`. This tool intends to erase the Optimised PCode and Inject new compressed source code into the Excel workbook.


## Authors

* openVBA Compiler - Sancarn 
* EvilClippy - Stan Hegt ([@StanHacked](https://twitter.com/StanHacked)) / [Outflank](https://www.outflank.nl)
* EvilClippy - Carrie Roberts ([@OrOneEqualsOne](https://twitter.com/OrOneEqualsOne) / Walmart).
* EvilClippy - Nick Landers ([@monoxgas](https://twitter.com/monoxgas) / Silent Break Security) for pointing me towards OpenMCDF.