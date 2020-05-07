/**
 * LibCompileVBA
 * Provides a class to add files to a Excel macro workbook.
 * 
 * Consumes an xlsx file and injects macros into it.
 */
using System;
using OpenMcdf;
using System.Text;
using System.Collections.Generic;
using Kavod.Vba.Compression;
using System.Linq;
using NDesk.Options;
using System.Net;
using System.Threading;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Collections;

public enum VBAProjectType {
  Excel,
  Word,
  PowerPoint
}


public class VBA {
  private string Path;
  private string unzipTempPath;
  private string sOLEFileName;
  private Boolean isOpenXML;
  private CompoundFile cf;
  private CFStorage commonStorage;
  
  private byte[] vbaProjectStream;
	private byte[] dirStream;
	private byte[] projectStream;
	private byte[] projectwmStream;

  List<ModuleInformation> vbaModules;

  public static VBA create(VBAProjectType type){
    //Copy template
    string templatePath = GetTemplateCopy(type);

    //Create and return VBA object
    return new VBA(templatePath);
  }
  
  /**
   * Create VBA Project from file
   *
   * @param sPath path of macro enabled Excel, Powerpoint or Word file.
   */
  public VBA(string sPath){
    
    

    //Get unzip location
    unzipTempPath = CreateUniqueTempDirectory();
    
    /* EvilClippy brute forces the file open, It does this in a fairly unclean manner. We will assume the file,
     * is named correctly (extension wise). This, although requiring more boiler plate will look a lot cleaner.
     */

    // Get extension of file:
    string ext = Path.GetExtension(sPath);
    
    //Parse extension and create sOLEFileName
    switch(ext){
      case ".xlsm":
      case ".xlam":
        isOpenXML = true;
        sOLEFilename = Path.Combine(unzipTempPath,"xl","vbaProject.bin");
        break;
      case ".docm":
        isOpenXML = true;
        sOLEFilename = Path.Combine(unzipTempPath,"word","vbaProject.bin");
        break;
      case ".pptm":
        isOpenXML = true;
        sOLEFilename = Path.Combine(unzipTempPath,"ppt","vbaProject.bin"); //untested
        break;
      case ".xls":
      case ".doc":
      case ".ppt":
        //Copy path to sOLEFileName to prevent overwriting (we'll overrite this file directly)
        sOLEFileName = Path.Combine(unzipTempPath,Path.GetFileName(sPath));
        File.Copy(sPath,sOLEFileName);
        break;    
      default:
        Console.WriteLine("ERROR: Could not open file " + filename);
			  Console.WriteLine("Please make sure this file exists, has a valid extension and is of a valid type.");
			  Console.WriteLine();
        break;
    }

    //Unzip to unzipTemoPath if isOpenXML. Otherwise create compound file.
    if(isOpenXML){
      ZipFile.ExtractToDirectory(sPath, unzipTempPath);
    } 

    //Create Compound file from VBProject.bin or xls,doc,ppt file:
    cf = new CompoundFile(sOLEFilename, CFSUpdateMode.Update, 0);
    
    // Read relevant streams
    switch(ext){
      case ".xlsm":
      case ".xlam":
      case ".docm":
      case ".pptm":
        commonStorage = cf.RootStorage;
        break;

      case ".doc":
        commonStorage = cf.RootStorage.GetStorage("Macros");   //Note you can also use `cf.RootStorage.TryGetStorage("Macros")` which returns null if not found.
        break;
      
      case ".ppt": //untested - don't know if this is the case for ppts
      case ".xls":
        commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR");
        break;

    }

		vbaProjectStream = commonStorage.GetStorage("VBA").GetStream("_VBA_PROJECT").GetData();
		projectStream = commonStorage.GetStream("project").GetData();
		projectwmStream = commonStorage.GetStream("projectwm").GetData();
		dirStream = Decompress(commonStorage.GetStorage("VBA").GetStream("dir").GetData());

		// Read project streams as string
		string projectStreamString = System.Text.Encoding.UTF8.GetString(projectStream);
		string projectwmStreamString = System.Text.Encoding.UTF8.GetString(projectwmStream);

		// Find all VBA modules in current file
		vbaModules = ParseModulesFromDirStream(dirStream);

		// Write streams to debug log (if verbosity enabled)
		DebugLog("Hex dump of original _VBA_PROJECT stream:\n" + Utils.HexDump(vbaProjectStream));
		DebugLog("Hex dump of original dir stream:\n" + Utils.HexDump(dirStream));
		DebugLog("Hex dump of original project stream:\n" + Utils.HexDump(projectStream));


	
  }

  /**
   * Add a file to the VBA project. Files added mostly based off of extension. So .wkb, .sht, .cls, .bas, .frm
   * 
   *
   *
   */
  public Boolean addFile(string sFile){

  }

  public boolean save(string sPath = ""){
    //Save to file
  }


  //HELPERS

  private static string CreateUniqueTempDirectory()
	{
		var uniqueTempDir = Path.GetFullPath(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));
		Directory.CreateDirectory(uniqueTempDir);
		return uniqueTempDir;
	}

  public class ModuleInformation
	{
		public string moduleName; // Name of VBA module stream
		public UInt32 textOffset; // Offset of VBA source code in VBA module stream
	}

  private static List<ModuleInformation> ParseModulesFromDirStream(byte[] dirStream)
	{
		// 2.3.4.2 dir Stream: Version Independent Project Information
		// https://msdn.microsoft.com/en-us/library/dd906362(v=office.12).aspx
		// Dir stream is ALWAYS in little endian

		List<ModuleInformation> modules = new List<ModuleInformation>();

		int offset = 0;
		UInt16 tag;
		UInt32 wLength;
		ModuleInformation currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };

		while (offset < dirStream.Length)
		{
			tag = GetWord(dirStream, offset);
			wLength = GetDoubleWord(dirStream, offset + 2);

			// The following idiocy is because Microsoft can't stick to their own format specification - taken from Pcodedmp
			if (tag == 9)
				wLength = 6;
			else if (tag == 3)
				wLength = 2;

			switch (tag)
			{
				case 26: // 2.3.4.2.3.2.3 MODULESTREAMNAME Record
					currentModule.moduleName = System.Text.Encoding.UTF8.GetString(dirStream, (int)offset + 6, (int)wLength);
					break;
				case 49: // 2.3.4.2.3.2.5 MODULEOFFSET Record
					currentModule.textOffset = GetDoubleWord(dirStream, offset + 6);
					modules.Add(currentModule);
					currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };
					break;
			}

			offset += 6;
			offset += (int)wLength;
		}

		return modules;
	}
}