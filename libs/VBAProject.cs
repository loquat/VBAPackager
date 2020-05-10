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

namespace VBAPackager {
  public class VBAProject {
    private string pPath;
    private string unzipTempPath;
    public string sOLEFileName;
    private Boolean isOpenXML;
    
    private byte[] vbaProjectStream;
    private byte[] dirStream;
    private byte[] projectStream;
    private byte[] projectwmStream;

    private List<ModuleInformation> vbaModulesEx;
    public List<VBAModule> vbaModules;
    private CompoundFile cf;
    private CFStorage commonStorage;

    /**
    * Create VBA Project from file
    *
    * @param sPath path of macro enabled Excel, Powerpoint or Word file.
    */
    public VBAProject(string sPath){
      //Initialize pPath
      pPath = sPath;
      
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
          sOLEFileName = Path.Combine(unzipTempPath,"xl","vbaProject.bin");
          break;
        case ".docm":
          isOpenXML = true;
          sOLEFileName = Path.Combine(unzipTempPath,"word","vbaProject.bin");
          break;
        case ".pptm":
          isOpenXML = true;
          sOLEFileName = Path.Combine(unzipTempPath,"ppt","vbaProject.bin"); //untested
          break;
        case ".xls":
        case ".doc":
        case ".ppt":
          //Copy path to sOLEFileName to prevent overwriting (we'll overrite this file directly)
          sOLEFileName = Path.Combine(unzipTempPath,Path.GetFileName(sPath));
          File.Copy(sPath,sOLEFileName);
          break;    
        default:
          Console.WriteLine("ERROR: Could not open file " + sPath);
          Console.WriteLine("Please make sure this file exists, has a valid extension and is of a valid type.");
          Console.WriteLine();
          break;
      }

      //Unzip to unzipTemoPath if isOpenXML. Otherwise create compound file.
      if(isOpenXML){
        ZipFile.ExtractToDirectory(sPath, unzipTempPath);
      } 

      //Create Compound file from VBProject.bin or xls,doc,ppt file:
      cf = new CompoundFile(sOLEFileName, CFSUpdateMode.Update, 0);
      
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
      
      //Contains VBA Project properties, including module names. Remove module names from this stream to hide them.
      //PROJECT stream:   https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/cc848a02-6f87-49a4-ad93-6edb3103f593
      //More information: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/d88cb9d8-a475-423d-b370-cc0caaf78628
      projectStream = commonStorage.GetStream("project").GetData();
      
      //PROJECTwm stream contains all names of all Modules.
      //PROJECTwm stream:  https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/514acc65-ea7b-4813-aaf7-fabb1bca0ba2
      //More Information:  https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/c458f2e6-f2cc-4c2d-96c7-91a3e63f2fe1
      projectwmStream = commonStorage.GetStream("PROJECTwm").GetData();                              
      
      //Contains names of visible modules and module source code text (compressed)
      //VBA Storage:     https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/b693e0ba-489f-4ac8-ac9d-6387fb5779bb
      //VBA StorageInfo: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/170f52a0-4cd6-4729-b51a-d08155cbd213
      //Dir Stream:      https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/672d276c-d83c-4452-993b-ca9eca3d8917
      dirStream = VbaCompression.Decompress(commonStorage.GetStorage("VBA").GetStream("dir").GetData());
      
      // Read project streams as string
      string projectStreamString = System.Text.Encoding.UTF8.GetString(projectStream);
      string projectwmStreamString = System.Text.Encoding.UTF8.GetString(projectwmStream);
      Console.WriteLine(System.Text.Encoding.UTF8.GetString(projectStream));
      Console.WriteLine(System.Text.Encoding.UTF8.GetString(projectwmStream));
      


      // Find all VBA modules in current file
      vbaModulesEx = ParseModulesFromDirStream(dirStream);

      // Write streams to debug log (if verbosity enabled)
      Console.WriteLine("Hex dump of original _VBA_PROJECT stream:\n" + Utils.HexDump(vbaProjectStream));
      Console.WriteLine("Hex dump of original dir stream:\n" + Utils.HexDump(dirStream));


    
    }

    //Finalizer i.e. Called when objects of this type are destroyed
    //Cleanup temporary files.
    ~VBAProject(){
      if(Directory.Exists(unzipTempPath)){
        Directory.Delete(unzipTempPath);
      }
    }

    // public static VBAProject create(VBAProjectType type){
    //   //Copy template
    //   string templatePath = GetTemplateCopy(type);

    //   //Create and return VBA object
    //   return new VBAProject(templatePath);
    // }

    /**
    * Add a file to the VBA project. Files added mostly based off of extension. So .wkb, .sht, .cls, .bas, .frm
    * 
    *
    *
    */
    public bool addFile(string sFile){
      return false;
    }

    public bool save(string sPath = ""){
      //Save to file
      return false;
    }


    //HELPERS

    private static string CreateUniqueTempDirectory()
    {
      string uniqueTempDir = Path.GetFullPath(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));
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
    private static UInt16 GetWord(byte[] buffer, int offset){
      var rawBytes = new byte[2];

      Array.Copy(buffer, offset, rawBytes, 0, 2);
      //if (!BitConverter.IsLittleEndian) {
      //	Array.Reverse(rawBytes);
      //}

      return BitConverter.ToUInt16(rawBytes, 0);
    }

    private static UInt32 GetDoubleWord(byte[] buffer, int offset){
      var rawBytes = new byte[4];

      Array.Copy(buffer, offset, rawBytes, 0, 4);
      //if (!BitConverter.IsLittleEndian) {
      //	Array.Reverse(rawBytes);
      //}

      return BitConverter.ToUInt32(rawBytes, 0);
    }
  }
  
  public class VBAModule {
    private string pModuleName; // Name of VBA module stream
    private UInt32 pTextOffset; // Offset of VBA source code in VBA module stream
    private byte[] pModuleStream; //Where the VBA Module byte code is stored.
    private VBAProject pParent;
    public bool hidden = false;

    //Constructor
    public VBAModule(VBAProject parent, byte[] moduleStream, string moduleName, UInt32 textOffset){
      pModuleStream = moduleStream;
      pModuleName = moduleName;
      pTextOffset = textOffset;
      pParent = parent;
    }

    public string moduleName {
      get {return moduleName;}
      set {
        //
        pModuleName = value;

        //Change in stream?

      }
    }

    public UInt32 textOffset {
      get {return pTextOffset;}
      set {
        //
        pTextOffset = value;
        
        //Change stream
      }
    }

    public string Source {
      get {
        //Skip bytes until pTextOffset
        byte[] bytes = pModuleStream.Skip((int)pTextOffset).ToArray();

        //Obtain module text as string
        string vbaModuleText = System.Text.Encoding.UTF8.GetString(VbaCompression.Decompress(bytes));

        //Return module text
        return vbaModuleText;
      }
      set {
        //Take the first pTextOffset chars, and concatenate compressed source.
        //Set pModuleStream to new value.
        pModuleStream = pModuleStream.Take((int)pTextOffset).Concat(VbaCompression.Compress(Encoding.UTF8.GetBytes(value))).ToArray();
      }
    }

    public byte[] build(){
      //TODO:
      //Create new module stream and return it from build.
      return pModuleStream;
    }
  
  // public static VBAModule FromFile(string sPath){
  //   //
    
  // }

  }
}