using System;
using VBAPackager;

class Program {
    static void Main(string[] args) {
        string pth = "Test.xlsm";
        Console.WriteLine("Decompressing " + pth);
        VBAProject x = new VBAProject(pth);
        return;
    }
}

