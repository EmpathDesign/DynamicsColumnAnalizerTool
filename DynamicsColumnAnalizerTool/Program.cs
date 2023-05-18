using System;
using ClosedXML;
using Classes;
using ClosedXML.Excel;

namespace Program
{
    class DynamicsColumnAnalizerTool
    {
        public static void Main(string[] args)
        {
            //get path to xlsx files
            Console.WriteLine("""Select An Inactive XLSX File   (format example: C\\Users\\user\\Desktop\\xlsxFile.xlsx)""");
            string dataPath = Console.ReadLine()+""; //concat for non null
            Console.WriteLine("""Select An Inactive XLSX Schema File   (format example: C\\Users\\user\\Desktop\\xlsxSchemaFile.xlsx)""");
            string schemaPath = Console.ReadLine() + ""; //concat for non null
            Console.WriteLine(dataPath + " & " + schemaPath);
            IXLWorkbook wb = new XLWorkbook(dataPath);
            IXLWorkbook schemaWb = new XLWorkbook(schemaPath);
            Table table = new Table(wb, schemaWb);
            //pause so i can read the console
            System.Threading.Thread.Sleep(5000);

        }
    }
}