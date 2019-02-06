using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using excelToJson;
using System.Threading.Tasks;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace excelToJson_console
{
    class Program
    {


        static void Main(string[] args)
        {
            Console.WriteLine("generuje plik....");

            var interfaceFile = new ExcelToJsonHelper("baxModelSpec.xlsm", "i-bax-model-spec-list.ts");
            var datalist= new ExcelToJsonHelper("baxModelSpec.xlsm", "bax-model-spec-list.json");

            interfaceFile.GenInterface();
            datalist.GenData();
            // Console.Read();
        }

    }

}
