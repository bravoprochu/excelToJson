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
            // onPostFile("baxModelSpec.xlsm");
            var interfaceFile = new ExcelToJsonHelper("baxModelSpec.xlsm", "i-bax-model-spec-list.ts");
            var datalist= new ExcelToJsonHelper("baxModelSpec.xlsm", "bax-model-spec-list.json");

            interfaceFile.GenInterface();
            datalist.GenData();

            // Console.Read();
        }




        static void onPostFile(string fileInput = "baxModelSpec.xlsm", string fileOutput = "bax-model-spec-list.ts")
        {
            //string fileFolder = Environment.CurrentDirectory;
            StringBuilder sb = new StringBuilder();
            MemoryStream memoryStream = new MemoryStream();

            string fileFolder = @"C:\\tool";
            //string fileInput = @"baxModelSpec.xlsm";
            //string fileOutput = @"test.txt";
            //var fileFolder = Environment.SpecialFolder.Desktop;

            StreamWriter streamWriter = new StreamWriter($"{fileFolder}\\{fileOutput}");

            FileStream fsInput;

            try
            {
                fsInput = new FileStream($"{fileFolder}\\{fileInput}", FileMode.Open, FileAccess.Read);
            }
            catch (Exception ex)
            {
                sb.Append(ex.Message);
                printIt(sb);
                return;
            }

            //if (fsInput.Length == 0) { return; }

            //fsInput.Position = 0;

            XSSFWorkbook hssfworkbook = new XSSFWorkbook(fsInput);
            string sheetName = "dane";
            // sheetName = Console.ReadLine();

            ISheet sheet = hssfworkbook.GetSheet(sheetName);
            if (sheet == null)
            {
                printIt($"Nie ma skoroszytu {sheetName}");
                return;
            }

            int dataRows = sheet.LastRowNum;

            //            sb.Append(rowCellsValues(sheet, 0));
            sb.Append(BuildInterface(sheet));
            sb.Append(BuildIt(sheet));


            streamWriter.Write(sb);
            streamWriter.Close();


            IRow row = sheet.GetRow(0);
            printIt(sb);

        }


        static int getColumnCountAtRow(ISheet sheet, int rowId = 0, int startColumnId = 0)
        {
            bool isCellEmpty = false;
            while (!isCellEmpty)
            {
                if (sheet.GetRow(rowId).GetCell(startColumnId) == null)
                {
                    isCellEmpty = true;
                }
                startColumnId++;
            }
            return startColumnId - 1;
        }

        
        static int getRowCountAtColumn(ISheet sheet, int columnId = 0, int startRowId = 0)
        {
            bool isCellEmpty = false;
            while (!isCellEmpty)
            {
                if (sheet.GetRow(startRowId)?.GetCell(columnId) == null)
                {
                    isCellEmpty = true;
                } else
                {
                    startRowId++;
                }
                
            }
            return startRowId;
        }


        static StringBuilder rowCellsValues(ISheet sheet, int rowId, int initColumnId = 0)
        {
            bool isCellEmpty = false;
            StringBuilder res = new StringBuilder();
            while (!isCellEmpty)
            {
                var actCell = sheet.GetRow(rowId).GetCell(initColumnId);
                res.Append(initColumnId > 0 ? " | " : null);
                res.Append(actCell);
                res.Append($" - {actCell.CellType}");
                initColumnId++;
                if (sheet.GetRow(rowId).GetCell(initColumnId) == null)
                {
                    isCellEmpty = true;
                }
            }

            return res;
        }


        static StringBuilder columnCellsValues(ISheet sheet, int columnId, int initRowId = 0)
        {
            bool isCellEmpty = false;
            StringBuilder res = new StringBuilder();
            while (!isCellEmpty)
            {
                var actCell = sheet.GetRow(initRowId).GetCell(columnId);
                res.Append(initRowId > 0 ? " | " : null);
                res.Append(actCell);
                res.Append($" - {actCell.CellType}");
                initRowId++;
                if (sheet.GetRow(initRowId).GetCell(columnId) == null)
                {
                    isCellEmpty = true;
                }
            }
            return res;
        }

        static CellType getFirstTypeInColumn(ISheet sheet, int columnId, int startRowId = 1)
        {
            bool found = false;
            int maxRows = getRowCountAtColumn(sheet);

            while (!found && (startRowId < maxRows))
            {
                var actCell = sheet.GetRow(startRowId).GetCell(columnId);
                if (actCell != null)
                {
                    return actCell.CellType;
                }
                startRowId++;
            }
            return CellType.String;
        }

        static void printIt(StringBuilder text)
        {
            Console.WriteLine("-----------------------");
            Console.WriteLine(text);
        }

        static void printIt(string text)
        {
            Console.WriteLine("-----------------------");
            Console.WriteLine(text);
        }

        static string getTypeBasedOnCell(CellType cellType)
        {
            string res;
            switch (cellType)
            {
                case CellType.Boolean:
                    res = "boolean";
                    break;
                case CellType.Numeric:
                    res = "number";
                    break;
                case CellType.String:
                    res = "string";
                    break;
                default:
                    res = "string";
                    break;
            }
            return res;
        }

        static StringBuilder BuildInterface(ISheet sheet) {
            var res = new StringBuilder();

            var colCount = getColumnCountAtRow(sheet);
            var rowCount = getRowCountAtColumn(sheet);
            var header = sheet.GetRow(0);
            if (rowCount < 2) { return res; }

            res.AppendLine("export interface IBAXModelSpec {");
            for (int c = 0; c < colCount; c++)
            {
                string headerCell = header.GetCell(c).ToString();
                string dataType = getTypeBasedOnCell(getFirstTypeInColumn(sheet, c));

                res.AppendLine($"{headerCell} : {dataType}");
            }
            res.AppendLine("}");



            return res;
        }

        static StringBuilder BuildIt(ISheet sheet, bool isObject = false)
        {
            var res = new StringBuilder("[");
            

            var colCount = getColumnCountAtRow(sheet);
            var rowCount = getRowCountAtColumn(sheet);
            var header = sheet.GetRow(0);

            if (rowCount < 2) { return res; }

            var firstRowData = sheet.GetRow(1);

            for (int r = 1; r < rowCount; r++)
            {
                var actRow = sheet.GetRow(r);
                for (int c = 0; c < colCount; c++)
                {
                    string headerCell = header.GetCell(c).ToString();
                    
                    var _cellType = getFirstTypeInColumn(sheet, c);
                    bool isNumeric = _cellType == CellType.Numeric ? true : false;

                    string _cellTypeRes = getTypeBasedOnCell(_cellType);

                        headerCell = "\"" + headerCell + "\"";
                        if (c == 0)
                        {
                            res.AppendLine("{");
                        }

                        if (actRow.GetCell(c) == null)
                        {
                            res.Append($"{headerCell}: null");
                        }
                        else
                        {
                            if (isNumeric)
                            {
                                var v = string.Format(actRow.GetCell(c).ToString()).Replace(",", ".");
                                res.Append($"{headerCell}:  {v}");
                            }
                            else
                            {
                                res.Append($"{headerCell}: "+ "\"" + actRow.GetCell(c) + "\"");
                            }
                        }


                    if (c == colCount - 1)
                    {
                        res.AppendLine("}");
                    }
                    else {
                        res.Append(",");
                        res.AppendLine();
                    }
                }
                if (r < rowCount - 1)
                {
                    res.Append(",");
                }
            }
            res.AppendLine("]");
            return res;
        }







    }

}
