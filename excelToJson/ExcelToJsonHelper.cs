using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace excelToJson
{
    public class ExcelToJsonHelper {

        public ExcelToJsonHelper(string fileInputName = "baxModelSpec.xlsm", string fileOutputName = "bax-model-spec-list.ts", string fileFolder = "c:\\tool\\")
        {
            this._fileInputName = fileInputName;
            this._fileOutputName = fileOutputName;
            this._fileFolder = fileFolder;
            this.SheetName = "dane";
            this.InterfaceName = "IBaxModelSpec";
            this._stringBuilderError = new StringBuilder();
            this.InitInputFile();
        }

        private int _colsCount { get; set; }
        private string _fileInputName { get; set; }
        private string _fileOutputName { get; set; }
        private string _fileFolder { get; set; }
        private List<KeyValuePair<string, string>> _header { get; set; }
        private FileStream _fileInputStream { get; set; }

        private int CountColumnsOnRowId(int rowId)
        {
            if (!_isSheetReady) { return 0; }

            bool keepSearching = true;
            int colId = 0;

            while (keepSearching)
            {
                if (this._selectedSheet.GetRow(rowId).GetCell(colId) == null)
                {
                    keepSearching = false;
                }
                else
                {
                    colId++;
                }
            }
            return colId;
        }

        private int CountRowsOnColumnId(int columnId)
        {
            if (!_isSheetReady) { return 0; }

            bool keepSearching = true;
            int rowId = 0;

            while (keepSearching)
            {
                if (this._selectedSheet.GetRow(rowId)?.GetCell(columnId) == null)
                {
                    keepSearching = false;
                }
                else
                {
                    rowId++;
                }
            }
            return rowId;
        }

        public void GenErrorFile() {
            var idx = this._fileInputName.IndexOf('.');
            var _filename = this._fileInputName.Substring(0, idx);

            this.streamWriter = new StreamWriter($"{_fileFolder}\\{_filename}_errorLog.txt");
            this.streamWriter.Write(this._stringBuilderError);
            this.streamWriter.Close();
        }
        public void GenInterface(bool isOptional=true) {
            this._isOptional = isOptional;

            this.initSheetCountsHeader();
            this.PrepInterface();

            if (this._isInterfaceReady)
            {
                this.streamWriter = new StreamWriter($"{_fileFolder}\\{_fileOutputName}");
                this.streamWriter.Write(this._stringBuilder);
                this.streamWriter.Close();
            }
            else {
                this.GenErrorFile();
            }
        }
        public void GenData(bool isObject=true) {
            this._isObject = isObject;

            this.initSheetCountsHeader();
            this.PrepData();

            if (this._isDataReady)
            {
                this.streamWriter = new StreamWriter($"{this._fileFolder}\\{this._fileOutputName}");
                this.streamWriter.Write(this._stringBuilder);
                this.streamWriter.Close();
            }
            else {
                this.GenErrorFile();
            }
            
        }

        private string GetCellTypeScriptType(ICell cell)
        {
            string res;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    res = "number";
                    break;
                case CellType.String:
                    res = "string";
                    break;
                case CellType.Boolean:
                    res = "boolean";
                    break;
                default:
                    res = "string";
                    break;
            }
            return res;
        }
        private ICell GetFirstNotEmptyCellByColumn(int columnId, int initRow = 0)
        {
            if (!this._isSheetDataReady || initRow > _rowsCount) { return null; }

            for (int r = initRow; r < this._rowsCount; r++)
            {
                var _cell = this._selectedSheet.GetRow(r).GetCell(columnId);
                if (_cell != null)
                {
                    return _cell;
                }
            }
            return null;
        }
        private ICell GetFirstNotEmptyCellByRow(int rowId, int initCol = 0)
        {

            if (!this._isSheetDataReady || initCol > this._colsCount) { return null; }



            for (int c = initCol; c < this._rowsCount; c++)
            {
                var _cell = this._selectedSheet.GetRow(rowId).GetCell(c);
                if (_cell != null)
                {
                    return _cell;
                }
            }
            return null;
        }
        public string InterfaceName { get; set; }
        private bool _isDataReady { get; set; }
        private bool _isInputDataReady { get; set; }
        private bool _isInterfaceReady { get; set; }
        private bool _isObject { get; set; }
        private bool _isOptional { get; set; }
        private bool _isSheetReady { get; set; }
        private bool _isSheetDataReady { get; set; }
        private bool _isSheetHeaderReady { get; set; }
        private int _rowsCount { get; set; }

        private ISheet _selectedSheet { get; set; }
        private StringBuilder _stringBuilder { get; set; }
        private StringBuilder _stringBuilderError { get; set; }
        private XSSFWorkbook _workbook { get; set; }

        public string SheetName { get; set; }
        private StreamWriter streamWriter { get; set; }
        private void PrepInterface() {
            if (!this._isSheetHeaderReady) { return; }
            this._isInterfaceReady = false;

            this._stringBuilder = new StringBuilder();
            this._stringBuilder.AppendLine($"export interface {this.InterfaceName}" + " {");
            foreach (var pos in this._header)
            {
                string _doubleQuouts = this._isObject ? null : "\"";
                string _optional = this._isOptional ? "?" : null;
                this._stringBuilder.AppendLine(_doubleQuouts + $"{pos.Key}"+ _doubleQuouts + _optional  + $": {pos.Value}");
            }
            this._stringBuilder.AppendLine("}");
            this._isInterfaceReady = true;
        }
        private void PrepData() {
            if (!this._isSheetHeaderReady) { return; }
            this._isInterfaceReady = false;

            this._stringBuilder = new StringBuilder();
            this._stringBuilder.AppendLine($"export const BaxModelSpecList = <{InterfaceName}>[");
            var _doubleQuotes = this._isObject ? null : "\"";

            for (int r = 1; r < _rowsCount; r++)
            {
                this._stringBuilder.AppendLine("{");
                var _actRow = this._selectedSheet.GetRow(r);

                for (int c = 0; c < this._colsCount; c++)
                {
                    var _activeHeader = _header[c];
                    var _activeValue = _actRow.GetCell(c);
                    var _doubleQuotesValue = _activeHeader.Value == "string" ? "\"" : null;
                    var _actCell = _actRow.GetCell(c);

                    string _value = "";
                    if (_actRow.GetCell(c) == null)
                    { _value = "null"; }
                    else
                    {
                        if (_actCell.CellType == CellType.Boolean) {
                            _value = _actCell.BooleanCellValue == true ? "true" : "false";
                        }

                        if (_actCell.CellType == CellType.Numeric)
                        {
                            _value = _actCell.ToString().Replace(",", ".");
                        }
                        if(_actCell.CellType == CellType.String)
                        {
                            _value = _doubleQuotesValue + _actRow.GetCell(c) + _doubleQuotesValue;
                        }
                    }
                    this._stringBuilder.AppendLine($"{_doubleQuotes}{_activeHeader.Key}{_doubleQuotes} : {_value},");
                    if (c == _colsCount - 1) {
                        this._stringBuilder.AppendLine("},");
                    }
                }
            
            }
            this._stringBuilder.AppendLine("]");
            this._isDataReady = true;
        }
        private void initSheetCountsHeader() {
            this.InitSheet();
            this.initSheetColRowCount();
            this.initSheetHeader();
        }
        private void InitInputFile() {
            this._isInputDataReady = false;

            try
            {
                this._fileInputStream = new FileStream($"{this._fileFolder}\\{this._fileInputName}", FileMode.Open, FileAccess.Read);
            }
            catch (Exception ex)
            {
                this._stringBuilderError.AppendLine("Init input file error: ");
                this._stringBuilderError.AppendLine(ex.Message?.ToString());
                // this._isInputDataReady = false;
                return;
                // throw ex.InnerException;
            }

            this._isInputDataReady = true;
        }
        private void InitSheet() {
            if (!this._isInputDataReady) { return; }
            this._isSheetReady = false;

            this._workbook = new XSSFWorkbook(this._fileInputStream);
            this._selectedSheet = _workbook.GetSheet(this.SheetName);
            if (_selectedSheet!=null)
            {
                this._isSheetReady = true;
            }
            else {
                this._stringBuilderError.AppendLine("Init sheet");
                this._isSheetReady = false;
            }

        }
        private void initSheetColRowCount() {
            if (!_isSheetReady) { return; }
            this._isSheetDataReady = false;

            this._colsCount = this.CountColumnsOnRowId(0);
            this._rowsCount = this.CountRowsOnColumnId(0);

            if (_rowsCount > 1 && _colsCount > 0) {
                this._isSheetDataReady = true;
            }
        }
        private void initSheetHeader() {
            if (!this._isSheetDataReady) { return; }
            this._isSheetHeaderReady = false;

            this._header = new List<KeyValuePair<string, string>>();

            for (int c = 0; c < this._colsCount; c++)
            {
                var firstRowValue = this._selectedSheet.GetRow(0).GetCell(c).ToString();
                string firstRowValueType;
                var firstNotEmptyCell = this.GetFirstNotEmptyCellByColumn(c, 1);
                if (firstNotEmptyCell != null)
                {
                    firstRowValueType = this.GetCellTypeScriptType(firstNotEmptyCell);
                    this._header.Add(new KeyValuePair<string, string>(firstRowValue, firstRowValueType));
                }
                else {
                    // if column is empty - set default type to string
                    this._header.Add(new KeyValuePair<string, string>(firstRowValue, "string"));
                }
            }
            this._isSheetHeaderReady = true;
        }






    }
}
