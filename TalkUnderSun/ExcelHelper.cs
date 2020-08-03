using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using LumenWorks.Framework.IO.Csv;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TalkUnderSun
{
    public class ExcelHelper
    {
        /// <summary>
        /// 导出EXCEL
        /// </summary>
        /// <param name="data">导出数据</param>
        /// <param name="type">导出类型</param>
        /// <returns></returns>
        public MemoryStream Export(DataTable data, EnumExcelType type)
        {
            var util = TransferDataFactory.GetUtil(type);
            var mStream = util.GetStream(data);
            var ms = (MemoryStream)mStream;
            return ms;
        }
    }

    /// <summary>
    /// 接口
    /// </summary>
    public interface ITransferData
    {
        Stream GetStream(DataTable table);//当为Stream时需要测试
        DataTable GetData(Stream stream);
    }
    /// <summary>
    /// 工厂类
    /// </summary>
    public class TransferDataFactory
    {
        public static ITransferData GetUtil(string fileName)
        {
            var array = fileName.Split('.');
            var typeName = array.Last().ToUpper();
            var dataType = (EnumExcelType)System.Enum.Parse(typeof(EnumExcelType), typeName);
            return GetUtil(dataType);
        }
        public static ITransferData GetUtil(EnumExcelType dataType)
        {
            switch (dataType)
            {
                case EnumExcelType.CSV: return new CsvTransferData();
                case EnumExcelType.XLS: return new XlsTransferData();
                case EnumExcelType.XLSX: return new XlsxTransferData();
                default: return new CsvTransferData();
            }
        }
    }
    /// <summary>
    /// 实现基类
    /// </summary>
    public abstract class ExcelTransferData : ITransferData
    {
        protected IWorkbook _workBook;
        public virtual Stream GetStream(DataTable table)
        {
            var sheet = _workBook.CreateSheet("Sheet1");
            if (table != null)
            {
                var rowCount = table.Rows.Count;
                var columnCount = table.Columns.Count;
                var row = sheet.CreateRow(0);
                for (int i = 0; i < columnCount; i++)
                {//将DataTable ColumnName作为第一行数据
                    var cell = row.CreateCell(i);
                    cell.SetCellValue(table.Columns[i].ColumnName);
                }
                for (int i = 0; i < rowCount; i++)
                {
                    var row1 = sheet.CreateRow(i + 1);
                    for (int j = 0; j < columnCount; j++)
                    {
                        var cell = row1.CreateCell(j);
                        if (table.Rows[i][j] != null)
                            cell.SetCellValue(table.Rows[i][j].ToString());
                    }
                }
            }
            var ms = new System.IO.MemoryStream();
            _workBook.Write(ms);
            return ms;
        }
        public virtual DataTable GetData(Stream stream)
        {
            using (stream)
            {
                var sheet = _workBook.GetSheetAt(0);
                if (sheet != null)
                {
                    var headerRow = sheet.GetRow(0);
                    DataTable dt = new DataTable();
                    int columnCount = headerRow.Cells.Count;
                    for (int i = 0; i < columnCount; i++)
                    {//第一行作为标题行
                        //dt.Columns.Add("col_" + i.ToString());
                        string cloName = headerRow.Cells[i].ToString();
                        if (!string.IsNullOrEmpty(cloName))
                            dt.Columns.Add(cloName);
                    }
                    var row = sheet.GetRowEnumerator();
                    while (row.MoveNext())
                    {
                        var dtRow = dt.NewRow();
                        var excelRow = row.Current as IRow;
                        for (int i = 0; i < columnCount; i++)
                        {
                            var cell = excelRow.GetCell(i);
                            if (cell != null)
                                dtRow[i] = GetValue(cell);
                        }
                        dt.Rows.Add(dtRow);
                    }
                    dt.Rows.RemoveAt(0);//去除第一行，第一行作为标题行
                    return dt;
                }
            }
            return null;
        }
        private object GetValue(ICell cell)
        {
            object value = null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    break;
                case CellType.Boolean:
                    value = cell.BooleanCellValue ? "1" : "0"; break;
                case CellType.Error:
                    value = cell.ErrorCellValue; break;
                case CellType.Formula:
                    value = "=" + cell.CellFormula; break;
                case CellType.Numeric:
                    if (HSSFDateUtil.IsCellDateFormatted(cell))//判断是否为日期格式 
                        value = cell.DateCellValue;
                    else
                        value = cell.NumericCellValue.ToString();
                    break;
                case CellType.String:
                    value = cell.StringCellValue; break;
                case CellType.Unknown:
                    break;
            }
            return value;
        }
    }
    /// <summary>
    /// 2007版本实现
    /// </summary>
    public class XlsxTransferData : ExcelTransferData
    {
        public override Stream GetStream(DataTable table)
        {
            base._workBook = new XSSFWorkbook();
            return base.GetStream(table);
        }
        public override DataTable GetData(Stream stream)
        {
            base._workBook = new XSSFWorkbook(stream);
            return base.GetData(stream);
        }
    }
    /// <summary>
    /// 2003版本实现
    /// </summary>
    public class XlsTransferData : ExcelTransferData
    {
        public override Stream GetStream(DataTable table)
        {
            base._workBook = new HSSFWorkbook();
            return base.GetStream(table);
        }
        public override DataTable GetData(Stream stream)
        {
            base._workBook = new HSSFWorkbook(stream);
            return base.GetData(stream);
        }
    }
    /// <summary>
    /// csv版本实现
    /// </summary>
    public class CsvTransferData : ITransferData
    {
        private Encoding _encode;
        public CsvTransferData()
        {
            //this._encode = Encoding.GetEncoding("utf-8");
            this._encode = Encoding.Default;
        }
        public Stream GetStream(DataTable table)
        {
            var sb = new StringBuilder();
            if (table != null && table.Columns.Count > 0 && table.Rows.Count > 0)
            {
                var columnCount = table.Columns.Count;
                for (int i = 0; i < columnCount; i++)
                {//将DataTable ColumnName作为第一行数据
                    var name = table.Columns[i].ColumnName;
                    if (i > 0)
                        sb.Append(",");
                    if (!string.IsNullOrEmpty(name))
                        sb.Append("\"").Append(name.Replace("\"", "\"\"")).Append("\"");
                }
                sb.Append("\n");
                foreach (DataRow item in table.Rows)
                {
                    for (int i = 0; i < columnCount; i++)
                    {
                        if (i > 0)
                            sb.Append(",");
                        if (item[i] != null)
                            sb.Append("\"").Append(item[i].ToString().Replace("\"", "\"\"")).Append("\"");
                    }
                    sb.Append("\n");
                }
            }
            MemoryStream stream = new MemoryStream(_encode.GetBytes(sb.ToString()));
            return stream;
        }
        public DataTable GetData(Stream stream)
        {
            using (stream)
            {
                using (StreamReader input = new StreamReader(stream, _encode))
                {
                    using (CsvReader csv = new CsvReader(input, false))
                    {
                        DataTable dt = new DataTable();
                        int columnCount = csv.FieldCount;
                        while (csv.ReadNextRecord())
                        {
                            if (csv.CurrentRecordIndex == 0)
                            {//将第一行作为列名
                                for (int i = 0; i < columnCount; i++)
                                {
                                    if (!string.IsNullOrWhiteSpace(csv[i]))
                                        dt.Columns.Add(csv[i]);
                                }
                            }
                            else
                            {
                                var dr = dt.NewRow();
                                for (int i = 0; i < columnCount; i++)
                                {
                                    if (!string.IsNullOrWhiteSpace(csv[i]))
                                        dr[i] = csv[i];
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                        return dt;
                    }
                }
            }
        }
    }
}
