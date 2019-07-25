using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Excel2vCard
{

    public class NPOIExcelIO
    {
        #region read
        public bool ifTrimSpace = true;

        IWorkbook GetWorkBook(FileStream stream, string fileName)
        {
            string extension = Path.GetExtension(fileName);
            switch (extension)
            {
                case ".xlsx":
                    return new XSSFWorkbook(stream);
                case ".xls":
                    return new HSSFWorkbook(stream);
                default:
                    return null;
            }
        }

        public DataTable Read(string fileName, List<string> strReadColumns, bool hasHeader = true, string sheetName = null)
        {
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = GetWorkBook(fileStream, fileName);
                    ISheet sheet;
                    if (string.IsNullOrEmpty(sheetName))
                        sheet = workbook.GetSheetAt(0);
                    else
                        sheet = workbook.GetSheet(sheetName);

                    if (sheet == null)
                        return null;
                    else
                    {
                        List<int> indexReadColumns = GetColumnsIndex(sheet.GetRow(0), strReadColumns);
                        if (indexReadColumns.Count != strReadColumns.Count)
                            return null;

                        DataTable dt = new DataTable();
                        ReadColumns(dt, sheet.GetRow(0), indexReadColumns);
                        ReadRows(sheet, dt, hasHeader, indexReadColumns);

                        return dt;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public DataTable Read(string fileName, List<int> indexReadColumns, bool hasHeader = true, string sheetName = null)
        {
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = GetWorkBook(fileStream, fileName);

                    ISheet sheet;
                    if (string.IsNullOrEmpty(sheetName))
                        sheet = workbook.GetSheetAt(0);
                    else
                        sheet = workbook.GetSheet(sheetName);

                    if (sheet == null)
                        return null;
                    else
                    {
                        if (!ifExistsNum(sheet, indexReadColumns[indexReadColumns.Count - 1]))
                            return null;

                        DataTable dt = new DataTable();
                        ReadColumns(dt, sheet.GetRow(0), indexReadColumns);
                        ReadRows(sheet, dt, hasHeader, indexReadColumns);

                        return dt;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public DataTable Read(string fileName, bool hasHeader = true, string sheetName = null)
        {
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = GetWorkBook(fileStream, fileName);

                    ISheet sheet;
                    if (string.IsNullOrEmpty(sheetName))
                        sheet = workbook.GetSheetAt(0);
                    else
                        sheet = workbook.GetSheet(sheetName);

                    if (sheet == null)
                        return null;
                    else
                    {
                        DataTable dt = new DataTable();
                        ReadColumns(dt, sheet);
                        ReadRows(sheet, dt, hasHeader);

                        return dt;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public DataTable Read(string fileName, int headerRowIndex, string sheetName = null)
        {
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = GetWorkBook(fileStream, fileName);

                    ISheet sheet;
                    if (string.IsNullOrEmpty(sheetName))
                        sheet = workbook.GetSheetAt(0);
                    else
                        sheet = workbook.GetSheet(sheetName);

                    if (sheet == null)
                        return null;
                    else
                    {
                        DataTable dt = new DataTable();
                        var headRow = sheet.GetRow(headerRowIndex);
                        int headCellsCount = headRow.Cells.Count;
                        for (int i = 0; i < headCellsCount; i++)
                        {
                            var headCell = headRow.Cells[i];
                            var fieldName = ColumnReadCell(headCell);
                            dt.Columns.Add(fieldName);
                        }
                        int lastRowIndex = sheet.LastRowNum;
                        for (int j = headerRowIndex + 1; j < lastRowIndex; j++)
                        {
                            var row = sheet.GetRow(j);
                            object[] drValue = new object[headCellsCount];
                            for (int i = 0; i < headCellsCount; i++)
                            {
                                var _value = ColumnReadCell(row.GetCell(i));
                                drValue[i] = _value;
                            }

                            dt.Rows.Add(drValue);
                        }
                        return dt;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        void ReadRows(ISheet sheet, DataTable dt, bool hasHeader, List<int> indexReadColumns = null)
        {
            int rowCount = sheet.LastRowNum;
            int i = 0;
            if (hasHeader)
                i = 1;
            for (; i <= rowCount; i++)
            {
                IRow row = sheet.GetRow(i) as IRow;
                if (indexReadColumns == null) ReadRow(dt, row);
                else ReadRow(dt, row, indexReadColumns);
            }
        }

        void ReadRow(DataTable dt, IRow row, List<int> indexReadColumns = null)
        {
            DataRow dr = dt.NewRow();
            if (row != null)
            {

                if (indexReadColumns == null)
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ICell cell = row.GetCell(i) as ICell;
                        dr[i] = ContentReadCell(cell);
                    }
                else
                    for (int i = 0; i < indexReadColumns.Count; i++)
                    {
                        ICell cell = row.GetCell(indexReadColumns[i]) as ICell;
                        dr[i] = ContentReadCell(cell);
                    }
            }
            dt.Rows.Add(dr);
        }

        string ColumnReadCell(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        if (ifTrimSpace)
                            return cell.StringCellValue.Trim();
                        else
                            return cell.StringCellValue;
                    case CellType.Blank:
                        return string.Empty;
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                            return cell.DateCellValue.ToString();
                        else
                            return cell.NumericCellValue.ToString();
                    case CellType.Error:
                        return cell.ErrorCellValue.ToString();
                    case CellType.Formula:
                        if (ifTrimSpace)
                            return cell.StringCellValue.Trim();
                        else
                            return cell.StringCellValue;
                    default:
                        return string.Empty; ;
                }
            }
            catch
            {
                return string.Empty; ;
            }
        }

        object ContentReadCell(ICell cell)
        {
            if (cell == null)
                return DBNull.Value;
            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        if (ifTrimSpace)
                            return cell.StringCellValue.Trim();
                        else
                            return cell.StringCellValue;
                    case CellType.Blank:
                        return DBNull.Value;
                    case CellType.Boolean:
                        return cell.BooleanCellValue;
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                            return cell.DateCellValue;
                        else
                            return cell.NumericCellValue;
                    case CellType.Error:
                        return cell.ErrorCellValue;
                    case CellType.Formula:
                        if (ifTrimSpace)
                            return cell.StringCellValue.Trim();
                        else
                            return cell.StringCellValue;
                    default:
                        return DBNull.Value;
                }
            }
            catch
            {
                return DBNull.Value;
            }
        }

        bool ifExistsNum(ISheet sheet, int number)
        {
            int rowCount = sheet.LastRowNum;
            int i = 0;
            for (; i <= rowCount; i++)
            {
                IRow row = sheet.GetRow(i) as IRow;
                if (row != null)
                    if (row.LastCellNum > number)
                        return true;
            }
            return false;
        }

        List<int> GetColumnsIndex(IRow row, IEnumerable<string> strColumns)
        {
            List<int> indexList = new List<int>();

            int rowCount = row.LastCellNum;
            foreach (string str in strColumns)
            {
                for (int i = 0; i < rowCount; i++)
                {
                    ICell cell = row.GetCell(i);

                    string value = ColumnReadCell(cell);
                    if (string.IsNullOrEmpty(value))
                        continue;

                    if (str.Equals(value))
                    {
                        indexList.Add(i);
                        break;
                    }

                }
            }
            return indexList;
        }

        void ReadColumns(DataTable dt, ISheet sheet)
        {
            int rowCount = sheet.LastRowNum;
            int j = 0;
            for (; j <= rowCount; j++)
            {
                IRow row = sheet.GetRow(j) as IRow;
                if (row != null)
                    if (j == 0)
                        for (int i = 0; i < row.LastCellNum; i++)
                        {
                            ICell cell = row.GetCell(i);
                            string value = ColumnReadCell(cell);
                            if (string.IsNullOrEmpty(value))
                                dt.Columns.Add();
                            else
                            {
                                if (!dt.Columns.Contains(value))
                                    dt.Columns.Add(value);
                                else
                                    dt.Columns.Add();
                            }
                        }
                    else
                    {
                        if (row.LastCellNum > dt.Columns.Count)
                            for (int i = dt.Columns.Count; i < row.LastCellNum; i++)
                                dt.Columns.Add();
                    }

            }
        }

        void ReadColumns(DataTable dt, IRow row, List<int> indexReadColumns)
        {
            foreach (int i in indexReadColumns)
            {
                ICell cell = row.GetCell(i);
                string value = ColumnReadCell(cell);
                if (string.IsNullOrEmpty(value))
                    dt.Columns.Add();
                else
                {
                    if (!dt.Columns.Contains(value))
                        dt.Columns.Add(value);
                    else
                        dt.Columns.Add();
                }
            }
        }
        #endregion

        #region write
        public bool Write(DataTable dt, string fileName)
        {
            try
            {
                using (FileStream fs = File.Create(fileName))
                {
                    IWorkbook workbook = GetWorkBook(fileName);
                    ISheet wSheet = workbook.CreateSheet();

                    int rowIndex = 0;
                    int columnIndex = 0;

                    IRow row = wSheet.CreateRow(rowIndex++);
                    foreach (DataColumn dc in dt.Columns)
                    {
                        ICell cell = row.CreateCell(columnIndex++);
                        cell.SetCellType(CellType.String);
                        cell.SetCellValue(dc.ColumnName);
                    }

                    int columnCount = dt.Columns.Count;
                    foreach (DataRow dr in dt.Rows)
                    {
                        row = wSheet.CreateRow(rowIndex++);
                        for (int i = 0; i < columnCount; i++)
                        {
                            ICell cell = row.CreateCell(i);
                            cell.SetCellType(CellType.String);
                            cell.SetCellValue(dr[i].ToString());
                        }
                    }

                    workbook.Write(fs);
                    fs.Close();

                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        IWorkbook GetWorkBook(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            switch (extension)
            {
                case ".xls":
                    return new HSSFWorkbook();
                case ".xlsx":
                    return new XSSFWorkbook();
                default:
                    return null;
            }
        }
        #endregion
    }
}