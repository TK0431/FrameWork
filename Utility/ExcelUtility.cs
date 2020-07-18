using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace FrameWork.Utility
{
    public class ExcelUtility : IDisposable
    {

        #region Common

        /// <summary>
        /// Excel
        /// </summary>
        private ExcelPackage _excel;

        /// <summary>
        /// Ctor
        /// </summary>
        public ExcelUtility(string excelName)
        {
            // License
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Excel
            _excel = new ExcelPackage(new FileInfo(excelName));
        }

        /// <summary>
        /// Ctor
        /// </summary>
        public ExcelUtility()
        {
            // License
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Excel
            _excel = new ExcelPackage();
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            if (_excel != null) _excel.Dispose();
        }

        #endregion

        /// <summary>
        /// Load DataTable Data
        /// </summary>
        /// <param name="data"></param>
        public void LoadDataTable(DataTable data, string posi = "A2")
            => this.GetSheet().Cells[posi].LoadFromDataTable(data, false, TableStyles.Medium9);

        /// <summary>
        /// Load DataTable Data
        /// </summary>
        /// <param name="list"></param>
        public void LoadDataList<T>(List<T> list, string posi = "A2")
        {
            DataTable data = new DataTable();
            List<PropertyInfo> properties = typeof(T).GetProperties()
                .Where(x => x.GetCustomAttribute<DisplayAttribute>() != null).ToList()
                .OrderBy(x => x.GetCustomAttribute<DisplayAttribute>().Order).ToList();

            properties.ForEach(p => data.Columns.Add(p.Name));

            foreach (var item in list)
            {
                DataRow row = data.NewRow();
                int i = 0;
                properties.ForEach(p =>
                {
                    row[i++] = p.GetValue(item);
                });
                data.Rows.Add(row);
            }

            this.LoadDataTable(data, posi);
        }

        /// <summary>
        /// Add Sheet
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelWorksheet AddSheet(string sheetName)
            => _excel.Workbook.Worksheets.Add(sheetName);

        /// <summary>
        /// Add Sheet
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelWorksheet GetSheet(int index = 0)
            => _excel.Workbook.Worksheets[index];

        /// <summary>
        /// Add Sheet
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelWorksheet GetSheet(string sheetName)
            => _excel.Workbook.Worksheets[sheetName];

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public int GetMaxRow(ExcelWorksheet sheet, int col)
        {
            for (int i = 1; i <= sheet.Cells.Rows; i++)
                if (string.IsNullOrWhiteSpace(sheet.Cells[i, col].Text))
                    return i - 1;
            return sheet.Cells.Rows;
        }

        /// <summary>
        /// Save
        /// </summary>
        public void Save()
            => _excel.Save();

        /// <summary>
        /// SaveAs
        /// </summary>
        /// <param name="file"></param>
        public void SaveAs(FileInfo file)
            => _excel.SaveAs(file);

        /// <summary>
        /// SaveAs
        /// </summary>
        /// <param name="file"></param>
        public void SaveAs(Stream file)
            => _excel.SaveAs(file);

        public void RefreshAll()
        {
            _excel.Workbook.FullCalcOnLoad = true;
        }
    }

    public static class ExcelHelper
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static int GetMaxRow(this ExcelWorksheet sheet)
            => sheet.Dimension.End.Row;

        public static int GetMaxRow(this ExcelWorksheet sheet, int col)
        {
            int cntRow = sheet.GetMaxRow();
            for (; cntRow > 1; cntRow--)
                if (!string.IsNullOrWhiteSpace(sheet.Cells[cntRow, col].Text))
                    break;
            return cntRow;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static int GetMaxColumn(this ExcelWorksheet sheet)
            => sheet.Dimension.End.Column;

        public static int GetMaxColumn(this ExcelWorksheet sheet, int row)
        {
            int cntCol = sheet.GetMaxColumn();
            for (; cntCol > 1; cntCol--)
                if (!string.IsNullOrWhiteSpace(sheet.Cells[row, cntCol].Text))
                    break;
            return cntCol;
        }

        public static int LoadData(this ExcelWorksheet sheet, string file, int row = 1, int col = 1)
        {
            ExcelTextFormat format = new ExcelTextFormat();
            format.Encoding = Encoding.GetEncoding(932);

            string[] lens = File.ReadAllLines(file);

            if (lens.Length > 0)
            {
                List<eDataTypes> allTypes = new List<eDataTypes>();
                Enumerable.Range(0, lens[0].Split(',').Length).ToList().ForEach(x => allTypes.Add(eDataTypes.String));
                format.DataTypes = allTypes.ToArray();
            }

            sheet.Cells[row, col].LoadFromText(new FileInfo(file), format);

            if (lens.Length > 0)
            {
                sheet.Cells[row, col, row, sheet.GetMaxColumn(row)].SetRangeColor(Color.FromArgb(91, 155, 213));
                sheet.Cells[row, col, row, sheet.GetMaxColumn(row)].SetRangeBorder(Color.DarkGray);
                sheet.Cells[row, col, row, sheet.GetMaxColumn(row)].SetFontColor(Color.White);
            }

            if (lens.Count() <= 0) return 0;
            if (lens.Count() == 1 && string.IsNullOrEmpty(lens[0])) return 0;

            return lens.Length;
        }

        public static ExcelShape AddShape(this ExcelWorksheet sheet, string name)
        {
            ExcelShape es = sheet.Drawings.AddShape(name, eShapeStyle.Rect);
            es.Border.LineStyle = eLineStyle.Solid;
            es.Border.Width = 1.5;
            es.Border.Fill.Color = Color.Red;
            es.Fill.Style = eFillStyle.NoFill;

            return es;
        }

        public static void SetRangeBorder(this ExcelRange range, ExcelBorderStyle type = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Left.Style = type;
            range.Style.Border.Top.Style = type;
            range.Style.Border.Right.Style = type;
            range.Style.Border.Bottom.Style = type;
        }

        public static void SetRangeBorder(this ExcelRange range, Color color, ExcelBorderStyle type = ExcelBorderStyle.Thin)
        {
            range.SetRangeBorder(type);
            range.Style.Border.Left.Color.SetColor(color);
            range.Style.Border.Top.Color.SetColor(color);
            range.Style.Border.Right.Color.SetColor(color);
            range.Style.Border.Bottom.Color.SetColor(color);
        }

        public static void SetFontColor(this ExcelRange range, Color color)
            => range.Style.Font.Color.SetColor(color);

        public static void SetRangeColor(this ExcelRange range, Color color)
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }

        public static void SetColor(this ExcelRow er, Color color)
        {
            er.Style.Fill.PatternType = ExcelFillStyle.Solid;
            er.Style.Fill.BackgroundColor.SetColor(color);
        }

        public static int GetWidthPix(this ExcelWorksheet sheet)
        {
            Font font = new Font(sheet.Cells["A1"].Style.Font.Name, sheet.Cells["A1"].Style.Font.Size, FontStyle.Regular);
            return (int)Math.Round((sheet.Column(1).Width * Math.Round((Graphics.FromHwnd(IntPtr.Zero).MeasureString("1234567890", font, int.MaxValue, StringFormat.GenericTypographic).Width / 10))));
        }

        public static int GetHeightPix(this ExcelWorksheet sheet)
            => (int)((GetDiY() * sheet.Row(1).Height) / 72);

        public static int GetDiY()
            => (int)Graphics.FromHwnd(IntPtr.Zero).DpiY;

        public static double GetRate()
        {
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                switch ((int)graphics.DpiY)
                {
                    case 96:
                        return 0.75;
                    case 120:
                        return 0.6;
                    case 144:
                        return 0.5;
                    case 192:
                        return 0.375;
                    default:
                        // 72
                        return 1;
                }
            }
        }
    }
}
