using Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Font = System.Drawing.Font;

namespace FrameWork.Utility
{
    public static class ExcelHelper
    {
        #region Epplus

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static int GetMaxRow(this ExcelWorksheet sheet)
            => sheet.Dimension == null ? 1 : sheet.Dimension.End.Row;

        public static int GetMaxRow(this ExcelWorksheet sheet, int col)
        {
            int cntRow = sheet.GetMaxRow();
            for (; cntRow > 1; cntRow--)
                if (!string.IsNullOrWhiteSpace(sheet.Cells[cntRow, col].Text) || !string.IsNullOrWhiteSpace(sheet.Cells[cntRow, col].Formula))
                    break;
            return cntRow;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static int GetMaxColumn(this ExcelWorksheet sheet)
            => sheet.Dimension == null ? 1 : sheet.Dimension.End.Column;

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
            //format.Encoding = Encoding.GetEncoding(932);
            format.Delimiter = ',';

            List<string> lens = File.ReadAllLines(file, Encoding.GetEncoding(932)).ToList();

            if (lens.Count > 0)
            {
                if (lens[0].StartsWith("\"")) format.Delimiter = '\0';

                List<eDataTypes> allTypes = new List<eDataTypes>();
                Enumerable.Range(0, Regex.Split(lens[0], format.Delimiter.ToString(), RegexOptions.IgnoreCase).Length).ToList().ForEach(x => allTypes.Add(eDataTypes.String));
                format.DataTypes = allTypes.ToArray();
            }

            int rowHead = row;
            //StringBuilder txtAll = new StringBuilder();
            foreach (string len in lens)
            {
                if (len.StartsWith("\""))
                {
                    string temp = len.Replace("\",\"", "\0");
                    temp = temp.Substring(0, temp.Length - 1).Substring(1, temp.Length - 2);
                    //txtAll.AppendLine(temp);
                    sheet.Cells[row++, col].LoadFromText(temp, format);
                }
                else
                    //txtAll.AppendLine(len);
                    sheet.Cells[row++, col].LoadFromText(len, format);
            }

            //sheet.Cells[row, col].LoadFromText(txtAll.ToString(), format);

            if (lens.Count > 0)
            {
                sheet.Cells[rowHead, col, rowHead, sheet.GetMaxColumn(rowHead)].SetRangeColor(Color.FromArgb(91, 155, 213));
                sheet.Cells[rowHead, col, rowHead, sheet.GetMaxColumn(rowHead)].SetRangeBorder(Color.DarkGray);
                sheet.Cells[rowHead, col, rowHead, sheet.GetMaxColumn(rowHead)].SetFontColor(Color.White);
            }

            if (lens.Count() <= 0) return 0;
            if (lens.Count() == 1 && string.IsNullOrEmpty(lens[0])) return 0;

            int cnt = lens.Count;
            lens = null;

            return cnt;
        }

        public static int LoadDataText(this ExcelWorksheet sheet, string file, int row = 1, int col = 1)
        {
            string[] txtLens = File.ReadAllLines(file, Encoding.GetEncoding(932));

            foreach (string len in txtLens)
                sheet.Cells[row++, col].Value = len;

            return txtLens.Count();
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

        #endregion

        #region NPOI

        public static void CreateRowCells(this ISheet sheet, int colCnt)
        {
            IRow row = sheet.CreateRow(sheet.LastRowNum + 1);
            for (int i = 0; i < colCnt; i++)
                row.CreateCell(i);

        }

        #endregion

        public static bool RefreshPivotTable(string path)
        {
            Application _app = null;
            Workbooks _wkbs = null;
            Workbook _wkb = null;

            try
            {
                _app = new Application();
                _app.Visible = false;
                _app.Application.ScreenUpdating = false;
                _app.Application.DisplayAlerts = false;
                LogUtility.WriteInfo($"【透视表】透视表刷新App");
                _wkbs = _app.Workbooks;
                _wkbs.Application.ScreenUpdating = false;
                _wkbs.Application.DisplayAlerts = false;
                LogUtility.WriteInfo($"【透视表】透视表刷新Workbooks");
                _wkb = _wkbs.Open(path);
                _wkb.Application.ScreenUpdating = false;
                _wkb.Application.CalculateBeforeSave = false;
                _wkb.Application.DisplayAlerts = false;

                LogUtility.WriteInfo($"【透视表】透视表刷新开始");
                foreach (Worksheet sheet in _wkb.Worksheets)
                {
                    foreach (PivotTable pivotTable in sheet.PivotTables())
                    {
                        pivotTable.PivotCache().Refresh();

                        //Marshal.ReleaseComObject(pivotTable);
                        LogUtility.WriteInfo($"【透视表】透视表刷新完了{pivotTable.Name}");
                    }
                }
                LogUtility.WriteInfo($"【透视表】透视表刷新结束");
                Thread.Sleep(1000);

                _wkb.Close(SaveChanges: true);
                _wkbs.Close();
                _app.Quit();

                return true;
            }
            catch (NullReferenceException)
            {
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(_wkb);
                Marshal.ReleaseComObject(_wkbs);
                Marshal.ReleaseComObject(_app);
            }
        }
    }
}
