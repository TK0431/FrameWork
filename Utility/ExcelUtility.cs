using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;
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
        public ExcelUtility(string excelName, string passWord = null)
        {
            // License
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Excel
            if (string.IsNullOrWhiteSpace(passWord))
                _excel = new ExcelPackage(new FileInfo(excelName));
            else
                try
                {
                    _excel = new ExcelPackage(new FileInfo(excelName), passWord);
                }
                catch (SecurityException)
                {
                    throw new Exception("Err:密码不正确");
                }
                catch (InvalidDataException)
                {
                    _excel = new ExcelPackage(new FileInfo(excelName));
                }
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

    }
}
