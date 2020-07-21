using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Utility
{
    public class ExcelXlsUtility : IDisposable
    {
        private IWorkbook _excel;

        public ExcelXlsUtility(string excelName)
        {
            _excel = new HSSFWorkbook(new FileStream(excelName, FileMode.Open, FileAccess.Read));
        }

        public ExcelXlsUtility()
        {
            _excel = new HSSFWorkbook();
        }

        public void Dispose()
        {
            if (_excel != null) _excel.Close();
        }

        public ISheet GetSheet(string sheetName)
            => _excel.GetSheet(sheetName);

        public ISheet GetSheet(int index = 0)
            => _excel.GetSheetAt(index);

        public void SaveAs(string path)
        {
            _ = new FileInfo(path);
            _excel.Write(File.OpenWrite(path));
        }
    }
}
