using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Utility
{
    public static class DbDataConvert
    {
        /// <summary>
        /// NULL判断
        /// </summary>
        /// <param name="value">判断値</param>
        /// <returns></returns>
        public static bool IsDBNull(object value) => value == DBNull.Value ? true : false;

        /// <summary>
        /// stringへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static string CStrDB(object value, string failValue = "") => (value == null || IsDBNull(value) || string.IsNullOrEmpty(value.ToString().Trim())) ? failValue : value.ToString().Trim();

        /// <summary>
        /// boolへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static bool CBoolDB(object value, bool failValue = false)
        {
            // 異常変更の場合
            if ((value == null || IsDBNull(value))) return failValue;

            // stringへの変更
            string result = value.ToString().Trim();
            return (result.ToUpper() == "FALSE" || result == "0" || result == string.Empty) ? false : true;
        }

        /// <summary>
        /// Decimalへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static Decimal CDecDB(object value, Decimal failValue = 0)
        {
            // 異常変更の場合
            if ((value == null || IsDBNull(value))) return failValue;

            // Decimalへの変更
            Decimal result;
            return Decimal.TryParse(value.ToString().Trim(), out result) ? result : failValue;
        }

        /// <summary>
        /// Decimalへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static Decimal? CDecDB2(object value, Decimal? failValue = null)
        {
            // 異常変更の場合
            if ((value == null || IsDBNull(value))) return failValue;

            string value2 = value.ToString();

            if (value2.Contains("%"))
            {
                Decimal? result = (Decimal?)(int.Parse(value2.Replace("%", "")) / 100.0);
                return result;
            }
            else
            {
                // Decimalへの変更
                Decimal result;
                return (Decimal.TryParse(value.ToString().Trim(), out result) ? result : failValue) as Decimal?;
            }
        }

        /// <summary>
        /// Decimalへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static int CIntDB(object value, int failValue = 0)
        {
            // 異常変更の場合
            if ((value == null || IsDBNull(value))) return failValue;

            // Decimalへの変更
            int result;
            return int.TryParse(value.ToString().Trim(), out result) ? result : failValue;
        }

        /// <summary>
        /// DateTimeへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static DateTime? CDateDB(object value, DateTime? failValue = null)
        {
            // 異常変更の場合
            if ((value == null || IsDBNull(value)) || string.IsNullOrWhiteSpace(value.ToString().Trim())) return failValue;

            // DateTimeへの変更
            DateTime? resultDate = DateTime.Parse(value.ToString()) as DateTime?;
            return resultDate == null ? failValue : resultDate;
        }

        /// <summary>
        /// DateTimeからStringへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="type">変更タイプ</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static string CDateToStrDB(object value, string type = "yyyy/MM/dd hh:mm:ss.fff", string failValue = "")
        {
            // 異常変更Falseへ戻る
            if ((value == null || IsDBNull(value))) return failValue;

            // DateTimeへの変更
            string result = value.ToString().Trim();
            DateTime resultDate;
            if (!DateTime.TryParse(result, out resultDate)) return failValue;

            // Stringへの変更
            return resultDate.ToString(type);
        }
    }
}
