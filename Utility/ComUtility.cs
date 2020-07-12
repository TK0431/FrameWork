using FrameWork.Consts;
using System;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace FrameWork.Utility
{
    public static class ComUtility
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static IPAddress GetLocalIPV4()
        {
            string name = Dns.GetHostName();
            IPAddress[] ipadrlist = Dns.GetHostAddresses(name);
            foreach (IPAddress ipa in ipadrlist)
            {
                if (ipa.AddressFamily == AddressFamily.InterNetwork)
                    return ipa;
            }
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="myString"></param>
        /// <returns></returns>
        public static string GetMD5(string myString)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] fromData = Encoding.UTF8.GetBytes(myString);
            byte[] targetData = md5.ComputeHash(fromData);
            string byte2String = null;

            for (int i = 0; i < targetData.Length; i++)
            {
                byte2String = byte2String + targetData[i].ToString("x2");
            }

            return byte2String;
        }

        public static void Map<S, T>(this T target, S source)
        {
            Type targetType = target.GetType();
            Type sourceType = source.GetType();

            foreach (PropertyInfo p in sourceType.GetProperties())
            {

                PropertyInfo targetP = targetType.GetProperty(p.Name);

                if (targetP != null)
                {
                    targetP.SetValue(target, p.GetValue(source));
                }
            }
        }
    }

    public static class DataUtility
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
        /// Intへの変更
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
            if ((value == null || IsDBNull(value))) return failValue;

            // DateTimeへの変更
            DateTime resultDate;
            if (DateTime.TryParse(value.ToString(), out resultDate))
                return resultDate as DateTime?;
            else
                return failValue;
        }

        /// <summary>
        /// DateTimeからStringへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="type">変更タイプ</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static string CDateToStrDB(object value, string type = ComConst.DATA_YYYY_MM_DD, string failValue = "")
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

        /// <summary>
        /// Intへの変更
        /// </summary>
        /// <param name="value">変更前値</param>
        /// <param name="failValue">変更失敗時の値</param>
        /// <returns></returns>
        public static string CIPDB(string value, string failValue = null)
        {
            // 異常変更の場合
            if ((value == null || IsDBNull(value))) return failValue;

            // Decimalへの変更
            return IPAddress.TryParse(value, out _) ? value : failValue;
        }
    }
}
