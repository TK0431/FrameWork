using FrameWork.Consts;
using Microsoft.VisualBasic;
using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace FrameWork.Utility
{
    public static class CheckUtil
    {
        /// <summary>
        /// ASCII
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsAscii(string value)
        {
            if (string.IsNullOrEmpty(value)) return true;

            return Regex.IsMatch(value, @"^[\u0000-\u00FF]*$");
        }

        /// <summary>
        /// 半角カナ
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsHanKana(string value)
        {
            if (string.IsNullOrEmpty(value)) return true;

            return Regex.IsMatch(value, @"^[\uFF61-\uFF9F]*$");
        }

        /// <summary>
        /// 半角カナ存在判断
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsHaveHanKana(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;

            return Regex.IsMatch(value, @".*[\uFF61-\uFF9F].*");
        }

        /// <summary>
        /// 半角カナを全角カナに変換
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string ToZenkakuKana(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            return Regex.Replace(value, @"[\uFF61-\uFF9F]+", new MatchEvaluator((Match m) => Strings.StrConv(m.Value, VbStrConv.Wide)));
        }

        /// <summary>
        /// 半角英数カナを全角英数カナに変換
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string ToZenkaku(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            return Regex.Replace(value, @"[0-9a-zA-Z\uFF61-\uFF9F]+", new MatchEvaluator((Match m) => Strings.StrConv(m.Value, VbStrConv.Wide)));
        }

        /// <summary>
        /// 文字数が規定値を超えている判定
        /// </summary>
        /// <param name="value"></param>
        /// <param name="maxLengthCnt"></param>
        /// <returns></returns>
        public static bool IsLengthOver(string value, int maxLengthCnt)
        {
            if (string.IsNullOrEmpty(value)) return false;

            return value.Length > maxLengthCnt ? true : false;
        }

        /// <summary>
        /// 半角数字判断
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsHanNum(string value)
        {
            if (string.IsNullOrEmpty(value)) return true;

            return Regex.IsMatch(value, @"^[0-9]+$");
        }

        /// <summary>
        /// 半角英数字
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsHanEngNum(string value)
        {
            if (string.IsNullOrEmpty(value)) return true;

            return Regex.IsMatch(value, @"^[a-zA-Z0-9]+$");
        }

        /// <summary>
        /// 年月日形式判断
        /// </summary>
        /// <param name="value"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static bool IsDateFormat(string value, string format = ComConst.DATA_YYYY_MM_DD)
        {
            if (string.IsNullOrEmpty(value)) return false;

            DateTime temp;

            return DateTime.TryParseExact(value, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out temp);
        }

        /// <summary>
        /// 大小関係判断
        /// </summary>
        /// <param name="valueFrom"></param>
        /// <param name="valueTo"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static bool CorrectRange(string valueFrom, string valueTo, string format = ComConst.DATA_YYYY_MM_DD)
        {
            if (string.IsNullOrEmpty(valueFrom) || string.IsNullOrEmpty(valueTo)) return true;

            DateTime dateFrom;
            DateTime dateTo;
            if (DateTime.TryParseExact(valueFrom, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateFrom)
                && DateTime.TryParseExact(valueTo, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTo))
            {
                if (dateFrom <= dateTo) return true; else return false;
            }

            Decimal decFrom;
            Decimal decTo;
            if (Decimal.TryParse(valueFrom, out decFrom) && Decimal.TryParse(valueTo, out decTo))
            {
                if (decFrom <= decTo) return true; else return false;
            }

            if (string.Compare(valueFrom, valueTo) == 1) return false; else return true;
        }
    }
}
