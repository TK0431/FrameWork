using FrameWork.Consts;
using FrameWork.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;

namespace FrameWork.Utility
{
    /// <summary>
    /// EnumUtil
    /// </summary>
    public static class EnumUtil
    {
        /// <summary>
        /// 获取消息
        /// </summary>
        /// <param name="code"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static string GetMessage(this EnumMessage code, params string[] para)
        =>  para.Length > 0 ? string.Format(code.GetDescription(), para) : code.GetDescription();

        /// <summary>
        /// 枚举Lsit信息获取
        /// </summary>
        /// <param name="enumType"></param>
        /// <param name="flgRemoveAll">如果不想要All,设置为true</param>
        /// <returns></returns>
        public static List<EnumItem> GetList(this Type enumType, bool flgRemoveAll = false, EnumLanguage lan = EnumLanguage.CN)
        {
            // 返回值
            List<EnumItem> result = new List<EnumItem>();

            // 枚举遍历
            foreach (var e in Enum.GetValues(enumType))
            {
                EnumItem item = new EnumItem();

                // 获取Index
                item.Index = (int)e;
                // 是否加入All枚举
                if (flgRemoveAll && item.Index == 0) continue;

                // 获取描述
                object[] objArr;
                switch (lan)
                {
                    // 中文
                    case EnumLanguage.CN:
                        objArr = e.GetType().GetField(e.ToString()).GetCustomAttributes(typeof(DescriptionAttribute), true);
                        if (objArr != null && objArr.Length > 0)
                        {
                            // 中文获取Description特性
                            DescriptionAttribute da = objArr[0] as DescriptionAttribute;
                            item.Description = da.Description;
                        }
                        break;
                    // 英文
                    case EnumLanguage.EN:
                        objArr = e.GetType().GetField(e.ToString()).GetCustomAttributes(typeof(EnglishAttribute), true);
                        if (objArr != null && objArr.Length > 0)
                        {
                            // 英文获取Description特性
                            EnglishAttribute da = objArr[0] as EnglishAttribute;
                            item.Description = da.Value;
                        }
                        break;
                    // 日文
                    default:
                        objArr = e.GetType().GetField(e.ToString()).GetCustomAttributes(typeof(JapaneseAttribute), true);
                        if (objArr != null && objArr.Length > 0)
                        {
                            // 日文获取Japanese特性
                            JapaneseAttribute da = objArr[0] as JapaneseAttribute;
                            item.Description = da.Value;
                        }
                        break;
                }

                // 获取Value值
                object[] dbValueArr = e.GetType().GetField(e.ToString()).GetCustomAttributes(typeof(ValueAttribute), true);
                if (dbValueArr != null && dbValueArr.Length > 0)
                {
                    ValueAttribute da = dbValueArr[0] as ValueAttribute;
                    item.Value = da.Value;
                }

                result.Add(item);
            }
            return result;
        }

        /// <summary>
        /// EnumのDescriptionを取得
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="code"></param>
        /// <param name="para"></param>
        /// <returns></returns>
        //public static List<EnumItem> GetValues(this Type enumType)
        //{
        //    List<EnumItem> result = new List<EnumItem>();

        //    foreach (var e in Enum.GetValues(enumType))
        //    {
        //        EnumItem item = new EnumItem();

        //        item.Index = (int)e;

        //        // DBValue
        //        object[] dbValueArr = e.GetType().GetField(e.ToString()).GetCustomAttributes(typeof(ValueAttribute), true);
        //        if (dbValueArr != null && dbValueArr.Length > 0)
        //        {
        //            ValueAttribute da = dbValueArr[0] as ValueAttribute;
        //            item.Value = da.Value;
        //        }

        //        result.Add(item);
        //    }

        //    return result;
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="code"></param>
        /// <returns></returns>
        public static EnumItem GetItem<T>(this T code) where T : Enum
            => new EnumItem()
            {
                Index = Convert.ToInt32(code),
                Description = code.GetDescription(),
                Value = code.GetValue(),
            };

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="code"></param>
        /// <returns></returns>
        public static string GetValue<T>(this T code) where T : Enum
            => code.GetAttribute<T, ValueAttribute>();

        /// <summary>
        /// Value
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="code"></param>
        /// <param name="para"></param>
        /// <returns></returns>
        public static string GetAttribute<T, M>(this T code) where T : Enum where M : BaseAttribute
        {
            object[] titleValueArr = code.GetType().GetField(code.ToString()).GetCustomAttributes(typeof(M), true);
            if (titleValueArr != null && titleValueArr.Length > 0)
            {
                M da = titleValueArr[0] as M;
                return da.Value;
            }

            return null;
        }

        /// <summary>
        /// Description
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="code"></param>
        /// <param name="para"></param>
        /// <returns></returns>
        public static string GetDescription<T>(this T code, EnumLanguage lan = EnumLanguage.CN) where T : Enum
        {
            Type type = code.GetType();
            MemberInfo[] memberInfos = type.GetMember(code.ToString());
            if (memberInfos != null && memberInfos.Length > 0)
            {
                switch (lan)
                {
                    // 中文
                    case EnumLanguage.CN:
                        DescriptionAttribute[] attrs2 = memberInfos[0].GetCustomAttributes(typeof(DescriptionAttribute), false) as DescriptionAttribute[];
                        if (attrs2 != null && attrs2.Length > 0)
                        {
                            // 中文获取Description特性
                            return attrs2[0].Description;
                        }
                        break;
                    // 英文
                    case EnumLanguage.EN:
                        EnglishAttribute[] attrs0 = memberInfos[0].GetCustomAttributes(typeof(EnglishAttribute), false) as EnglishAttribute[];
                        if (attrs0 != null && attrs0.Length > 0)
                        {
                            // 英文获取English特性
                            return attrs0[0].Value;
                        }
                        break;
                    // 日文
                    default:
                        JapaneseAttribute[] attrs1 = memberInfos[0].GetCustomAttributes(typeof(JapaneseAttribute), false) as JapaneseAttribute[];
                        if (attrs1 != null && attrs1.Length > 0)
                        {
                            // 日文获取Japanese特性
                            return attrs1[0].Value;
                        }
                        break;
                }
            }
            return code.ToString();
        }
    }
}
