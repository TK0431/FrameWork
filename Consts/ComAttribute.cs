using System;

namespace FrameWork.Consts
{
    /// <summary>
    /// 基础特性
    /// </summary>
    public class BaseAttribute : Attribute
    {
        /// <summary>
        /// DB中の値
        /// </summary>
        public string Value { get; set; }
    }

    /// <summary>
    /// 值特性
    /// </summary>
    public class ValueAttribute : BaseAttribute
    {
        public ValueAttribute(string value)
        {
            this.Value = value;
        }
    }

    /// <summary>
    /// 英语描述
    /// </summary>
    public class EnglishAttribute : BaseAttribute
    {
        public EnglishAttribute(string value)
        {
            this.Value = value;
        }
    }

    /// <summary>
    /// 日语描述
    /// </summary>
    public class JapaneseAttribute : BaseAttribute
    {
        public JapaneseAttribute(string value)
        {
            this.Value = value;
        }
    }
}
