namespace FrameWork.Models
{
    /// <summary>
    /// 枚举类
    /// </summary>
    public class EnumItem
    {
        /// <summary>
        /// 枚举Index
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 描述（中日英）
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 值(和数据库相同)
        /// </summary>
        public string Value { get; set; }
    }
}
