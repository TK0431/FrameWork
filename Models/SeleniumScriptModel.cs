using System;
using System.Collections.Generic;

namespace FrameWork.Models
{
    public class SeleniumScriptModel
    {
        /// <summary>
        /// 超时(秒)
        /// </summary>
        public int OutTime { get; set; } = 20;

        /// <summary>
        /// 延时(毫秒)
        /// </summary>
        public int ReTry { get; set; } = 500;

        /// <summary>
        /// 导出路径
        /// </summary>
        public string OutPath { get; set; } = Environment.CurrentDirectory;

        /// <summary>
        /// 模板名
        /// </summary>
        public string TemplateFile { get; set; }

        /// <summary>
        /// 模板页
        /// </summary>
        public string TemplateSheet { get; set; }

        /// <summary>
        /// 图开始行
        /// </summary>
        public int StartRow { get; set; } = 1;

        /// <summary>
        /// 图开始列
        /// </summary>
        public int StartCol { get; set; } = 1;

        /// <summary>
        /// 替换参数
        /// </summary>
        public Dictionary<string, string> Args { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// WMenu Sheet
        /// </summary>
        public List<SeleniumOrder> Orders { get; set; } = new List<SeleniumOrder>();

        /// <summary>
        /// Events
        /// </summary>
        public Dictionary<string, Dictionary<string, List<SeleniumEvent>>> Events { get; set; } = new Dictionary<string, Dictionary<string, List<SeleniumEvent>>>();
    }

    /// <summary>
    /// WMenu Sheet
    /// </summary>
    public class SeleniumOrder
    {
        /// <summary>
        /// 导出文件
        /// </summary>
        public string File { get; set; }

        /// <summary>
        /// Case
        /// </summary>
        public string Case { get; set; }

        /// <summary>
        /// Sheet名
        /// </summary>
        public string Sid { get; set; }

        /// <summary>
        /// 备考
        /// </summary>
        public string Back { get; set; }

        /// <summary>
        /// 参数
        /// </summary>
        public Dictionary<string, string> Args { get; set; } = new Dictionary<string, string>();
    }

    public class SeleniumEvent
    {
        /// <summary>
        /// No
        /// </summary>
        public string No { get; set; }

        /// <summary>
        /// Key
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Cmd
        /// </summary>
        public string Cmd { get; set; }

        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// 框
        /// </summary>
        public string Range { get; set; }
    }
}
