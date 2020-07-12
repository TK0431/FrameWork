using log4net;
using System;

namespace FrameWork.Utility
{
    public class LogUtility
    {
        /// <summary>
        /// 异常
        /// </summary>
        /// <param name="e"></param>
        public static void WriteError(object text, Exception e)
        {
            ILog log = LogManager.GetLogger("Error");
            log.Error(text, e);
        }

        /// <summary>
        /// 测试
        /// </summary>
        /// <param name="text"></param>
        public static void WriteDebug(object text)
        {
            ILog log = LogManager.GetLogger("Debug");
            log.Debug(text);
        }

        /// <summary>
        /// 警告
        /// </summary>
        /// <param name="text"></param>
        public static void WarnInfo(object text)
        {
            ILog log = LogManager.GetLogger("Warn");
            log.Warn(text);
        }

        /// <summary>
        /// 正常信息
        /// </summary>
        /// <param name="text"></param>
        public static void WriteInfo(object text)
        {
            ILog log = LogManager.GetLogger("Info");
            log.Info(text);
        }
    }
}
