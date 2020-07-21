using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Utility
{
    public static class ExceptionUtility
    {
        public static void LogStop(Exception ex,string msg)
        {
            LogUtility.WriteError(ex.Message, ex);
            throw new Exception($"Err:{msg}");
        }
    }
}
