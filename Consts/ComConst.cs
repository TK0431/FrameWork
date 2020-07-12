using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Consts
{
    public class ComConst
    {
        #region DB定数

        /// <summary>
        /// 戻る状態：正常
        /// </summary>
        public const string DB_RESULT_STATUS_SAFE = "0";

        /// <summary>
        /// 戻る状態：エラー(想定内エラー)
        /// </summary>
        public const string DB_RESULT_STATUS_ERR_SAFE = "1";

        /// <summary>
        /// 戻る状態：エラー(想定外エラー)
        /// </summary>
        public const string DB_RESULT_STATUS_ERR_EX = "9";

        #endregion

        #region 時間定数

        /// <summary>
        /// YYYYMMDD
        /// </summary>
        public const string DATA_YYYYMMDD = "yyyyMMdd";

        /// <summary>
        /// YYYY/MM/DD
        /// </summary>
        public const string DATA_YYYY_MM_DD = "yyyy/MM/dd";

        /// <summary>
        /// YYYY-MM-DD
        /// </summary>
        public const string DATA_YYYYHMMHDD = "yyyy-MM-dd";

        /// <summary>
        /// YYYY/MM/DD HH:MM:SS
        /// </summary>
        public const string DATA_YYYYMMDD_HHMMSS = "yyyy/MM/dd HH:mm:ss";

        /// <summary>
        /// YYYYMMDDHHMMSS
        /// </summary>
        public const string DATA_YYYYMMDDHHMMSS = "yyyyMMddHHmmss";

        #endregion
    }
}
