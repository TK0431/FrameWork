namespace FrameWork.Models
{
    public class SqlServerModel
    {
        /// <summary>
        /// データベース
        /// </summary>
        public string Db { get; set; }

        /// <summary>
        /// サーバー：Ip
        /// </summary>
        public string Server { get; set; }

        /// <summary>
        /// DB登録User
        /// </summary>
        public string User { get; set; }

        /// <summary>
        /// DBパスワード
        /// </summary>
        public string PassWord { get; set; }
    }
}
