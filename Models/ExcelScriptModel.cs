using System.Collections.Generic;

namespace FrameWork.Models
{
    /// <summary>
    /// ExcelScriptModel
    /// </summary>
    public class ExcelScriptModel
    {
        /// <summary>
        /// Sheets
        /// </summary>
        public List<string> Sheets { get; set; } = new List<string>();

        /// <summary>
        /// Caseレコード
        /// </summary>
        public Dictionary<string, List<OrderItem>> OrderList { get; set; } = new Dictionary<string, List<OrderItem>>();

        /// <summary>
        /// TC Sheet Data
        /// </summary>
        public Dictionary<string, TCSheet> TCList { get; set; } = new Dictionary<string, TCSheet>();

        /// <summary>
        /// DB Sheet Data
        /// </summary>
        public Dictionary<string, DBSheet> DBList { get; set; } = new Dictionary<string, DBSheet>();
    }

    public class OrderItem
    {
        // [手順]No
        public string CaseNo { get; set; }
        // [手順]画面
        public string ViewName { get; set; }
        // [手順]Sheet
        public string Sheet { get; set; }
        // [手順]手順
        public string Order { get; set; }
    }

    public class TCSheet
    {
        // Control List
        public Dictionary<string, ControlInfo> ControlItems = new Dictionary<string, ControlInfo>();

        // Case のデータ
        public Dictionary<string, Dictionary<string,string>> CaseDatas { get; set; } = new Dictionary<string, Dictionary<string, string>>();
    }

    public class ControlInfo
    {
        // 番号
        public string Id { get; set; }
        // 画面項目論理名
        public string Name { get; set; }
        // 画面項目タイプ
        public string Class { get; set; }
        // 座標X
        public int X { get; set; }
        // 座標Y
        public int Y { get; set; }
    }

    public class DBSheet
    {
        // Case List
        public Dictionary<string, DBSheetInfo> CaseDatas { get; set; } = new Dictionary<string, DBSheetInfo>();
    }

    public class DBSheetInfo
    {
        // 番号
        public string Id { get; set; }
        // サーバ
        public string Server { get; set; }
        // データベース
        public string DataBase { get; set; }
        // ユーザ
        public string User { get; set; }
        // パスワード
        public string PassWord { get; set; }
        // テーブル名
        public string Table { get; set; }
        // タイプ
        public string IsDown { get; set; }
        // ファイル名
        public string FileName { get; set; }
        // SQL
        public string Sql { get; set; }
        public int Sleep { get; set; }
    }
}
