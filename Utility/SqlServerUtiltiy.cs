using FrameWork.Models;
using System;
using System.Data;
using System.Data.SqlClient;

namespace FrameWork.Utility
{
    public class SqlServerUtiltiy : IDisposable
    {
        #region 変数
        /// <summary>
        /// SqlConnection
        /// </summary>
        private SqlConnection _conn;

        /// <summary>
        /// SqlCommand
        /// </summary>
        private SqlCommand _command;

        /// <summary>
        /// SqlTransaction
        /// </summary>
        private SqlTransaction _trn;

        #endregion

        #region コンストラクタ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SqlServerUtiltiy(SqlServerModel model)
        {
            // DB接続
            OpenConnection(model);

            // SqlTransaction開始
            _trn = _conn.BeginTransaction();

            // SqlCommand新規
            _command = new SqlCommand()
            {
                Connection = _conn,
                Transaction = _trn
            };
        }

        #endregion

        #region DB接続関数

        /// <summary>
        /// DB接続
        /// </summary>
        private void OpenConnection(SqlServerModel model)
        {
            // DB接続文字列の取得
            string connectionString = $@"Server={model.Server};Database={model.Db};User ID={model.User};Password={model.PassWord};Trusted_Connection=False;";

            // DB接続
            _conn = new SqlConnection(connectionString);
            _conn.Open();
        }

        #endregion

        #region パラメータ関数

        /// <summary>
        /// 参数設定
        /// </summary>
        /// <param name="name">項目</param>
        /// <param name="value">値</param>
        public void AddParameters(string name, object value)
        {
            value = value ?? DBNull.Value;

            _command.Parameters.AddWithValue(name, value);

        }

        /// <summary>
        /// 参数削除
        /// </summary>
        private void ClearParameters()
        {
            _command.Parameters.Clear();
        }

        /// <summary>
        /// SQL実行(SELECT)
        /// </summary>
        /// <param name="sql">SQL文</param>
        /// <returns></returns>
        public SqlDataReader ExecuteReader(string sql)
        {
            // SQL文設定
            _command.CommandText = sql;

            // データ取得
            SqlDataReader result = null;

            result = _command.ExecuteReader();

            // 参数削除
            ClearParameters();

            return result;
        }

        /// <summary>
        /// SQL実行(SELECT)
        /// </summary>
        /// <param name="sql">SQL文</param>
        /// <returns></returns>
        public DataTable ExecuteTable(string sql)
        {
            // SQL文設定
            _command.CommandText = sql;

            // データ取得
            DataTable result = new DataTable();

            SqlDataAdapter adp = new SqlDataAdapter(_command);

            adp.Fill(result);

            // 参数削除
            ClearParameters();

            return result;
        }

        /// <summary>
        /// SQL実行(SELECT)
        /// </summary>
        /// <param name="sql">SQL文</param>
        /// <returns></returns>
        public object ExecuteScalar(string sql)
        {
            // SQL文設定
            _command.CommandText = sql;

            // データ取得
            object result = null;

            result = _command.ExecuteScalar();


            // 参数削除
            ClearParameters();

            return result;
        }

        /// <summary>
        /// SQL実行(INSERT、UPDATE、DELETE)
        /// </summary>
        /// <param name="sql">SQL文</param>
        /// <returns>変更レコード数</returns>
        public int ExecuteNonQuery(string sql)
        {
            // SQL文設定
            _command.CommandText = sql;

            // データ更新
            int result = 0;

            result = _command.ExecuteNonQuery();

            // 参数削除
            ClearParameters();

            return result;
        }

        #endregion

        #region クローズ

        /// <summary>
        /// クローズ
        /// </summary>
        public void Dispose()
        {
            _trn.Commit();

            if (_command != null)
            {
                _command.Dispose();
            }

            // DBクローズ
            if (_conn != null)
            {
                _conn.Close();
                _conn.Dispose();
                _conn = null;
            }
        }

        #endregion
    }
}
