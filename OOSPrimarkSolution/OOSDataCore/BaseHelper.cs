using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace OOSDataCore
{
    /// <summary>
    /// 数据库操作基类
    /// </summary>
    public class BaseHelper : IDisposable
    {
        private AbsConnString _connString;
        private IDbConnection _dbconnection;
        /// <summary>
        /// 实现构造函数
        /// </summary>
        /// <param name="conn">必须是集成自AbsConnString的类</param>
        public BaseHelper(AbsConnString conn)
        {
            _connString = conn;
            _dbconnection = new SqlConnection(_connString.ConnectionString);
        }

        public void BaseHelperConnReset(AbsConnString conn)
        {
            _connString = conn;
            _dbconnection = new SqlConnection(_connString.ConnectionString);
        }

        private IDbCommand CreateCommand()
        {
            IDbCommand cmd = _dbconnection.CreateCommand();
            cmd.CommandTimeout = _connString.TimeOut;
            return cmd;
        }
        private void PrepareCommand(IDbCommand cmd, IDbTransaction trans, CommandType cmdType, string cmdText, IDbDataParameter[] cmdParms)
        {
            if (_dbconnection.State != ConnectionState.Open)
            {
                _dbconnection.Open();
            }
            cmd.CommandText = cmdText;
            if (trans != null)
            {
                cmd.Transaction = trans;
            }
            cmd.CommandType = cmdType;
            if (cmdParms != null && cmdParms.Length > 0)
            {
                foreach (IDbDataParameter parm in cmdParms)
                {
                    if (parm != null)
                    {
                        cmd.Parameters.Add(parm);
                    }
                }
            }
        }

        public int ExecuteNonQuery(string cmdText, params IDbDataParameter[] commandParameters)
        {
            return ExecuteNonQuery(CommandType.Text, cmdText, commandParameters);
        }
        public int ExecuteNonQuery(CommandType cmdType, string cmdText, params IDbDataParameter[] commandParameters)
        {
            IDbCommand cmd = CreateCommand();
            PrepareCommand(cmd, null, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        public IDataReader ExecuteReader(string cmdText, params IDbDataParameter[] commandParameters)
        {
            return ExecuteReader(CommandType.Text, cmdText, commandParameters);
        }
        public IDataReader ExecuteReader(CommandType cmdType, string cmdText, params IDbDataParameter[] commandParameters)
        {
            IDbCommand cmd = CreateCommand();
            PrepareCommand(cmd, null, cmdType, cmdText, commandParameters);
            IDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            cmd.Parameters.Clear();
            return rdr;
        }
        public object ExecuteScalar(string cmdText, params IDbDataParameter[] commandParameters)
        {
            return ExecuteScalar(CommandType.Text, cmdText, commandParameters);
        }

        public object ExecuteScalar(CommandType cmdType, string cmdText, params IDbDataParameter[] commandParameters)
        {
            IDbCommand cmd = CreateCommand();
            PrepareCommand(cmd, null, cmdType, cmdText, commandParameters);
            object obj = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return obj;
        }
        public DataTable ExecuteTable(CommandType cmdType, string cmdText, params IDbDataParameter[] commandParameters)
        {
            IDbCommand cmd = CreateCommand();
            PrepareCommand(cmd, null, cmdType, cmdText, commandParameters);
            DataSet st = new DataSet();
            IDbDataAdapter ap = new SqlDataAdapter();
            ap.SelectCommand = cmd;
            ap.Fill(st);
            cmd.Parameters.Clear();
            return st.Tables[0];
        }
        public DataSet ExecuteDataSet(CommandType cmdType, string cmdText, params IDbDataParameter[] commandParameters)
        {
            IDbCommand cmd = CreateCommand();
            PrepareCommand(cmd, null, cmdType, cmdText, commandParameters);
            DataSet st = new DataSet();
            IDbDataAdapter ap = new SqlDataAdapter();
            ap.SelectCommand = cmd;
            ap.Fill(st);
            cmd.Parameters.Clear();
            return st;
        }
        public DataTable ExecuteTable(string cmdText, params IDbDataParameter[] commandParameters)
        {
            return ExecuteTable(CommandType.Text, cmdText, commandParameters);
        }
        public object ExecuteScalarByTrans(CommandType cmdType, string cmdText, params IDbDataParameter[] commandParameters)
        {
            IDbCommand cmd = CreateCommand();
            if (_dbconnection.State != ConnectionState.Open)
            {
                _dbconnection.Open();
            }
            IDbTransaction trans = _dbconnection.BeginTransaction();
            PrepareCommand(cmd, trans, cmdType, cmdText, commandParameters);
            try
            {
                object val = cmd.ExecuteScalar();

                if (val != null)
                {
                    trans.Commit();
                    cmd.Parameters.Clear();
                    return val;
                }
                else
                {
                    trans.Rollback();
                }
            }
            catch
            {
                trans.Rollback();
            }
            return null;
        }
        public int ExecuteNonQueryByTrans(CommandType cmdType, string cmdText, params IDbDataParameter[] commandParameters)
        {
            IDbCommand cmd = CreateCommand();
            if (_dbconnection.State != ConnectionState.Open)
            {
                _dbconnection.Open();
            }
            //_dbconnection.Open();
            IDbTransaction trans = _dbconnection.BeginTransaction();
            PrepareCommand(cmd, trans, cmdType, cmdText, commandParameters);
            try
            {
                int val = cmd.ExecuteNonQuery();
                if (val > 0)
                {
                    trans.Commit();
                    cmd.Parameters.Clear();
                    return val;
                }
                else
                {
                    trans.Rollback();
                }
            }
            catch
            {
                trans.Rollback();
            }
            return -1;

        }
        public IDbDataParameter CreateParameter(string parameterName)
        {
            SqlParameter p = new SqlParameter();
            p.ParameterName = parameterName;
            return p;
        }

        public IDbDataParameter CreateParameter(string parameterName, object value)
        {
            if (value == null)
            {
                value = (object)DBNull.Value;
            }

            SqlParameter p = new SqlParameter(parameterName, value);
            return p;
        }
        public IDbDataParameter CreateParameter(string parameterName, SqlDbType dbType, int size, object value)
        {
            if (value == null)
            {
                value = (object)DBNull.Value;
            }
            SqlParameter p = new SqlParameter(parameterName, dbType, size);
            p.Value = value;
            return p;
        }
        public IDbDataParameter CreateParameter(string parameterName, SqlDbType dbType, object value)
        {
            if (value == null)
            {
                value = (object)DBNull.Value;
            }

            SqlParameter p = new SqlParameter(parameterName, dbType);
            p.Value = value;
            return p;
        }

        #region IDisposable And Close
        public void Close()
        {
            if ((_dbconnection != null) && (_dbconnection.State != ConnectionState.Closed))
            {
                _dbconnection.Close();
            }
        }
        public void Dispose()
        {
            Close();
            if (_dbconnection != null)
            {
                _dbconnection.Dispose();
            }
        }
        #endregion

        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">多条SQL语句</param>  
        public int ExecuteSqlTran(List<String> SQLStringList)
        {
            using (SqlConnection conn = new SqlConnection(_dbconnection.ConnectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                SqlTransaction tx = conn.BeginTransaction();
                cmd.Transaction = tx;
                try
                {
                    int count = 0;
                    for (int n = 0; n < SQLStringList.Count; n++)
                    {
                        string strsql = SQLStringList[n];
                        if (strsql.Trim().Length > 1)
                        {
                            cmd.CommandText = strsql;
                            count += cmd.ExecuteNonQuery();
                        }
                    }
                    tx.Commit();
                    return count;
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    //return 0;
                    throw new Exception(ex.Message, ex);
                }
            }
        }
    }
}
