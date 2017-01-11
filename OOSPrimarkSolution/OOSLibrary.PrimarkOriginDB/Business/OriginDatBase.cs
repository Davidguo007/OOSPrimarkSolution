using OOSLibrary.PrimarkOriginDB.Access;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOSLibrary.PrimarkOriginDB.Business
{
    public class OriginDatBase
    {
        
        public OriginDatBase()
        {
            
        }

        /// <summary>
        /// 执行sql返回Datatable
        /// </summary>
        /// <param name="_sql">SQL语句</param>
        /// <returns>Datatable</returns>
        public DataTable GetDBDataBySQL(string _sql)
        {
            DataTable _dt = new DataTable();

            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                _dt = oosdbh.ExecuteTable(CommandType.Text, _sql);
            }

            return _dt;
        }

        /// <summary>
        /// 执行sql返回Datatable
        /// </summary>
        /// <param name="_sql">SQL语句</param>
        /// <param name="spara">SqlParameter 参数</param>
        /// <returns>DataTable</returns>
        public DataTable GetDBDataBySQL(string _sql, params SqlParameter[] spara)
        {
            DataTable _dt = new DataTable();

            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                _dt = oosdbh.ExecuteTable(CommandType.Text, _sql, spara);
            }

            return _dt;
        }

        /// <summary>
        /// 执行sql返回DataSet
        /// </summary>
        /// <param name="_sql">SQL语句</param>
        /// <returns>Datatable</returns>
        public DataSet GetDBDataSetBySQL(string _sql)
        {
            DataSet _ds = new DataSet();

            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                _ds = oosdbh.ExecuteDataSet(CommandType.Text, _sql);
            }

            return _ds;
        }

        /// <summary>
        /// 执行SQL,(查询单个值)
        /// </summary>
        /// <param name="_sql">sql语句</param>
        /// <returns>object</returns>
        public object GetDBDataScalarBySQL(string _sql)
        {
            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                object obj = oosdbh.ExecuteScalar(CommandType.Text, _sql);
                return obj;
            }
        }

        /// <summary>
        /// 执行SQL,(查询单个值)
        /// </summary>
        /// <param name="_sql">sql语句</param>
        /// <param name="spara">SqlParameter 参数</param>
        /// <returns>object</returns>
        public object GetDBDataScalarBySQL(string _sql, params SqlParameter[] spara)
        {
            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                object obj = oosdbh.ExecuteScalar(CommandType.Text, _sql, spara);
                return obj;
            }
        }

        /// <summary>
        /// [2]执行非查询语句 (增加，删除，修改)
        /// </summary>
        /// <param name="sql"> sql语句</param>
        /// <returns>int 受影响的行数</returns>
        public int NoQueryDataBySQL(string sql)
        {
            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                return oosdbh.ExecuteNonQuery(CommandType.Text, sql);
            }
        }

        /// <summary>
        /// [2]执行非查询语句 (增加，删除，修改)
        /// </summary>
        /// <param name="sql"> sql语句</param>
        /// <param name="spara">SqlParameter 参数</param>
        /// <returns>int，受影响的行数</returns>
        public int NoQueryDataBySQL(string sql, params SqlParameter[] spara)
        {
            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                return oosdbh.ExecuteNonQuery(CommandType.Text, sql, spara);
            }
        }

        public int NoQueryDataBySQLByTrans(string sql)
        {
            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                return oosdbh.ExecuteNonQueryByTrans(CommandType.Text, sql);
            }
        }

        /// <summary>
        /// 执行多语句的事务
        /// </summary>
        /// <param name="SQLStringList"></param>
        /// <returns></returns>
        public int ExecuteSqlTran(List<String> SQLStringList)
        {
            using (OriginDBDAL oosdbh = new OriginDBDAL())
            {
                return oosdbh.ExecuteSqlTran(SQLStringList);
            }
        }
    }
}
