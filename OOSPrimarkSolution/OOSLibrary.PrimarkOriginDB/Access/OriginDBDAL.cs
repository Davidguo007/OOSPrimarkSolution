using OOSDataCore;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOSLibrary.PrimarkOriginDB.Access
{
    /// <summary>
    /// 对象型数据库操作类
    /// </summary>
    public class OriginDBDAL : BaseHelper
    {
        private static DBConnString conn = new DBConnString();
        public OriginDBDAL()
            : base(conn)
        {

        }
        public void SetParameter(IDataParameter parameter, object value)
        {
            if (value == null)
            {
                value = (object)DBNull.Value;
            }
            parameter.Value = value;
        }

    }
}
