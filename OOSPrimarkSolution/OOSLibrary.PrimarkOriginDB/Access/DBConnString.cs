using OOSDataCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOSLibrary.PrimarkOriginDB.Access
{
    public class DBConnString : AbsConnString
    {
        public DBConnString()
        {

        }
        public static string _connectionstring;
        public override string ConnectionString
        {
            get
            {
                if (string.IsNullOrEmpty(_connectionstring))
                {
                    _connectionstring = System.Configuration.ConfigurationManager.ConnectionStrings["PrimarkOriginDBConnString"].ConnectionString;
                    //_connectionstring = HashEncrypt.DecryptForStr(_connectionstring);
                }
                return _connectionstring;
            }
        }
    }
}
