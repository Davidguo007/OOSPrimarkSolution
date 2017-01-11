using OOSCommon;
using OOSLibrary.PrimarkOriginDB.Access;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOSLibrary.PrimarkOriginDB.Business
{
    public class PrimarkFTPDatAgent
    {

        public bool RecivePrimarkOrigFile(string OrigFilePath, string filename, ref StringWriter sw, ref string _ErrCode)
        {
            //ConfigLoad();
            //StringWriter sw = new StringWriter();
            string status = "True";
            sw.WriteLine("Start Time:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff"));
            //sw.WriteLine("客户端IP：" + HttpContext.Current.Request.UserHostAddress);
            try
            {
                string filepath = OrigFilePath + filename;
                sw.WriteLine("文件路径：" + filepath);
                if (!File.Exists(filepath))
                {
                    throw new Exception("File does not exist");
                }
                string filemd5 = HashEncrypt.GetFileMD5(filepath);
                sw.WriteLine("当前文件MD5值为：" + filemd5);

                FileInfo fi = new FileInfo(filepath);

                string FiledataLogID = Guid.NewGuid().ToString();
                sw.WriteLine("FiledataLogID:" + FiledataLogID);
                List<string> list = new List<string>();
                string sql_FiledataLog = string.Format("insert into OriginOrderHead(FileID, FilePath, FileName,FileSize_Bytes,tStaus) values('{0}','{1}','{2}','{3}','{4}')", FiledataLogID, filepath.Replace("'","''"), filename.Replace("'", "''"), fi.Length, 5);
                list.Add(sql_FiledataLog);

                List<string> listData = Utils.ReadDatFile(filepath, ref sw);
                if (list == null)
                {
                    throw new Exception("读取Dat文件失败");
                }

                DataSet ds = PrimarkOrig.DatToDataSet(listData, filename, ref sw);
                DataTable dt_primark = ds.Tables["OrderHeader"];
                sw.WriteLine("本次订单共" + dt_primark.Rows.Count + "笔明细");
                if (dt_primark.Rows.Count <= 0)
                {
                    _ErrCode = "no data";
                    return false;
                }

                string sqlinsert = "insert into OriginOrderDetail ({0}) values ({1})";
                string column = string.Empty;

                for (int i = 0; i < dt_primark.Columns.Count; i++)
                {
                    column += "," + dt_primark.Columns[i].ColumnName;

                }
                column = "FileID,Status" + column;

                for (int i = 0; i < dt_primark.Rows.Count; i++)
                {
                    string values = string.Empty;
                    for (int j = 0; j < ds.Tables["OrderHeader"].Columns.Count; j++)
                    {
                        values += ",N'" + ds.Tables["OrderHeader"].Rows[i][j].ToString().Replace("'", "''") + "'";
                    }
                    values = "'" + FiledataLogID + "',5" + values;

                    string sql = string.Format(sqlinsert, column, values);
                    list.Add(sql);
                }
                sw.WriteLine("sql新增语句组成完成，共" + list.Count + "笔");
                foreach (string item in list)
                {
                    sw.WriteLine(item);
                }

                OriginDatBase help = new OriginDatBase();
                string SqlErr = string.Empty;
                int sqlcount = help.ExecuteSqlTran(list);
                if (sqlcount != list.Count)
                {
                    sw.WriteLine("新增数据失败.");
                    sw.WriteLine(SqlErr);
                    throw new Exception("批量插入数据到DB失败。");
                }
                sw.WriteLine();
                sw.WriteLine("新增成功");
                return true;
            }
            catch (Exception ex)
            {
                sw.WriteLine(ex.Message);
                sw.WriteLine(ex.StackTrace);
                status = "False";
                return false;
            }
            //finally
            //{
            //    sw.WriteLine("End Time:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff"));
            //    LogPath += @"RecivePrimarkOrigFile\";
            //    string logfilename = string.Format("{1}-RecivePrimarkOrigFile-{0}.txt", DateTime.Now.ToFileTime(), status);
            //    LogAgent.WriteFileFunctionByTextEncoding(LogPath, logfilename, sw.ToString(), Encoding.UTF8);
            //}
        }
    }
}
