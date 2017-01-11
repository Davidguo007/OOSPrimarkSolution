using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;

namespace OOSCommon
{
    public class Utils
    {
        public static readonly string EN = "en-US";
        private static Object GetAttributeClass(Enum enumSubitem, Type attributeType)
        {
            FieldInfo fieldinfo = enumSubitem.GetType().GetField(enumSubitem.ToString());
            Object[] objs = fieldinfo.GetCustomAttributes(attributeType, false);
            if (objs == null || objs.Length == 0)
            {
                return null;
            }
            return objs[0];
        }

        /// <summary>
        /// 获得某个Enum类型的绑定列表
        /// </summary>
        /// <param name="enumType">枚举的类型，例如：typeof(SysStatus)</param>
        /// <returns>
        /// 返回一个DataTable
        /// DataTable 有两列：    "StatusText"    : System.String;
        ///                        "StatusValue"    : System.Char
        /// </returns>
        public static DataTable EnumListTable(string Culture, Type enumType)
        {
            if (enumType.IsEnum != true)
            {    //不是枚举的要报错
                throw new InvalidOperationException();
            }
            //建立DataTable的列信息
            DataTable dt = new DataTable();
            dt.Columns.Add("StatusText", typeof(System.String));
            dt.Columns.Add("StatusValue", typeof(System.Int32));

            //获得特性Description的类型信息
            Type typeDescription = typeof(DescriptionAttribute);
            Type typeselfAttribute = typeof(selfAttribute);

            //获得枚举的字段信息（因为枚举的值实际上是一个static的字段的值）
            System.Reflection.FieldInfo[] fields = enumType.GetFields();

            //检索所有字段
            foreach (FieldInfo field in fields)
            {
                //过滤掉一个不是枚举值的，记录的是枚举的源类型
                if (field.FieldType.IsEnum)
                {
                    DataRow dr = dt.NewRow();

                    // 通过字段的名字得到枚举的值
                    dr["StatusValue"] = (int)enumType.InvokeMember(field.Name, BindingFlags.GetField, null, null, null);
                    if (string.IsNullOrEmpty(Culture))
                    {
                        //获得这个字段的所有自定义特性，这里只查找Description特性
                        object[] arr = field.GetCustomAttributes(typeDescription, true);
                        if (arr.Length > 0)
                        {
                            //因为Description这个自定义特性是不允许重复的，所以我们只取第一个就可以了！
                            DescriptionAttribute aa = (DescriptionAttribute)arr[0];
                            //获得特性的描述值，也就是中文描述
                            dr["StatusText"] = aa.Description;

                        }
                        else
                        {
                            //如果没有特性描述
                            dr["StatusText"] = field.Name;
                        }
                    }
                    else//has culture
                    {

                        //获得这个字段的所有自定义特性，这里只查找Description特性
                        object[] arr = field.GetCustomAttributes(typeselfAttribute, true);
                        if (arr.Length > 0)
                        {
                            //因为Description这个自定义特性是不允许重复的，所以我们只取第一个就可以了！
                            selfAttribute aa = (selfAttribute)arr[0];
                            //获得特性的描述值，也就是中文描述
                            if (Culture == EN || Culture.Contains("En"))
                            {
                                dr["StatusText"] = aa.ENDisplayText;
                            }
                            else
                            {
                                dr["StatusText"] = aa.CNDisplayText;
                            }

                        }
                        else
                        {
                            //如果没有特性描述
                            dr["StatusText"] = field.Name;
                        }

                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        /// <summary>
        /// 获得某个Enum类型的绑定列表
        /// </summary>
        /// <param name="enumType">枚举的类型，例如：typeof(SysStatus)</param>
        /// <returns>
        /// 返回一个DataTable
        /// DataTable 有两列：    "StatusText"    : System.String;
        ///                        "StatusValue"    : System.Char
        /// </returns>
        public static DataTable EnumListTable(Type enumType)
        {
            if (enumType.IsEnum != true)
            {    //不是枚举的要报错
                throw new InvalidOperationException();
            }
            //建立DataTable的列信息
            DataTable dt = new DataTable();
            dt.Columns.Add("StatusText", typeof(System.String));
            dt.Columns.Add("StatusValue", typeof(System.Int32));

            //获得特性Description的类型信息
            Type typeDescription = typeof(DescriptionAttribute);

            //获得枚举的字段信息（因为枚举的值实际上是一个static的字段的值）
            System.Reflection.FieldInfo[] fields = enumType.GetFields();

            //检索所有字段
            foreach (FieldInfo field in fields)
            {
                //过滤掉一个不是枚举值的，记录的是枚举的源类型
                if (field.FieldType.IsEnum)
                {
                    DataRow dr = dt.NewRow();

                    // 通过字段的名字得到枚举的值
                    dr["StatusValue"] = (int)enumType.InvokeMember(field.Name, BindingFlags.GetField, null, null, null);

                    //获得这个字段的所有自定义特性，这里只查找Description特性
                    object[] arr = field.GetCustomAttributes(typeDescription, true);
                    if (arr.Length > 0)
                    {
                        //因为Description这个自定义特性是不允许重复的，所以我们只取第一个就可以了！
                        DescriptionAttribute aa = (DescriptionAttribute)arr[0];
                        //获得特性的描述值，也就是中文描述
                        dr["StatusText"] = aa.Description;

                    }
                    else
                    {
                        //如果没有特性描述
                        dr["StatusText"] = field.Name;
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        /// <summary>
        /// 获得某个Enum类型的绑定列表
        /// </summary>
        /// <param name="enumType">枚举的类型，例如：typeof(SysStatus)</param>
        /// <returns>
        /// 返回一个string  
        /// 有两列：    "StatusText"    : System.String;
        ///                "StatusValue"    : System.Char
        /// </returns>
        public static string GetEnumListDesc(Type enumType, object status_or_kindobj)
        {
            if (Convert.IsDBNull(status_or_kindobj)) return string.Empty;
            int status_or_kind = Int32.Parse(status_or_kindobj.ToString());
            if (enumType.IsEnum != true)
            {    //不是枚举的要报错
                throw new InvalidOperationException();
            }




            //获得特性Description的类型信息
            Type typeDescription = typeof(DescriptionAttribute);

            //获得枚举的字段信息（因为枚举的值实际上是一个static的字段的值）
            System.Reflection.FieldInfo[] fields = enumType.GetFields();

            //检索所有字段
            foreach (FieldInfo field in fields)
            {
                //过滤掉一个不是枚举值的，记录的是枚举的源类型
                if (field.FieldType.IsEnum)
                {


                    // 通过字段的名字得到枚举的值             


                    if ((int)enumType.InvokeMember(field.Name, BindingFlags.GetField, null, null, null) == status_or_kind)
                    {
                        string desc = string.Empty;
                        //获得这个字段的所有自定义特性，这里只查找Description特性
                        object[] arr = field.GetCustomAttributes(typeDescription, true);
                        if (arr.Length > 0)
                        {
                            //因为Description这个自定义特性是不允许重复的，所以我们只取第一个就可以了！
                            DescriptionAttribute aa = (DescriptionAttribute)arr[0];
                            //获得特性的描述值，也就是中文描述
                            //dr["StatusText"] = aa.Description;
                            desc = aa.Description;
                        }
                        else
                        {
                            //如果没有特性描述
                            // dr["StatusText"] = field.Name;
                            desc = field.Name;
                        }
                        return desc;

                    }
                }
            }
            return string.Empty;
        }


        /// <summary>
        /// 获得某个Enum类型的绑定列表
        /// </summary>
        /// <param name="enumType">枚举的类型，例如：typeof(SysStatus)</param>
        /// <returns>
        /// 返回一个string  
        /// 有两列：    "StatusText"    : System.String;
        ///                "StatusValue"    : System.Char
        /// </returns>
        public static string GetEnumListDesc(string Culture, Type enumType, object status_or_kindobj)
        {
            if (Convert.IsDBNull(status_or_kindobj)) return string.Empty;
            int status_or_kind = Int32.Parse(status_or_kindobj.ToString());
            if (enumType.IsEnum != true)
            {    //不是枚举的要报错
                throw new InvalidOperationException();
            }




            //获得特性Description的类型信息
            Type typeDescription = typeof(DescriptionAttribute);
            Type typeselfAttribute = typeof(selfAttribute);

            //获得枚举的字段信息（因为枚举的值实际上是一个static的字段的值）
            System.Reflection.FieldInfo[] fields = enumType.GetFields();

            //检索所有字段
            foreach (FieldInfo field in fields)
            {
                //过滤掉一个不是枚举值的，记录的是枚举的源类型
                if (field.FieldType.IsEnum)
                {


                    // 通过字段的名字得到枚举的值             

                    string desc = string.Empty;
                    if ((int)enumType.InvokeMember(field.Name, BindingFlags.GetField, null, null, null) == status_or_kind)
                    {


                        if (string.IsNullOrEmpty(Culture))
                        {
                            //获得这个字段的所有自定义特性，这里只查找Description特性
                            object[] arr = field.GetCustomAttributes(typeDescription, true);
                            if (arr.Length > 0)
                            {
                                //因为Description这个自定义特性是不允许重复的，所以我们只取第一个就可以了！
                                DescriptionAttribute aa = (DescriptionAttribute)arr[0];
                                //获得特性的描述值，也就是中文描述
                                //dr["StatusText"] = aa.Description;
                                desc = aa.Description;
                            }
                            else
                            {
                                //如果没有特性描述
                                // dr["StatusText"] = field.Name;
                                desc = field.Name;
                            }
                            return desc;

                        }
                        else// find
                        {

                            //获得这个字段的所有自定义特性，这里只查找Description特性
                            object[] arr = field.GetCustomAttributes(typeselfAttribute, true);
                            if (arr.Length > 0)
                            {
                                //因为Description这个自定义特性是不允许重复的，所以我们只取第一个就可以了！
                                selfAttribute aa = (selfAttribute)arr[0];
                                //获得特性的描述值，也就是中文描述
                                if (Culture == EN || Culture.Contains("En"))
                                {

                                    desc = aa.ENDisplayText;
                                }
                                else
                                {
                                    desc = aa.CNDisplayText;
                                }
                                return desc;

                            }
                            else
                            {
                                //如果没有特性描述

                                return field.Name;
                            }

                        }
                    }

                }
            }
            return string.Empty;
        }


        //System.Net.HttpWebRequest
        public static bool GetSrvPdfFile(string ftppath, string localfilepath, string filename, ref string _newFileName, string username, string pass, ref StringWriter sw, ref string ErrDsc)
        {
            ErrDsc = string.Empty;
            // string ftppath = @"ftp://192.168.30.33/20100515/";
            // string localfilepath = @"c:\log\TestBtw\";
            // string filename = @"3.btw";
            // string username = @"mscrmdev01\administrator";
            // string pass = @"-0p9o8i7u";

            string FtpFilePathName = ftppath + filename;
            string _localfileName = filename;
            if (filename.Contains("/"))
            {
                _localfileName = filename.Split('/')[1];
            }

            _localfileName = _localfileName.Replace(".PDF", "").Replace(".pdf", "") + "-" + DateTime.Now.ToString("MM-dd HH:mm:ss").Replace("-", "").Replace(":", "").Replace(" ", "") + ".pdf";
            _newFileName = _localfileName;

            string LocalFilePathName = localfilepath + _localfileName;

            if (File.Exists(LocalFilePathName))
            {
                FileInfo fi = new FileInfo(LocalFilePathName);
                if (fi.Length <= 1)
                {
                    File.Copy(LocalFilePathName, LocalFilePathName.Replace(".pdf", "-GetSrvPKFileRename" + DateTime.Now.ToFileTime() + ".pdf"));
                    File.Delete(LocalFilePathName);
                }
                else
                {
                    return true;
                }
            }

            if (!System.IO.Directory.Exists(localfilepath))
                System.IO.Directory.CreateDirectory(localfilepath);

            sw.WriteLine();
            //sw.WriteLine("Ftp is: " + FtpFilePathName);

            sw.WriteLine("Start Down Load File [" + ftppath + "], the server file [" + FtpFilePathName + "]. " + System.DateTime.Now.ToString());
            HttpWebRequest webRequest;

            try
            {

                FileStream outputStream = new FileStream(LocalFilePathName, FileMode.Create);

                webRequest = (HttpWebRequest)WebRequest.Create(new Uri(FtpFilePathName));

                //webRequest.Credentials = new NetworkCredential(username, pass);

                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

                Stream httpStream = webResponse.GetResponseStream();

                int bufferSize = 2048;

                int readCount;

                byte[] buffer = new byte[bufferSize];

                try
                {

                    readCount = httpStream.Read(buffer, 0, bufferSize);

                    while (readCount > 0)
                    {
                        outputStream.Write(buffer, 0, readCount);

                        readCount = httpStream.Read(buffer, 0, bufferSize);
                    }

                    httpStream.Close();

                    outputStream.Close();

                    webResponse.Close();
                }
                catch (Exception ex)
                {
                    sw.WriteLine();
                    sw.WriteLine("Exception:" + ex.Message);
                }

                sw.WriteLine("OK Down Load File to local. the file is : [" + LocalFilePathName + "]. " + System.DateTime.Now.ToString());
                return true;
            }
            catch (Exception ex)
            {
                ErrDsc = ex.Message;
                sw.WriteLine("Error Down Load File to local. the file is : [" + LocalFilePathName + "]. " + System.DateTime.Now.ToString());
                sw.WriteLine("ex.Message is:  " + ex.Message + ex.StackTrace);
                return false;
            }
        }

        public static List<string> ReadDatFile(string filename, ref StringWriter sw)
        {
            try
            {
                List<string> list = new List<string>();

                StreamReader sr = new StreamReader(filename, Encoding.Default);
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    list.Add(line.ToString());
                }

                sr.Close();
                sr.Dispose();

                return list;

            }
            catch (Exception ex)
            {
                sw.WriteLine(ex.Message);
                sw.WriteLine(ex.StackTrace);
                return null;
            }            
        }
    }
}
