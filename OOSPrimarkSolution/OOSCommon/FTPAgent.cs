using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Collections;

namespace OOSCommon
{
    public class FTPAgent
    {
        private string path = "";//System.Configuration.ConfigurationSettings.AppSettings["FtpPath"];
        private string username = "";//ConfigurationManager.AppSettings["UserId"];
        private string password = "";//ConfigurationManager.AppSettings["UserPwd"];

        public FTPAgent(string host, string useid, string usePwd)
        {
            path = host;
            username = useid;
            password = usePwd;
        }

        /// <summary>
        /// 从ftp服务器上获得文件夹列表
        /// </summary>
        /// <param name="RequedstPath">服务器下的相对路径</param>
        /// <returns></returns>
        public List<string> GetDirctory(string RequedstPath)
        {
            List<string> strs = new List<string>();
            try
            {
                string uri = path + RequedstPath;   //目标路径 path为服务器地址
                FtpWebRequest reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
                // ftp用户名和密码
                reqFTP.Credentials = new NetworkCredential(username, password);
                reqFTP.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                WebResponse response = reqFTP.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream());//中文文件名

                string line = reader.ReadLine();
                while (line != null)
                {
                    if (line.Contains("<DIR>"))
                    {
                        string msg = line.Substring(line.LastIndexOf("<DIR>") + 5).Trim();
                        strs.Add(msg);
                    }
                    line = reader.ReadLine();
                }
                reader.Close();
                response.Close();
                return strs;
            }
            catch (Exception ex)
            {
                Console.WriteLine("获取目录出错：" + ex.Message);
            }
            return strs;
        }

        /// <summary>
        /// 判断当前目录下指定的子目录是否存在
        /// </summary>
        /// <param name="RemoteDirectoryName">指定的目录名</param>
        public bool DirectoryExist(string RemoteDirectoryName)
        {
            List<string> dirList = GetDirctory("");
            foreach (string str in dirList)
            {
                if (str.Trim() == RemoteDirectoryName.Trim())
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 获取当前目录下文件列表(仅文件)
        /// </summary>
        /// <returns></returns>
        public string[] GetFileList(string mask)
        {
            string _path = path + mask;
            //string[] downloadFiles;
            StringBuilder result = new StringBuilder();
            FtpWebRequest reqFTP;
            try
            {
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(_path));
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(username, password);
                reqFTP.Method = WebRequestMethods.Ftp.ListDirectory;
                WebResponse response = reqFTP.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);

                string line = reader.ReadLine();
                while (line != null)
                {
                    result.Append(line);
                    result.Append("\n");
                    line = reader.ReadLine();
                }
                result.Remove(result.ToString().LastIndexOf('\n'), 1);
                reader.Close();
                response.Close();
                return result.ToString().Split('\n');
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
                //   downloadFiles = null;
                //if (ex.Message.Trim() != "远程服务器返回错误: (550) 文件不可用(例如，未找到文件，无法访问文件)。")
                //{
                //    //.Insert("FtpWeb", "GetFileList Error --> " + ex.Message.ToString());
                //}
                // return downloadFiles;
            }
        }

        /// <summary>
        /// 判断指定目录下指定的文件是否存在
        /// </summary>
        /// <param name="direName">远程文件夹名</param>
        /// <param name="RemoteFileName">远程文件名</param
        public bool FileExist(string direName, string RemoteFileName)
        {
            string[] fileList = GetFileList(direName);
            foreach (string str in fileList)
            {
                if (str.Trim() == RemoteFileName.Trim())
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 从 FTP 上下载一个文件，
        /// </summary>
        /// <param name="ftppath">FTP目录：形如：ftp://192.168.1.113/MS/MS1207050006/</param>
        /// <param name="localfilepath">本地目标路径，形如：C:\log\FTP_TempFile\</param>
        /// <param name="filename">要从FTP上下载的文件名</param>
        /// <param name="username"></param>
        /// <param name="pass"></param>
        /// <param name="sw"></param>
        /// <param name="ErrDsc"></param>
        /// <returns></returns>
        public bool DownLoadFileFromFTP(string ftppath, string localfilepath, string filename, ref StringWriter sw, ref string ErrDsc)
        {

            // Path = @"ftp://192.168.1.113/";   localfilepath = @"C:\Log\";
            if (!ftppath.EndsWith("/"))
            {
                ftppath = ftppath + "/";
            }
            if (!localfilepath.EndsWith("\\"))
            {
                localfilepath = localfilepath + "\\";
            }

            ErrDsc = string.Empty;
            string FtpFilePathName = ftppath + filename;
            string LocalFilePathName = localfilepath + filename;

            if (!System.IO.Directory.Exists(localfilepath))
                System.IO.Directory.CreateDirectory(localfilepath);

            sw.WriteLine("DownLoadFileFromFTP  Start Down Load File " + LocalFilePathName + System.DateTime.Now.ToString());

            FtpWebRequest reqFTP;
            FileStream outputStream = new FileStream(LocalFilePathName, FileMode.Create);
            try
            {

                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(FtpFilePathName));
                reqFTP.Method = WebRequestMethods.Ftp.DownloadFile;
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(username, password);
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream ftpStream = response.GetResponseStream();
                bool _isDownSucc = true;
                try
                {
                    long cl = response.ContentLength;
                    int bufferSize = 2048;
                    int readCount;
                    byte[] buffer = new byte[bufferSize];
                    readCount = ftpStream.Read(buffer, 0, bufferSize);

                    while (readCount > 0)
                    {
                        outputStream.Write(buffer, 0, readCount);
                        readCount = ftpStream.Read(buffer, 0, bufferSize);
                    }
                }
                catch (Exception exiner)
                {
                    sw.WriteLine("Download Exception:" + exiner.Message);
                    _isDownSucc = false;
                }
                finally
                {
                    ftpStream.Close();
                    response.Close();
                }

                reqFTP.Abort();

                if (_isDownSucc)
                {
                    sw.WriteLine("DownLoadFileFromFTP Success Down Load File " + LocalFilePathName + System.DateTime.Now.ToString());
                    return true;
                }
                else
                {
                    sw.WriteLine("DownLoadFileFromFTP failed Down Load File " + LocalFilePathName + System.DateTime.Now.ToString());
                    return false;
                }
            }
            catch (Exception ex)
            {
                ErrDsc = ex.Message;
                sw.WriteLine("DownLoadFileFromFTP Exception Error Down Load File " + LocalFilePathName + System.DateTime.Now.ToString() + "  ex.Message " + ex.Message + ex.StackTrace);
                return false;
            }
            finally
            {                
                outputStream.Close();
            }
        }


        /// <summary>
        /// 文件移动
        /// </summary>
        /// <param name="currentFilename"></param>
        /// <param name="newFilename"></param>
        public void ReName(string currentFilename, string newFilename)
        {
            FtpWebRequest reqFTP;
            try
            {
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(path + currentFilename));
                reqFTP.Method = WebRequestMethods.Ftp.Rename;
                reqFTP.RenameTo = newFilename;
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(username, password);
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream ftpStream = response.GetResponseStream();

                ftpStream.Close();
                response.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
                //  Insert_Standard_ErrorLog.Insert("FtpWeb", "ReName Error --> " + ex.Message);
            }
        }

        /// <summary>
        /// 移动文件
        /// </summary>
        /// <param name="currentFilename"></param>
        /// <param name="newFilename"></param>
        public void MovieFile(string currentFilename, string newDirectory)
        {
            ReName(currentFilename, newDirectory);
        }
                
        /// <summary>
        /// Methods to upload file to FTP Server
        /// </summary>
        /// <param name="_FileName">local source file name，such as: E:\CRMWeb\WebDocuments\ArtWork\MS1207050006\abc.pdf</param>
        /// <param name="_UploadPath">Upload FTP path including Host name, such as:ftp://192.168.1.113/MS/MS1207050006/abc.pdf</param>
        /// <param name="_FTPUser">FTP login username</param>
        /// <param name="_FTPPass">FTP login password</param>
        public bool UploadFile(string _FileName, string _UploadPath, ref StringWriter sw, ref String ErrDsc)
        {
            try
            {
                sw.WriteLine("UploadFile  Start Time " + System.DateTime.Now.ToString());
                System.IO.FileInfo _FileInfo = new System.IO.FileInfo(_FileName);

                // Create FtpWebRequest object from the Uri provided
                System.Net.FtpWebRequest _FtpWebRequest = (System.Net.FtpWebRequest)System.Net.FtpWebRequest.Create(new Uri(_UploadPath));
                
                // Provide the WebPermission Credintials
                _FtpWebRequest.Credentials = new System.Net.NetworkCredential(username, password);

                // By default KeepAlive is true, where the control connection is not closed
                // after a command is executed.
                _FtpWebRequest.KeepAlive = false;
                // set timeout for 20 seconds
                _FtpWebRequest.Timeout = 20000;
                // Specify the command to be executed.
                _FtpWebRequest.Method = System.Net.WebRequestMethods.Ftp.UploadFile;
                // Specify the data transfer type.
                _FtpWebRequest.UseBinary = true;
                // Notify the server about the size of the uploaded file
                _FtpWebRequest.ContentLength = _FileInfo.Length;
                // The buffer size is set to 2kb
                int buffLength = 2048;
                byte[] buff = new byte[buffLength];
                // Opens a file stream (System.IO.FileStream) to read the file to be uploaded
                using (System.IO.FileStream _FileStream = _FileInfo.OpenRead())
                // Stream to which the file to be upload is written                
                using (System.IO.Stream _Stream = _FtpWebRequest.GetRequestStream())
                {
                    // Read from the file stream 2kb at a time
                    int contentLen = _FileStream.Read(buff, 0, buffLength);
                    // Till Stream content ends
                    while (contentLen != 0)
                    {
                        // Write Content from the file stream to the FTP Upload Stream
                        _Stream.Write(buff, 0, contentLen);
                        contentLen = _FileStream.Read(buff, 0, buffLength);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                sw.WriteLine(ex.StackTrace);
                ErrDsc = ex.Message;
                return false;
            }
            finally
            {
                sw.WriteLine("UploadFile  End Time " + System.DateTime.Now.ToString());
            }
        }
        
        public bool DelFile(string _FileName, string _UploadPath, ref StringWriter sw, ref string ErrDsc)
        {
            try
            {
                FileInfo fileInf = new FileInfo(_FileName);
                string uri = _UploadPath + fileInf.Name;

                sw.WriteLine("UploadFile  Start Time " + System.DateTime.Now.ToString());
                System.Net.FtpWebRequest _FtpWebRequest = (System.Net.FtpWebRequest)System.Net.FtpWebRequest.Create(new Uri(uri));
                _FtpWebRequest.Credentials = new System.Net.NetworkCredential(username, password);

                _FtpWebRequest.UseBinary = true;
                _FtpWebRequest.KeepAlive = false;
                
                //Connect(uri);
                _FtpWebRequest.Method = WebRequestMethods.Ftp.DeleteFile;
                FtpWebResponse response = (FtpWebResponse)_FtpWebRequest.GetResponse();
                response.Close();

                return true;
            }
            catch (Exception ex)
            {
                ErrDsc = ex.Message;
                sw.WriteLine(ex.Message);
                sw.WriteLine(ex.StackTrace);
                return false;
            }
        }

        /// <summary>
        /// 从 FTP上 取得所给目录的文件列表。
        /// </summary>
        /// <param name="FtpServerPath">FTP目录, 形如：ftp://192.168.1.113/MS/MS1207050006/</param>
        /// <param name="ResultData"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="sw"></param>
        /// <param name="ErrDsc"></param>
        /// <returns></returns>
        public bool GetFileListFromFTPPath(string FtpServerPath, ref ArrayList ResultData, ref StringWriter sw, ref String ErrDsc)
        {
            //"ftp://" + ftpServerIP + "/"ime
            sw.WriteLine("GetFileListFromFTPPath  Start Time " + System.DateTime.Now.ToString());

            StringBuilder result = new StringBuilder();
            FtpWebRequest reqFTP;
            try
            {
                if (FtpServerPath.EndsWith("/"))
                {
                    FtpServerPath = FtpServerPath.Substring(0, FtpServerPath.Length - 1);
                }
                reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(FtpServerPath));
                reqFTP.UseBinary = true;
                reqFTP.Credentials = new NetworkCredential(username, password);
                reqFTP.Method = WebRequestMethods.Ftp.ListDirectory;
                WebResponse response = reqFTP.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream());
                string line = reader.ReadLine();
                while (line != null)
                {
                    result.Append(line);
                    sw.WriteLine("取出一个文件：" + line);
                    #region 去掉目录名，只需文件名 ***备注：Line会包含：最后一层目录名。比如：如果是取这个 ftp://192.168.1.113/MS/MS1207050006 目录下的，那么它的文件名会包含：MS1207050006/xxxx.xxx
                    string[] _str = line.Split('/');
                    ResultData.Add(_str[_str.Length - 1]);
                    #endregion
                    //ResultData.Add(line);

                    line = string.Empty;
                    line = reader.ReadLine();
                }
                // to remove the trailing '\n'     
                //result.Remove(result.ToString().LastIndexOf('\n'), 1); 
                reader.Close();
                response.Close();
                return true;
            }
            catch (Exception ex)
            {
                // System.Windows.Forms.MessageBox.Show(ex.Message);                  
                sw.WriteLine(ex.StackTrace);
                ErrDsc = ex.Message;
                return false;
            }
            finally
            {
                sw.WriteLine("GetFileListFromFTPPath  End Time " + System.DateTime.Now.ToString());
            }
        }
            
        /// <summary>
        /// 在 FTP 上建立目录，备注：建立时必须一次只建立一级目录
        /// </summary>
        /// <param name="directory">FTP目录，形如：ftp://192.168.1.113/MS/MS1207050006</param>
        /// <param name="sw"></param>
        /// <param name="ErrDsc"></param>
        /// <returns></returns>
        public bool CreateFTPDirectory(string directory, ref StringWriter sw, ref String ErrDsc)
        {
            try
            {
                if (directory.EndsWith("/"))
                {
                    directory = directory.Substring(0, directory.Length - 1);
                }

                sw.WriteLine(" CreateFTPDirectory Start Time " + System.DateTime.Now.ToString());

                //create the directory 
                FtpWebRequest requestDir = (FtpWebRequest)FtpWebRequest.Create(new Uri(directory));

                requestDir.Method = WebRequestMethods.Ftp.MakeDirectory;

                requestDir.Credentials = new NetworkCredential(username, password);

                requestDir.UsePassive = true;

                requestDir.UseBinary = true;

                requestDir.KeepAlive = false;

                FtpWebResponse response = (FtpWebResponse)requestDir.GetResponse();

                Stream ftpStream = response.GetResponseStream();

                ftpStream.Close();

                response.Close();

                return true;
            }
            catch (WebException ex)
            {
                FtpWebResponse response = (FtpWebResponse)ex.Response;

                if (response.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                {
                    response.Close();
                    return true;
                }
                else
                {
                    ErrDsc = ex.Message;
                    sw.WriteLine(ex.StackTrace);
                    response.Close();
                    return false;
                }
            }
            finally
            {
                sw.WriteLine("CreateFTPDirectory End Time " + System.DateTime.Now.ToString());
            }
        }
        

    }
}
