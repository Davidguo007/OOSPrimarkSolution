using OOSCommon;
using OOSLibrary.PrimarkOriginDB.Business;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;

namespace OOSPrimarkFTPAgent
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Diagnostics.Process[] myProcesses = System.Diagnostics.Process.GetProcessesByName(System.Diagnostics.Process.GetCurrentProcess().ProcessName);
            if (myProcesses.Length > 1)
            {
                Console.WriteLine("程序已经在运行! 请关闭之前的程序。");
                Console.ReadLine();
                return;
            }

            StringWriter sw = new StringWriter();

            string MailSendFrom = System.Configuration.ConfigurationManager.AppSettings["MailFromAddress"];
            string MailSendTo = System.Configuration.ConfigurationManager.AppSettings["MailSendToAddress"];
            string _errMailSubject = "OOS错误邮件: OOSPrimarkSFTP Download and inport Data failed: ";

            //string _SYS_Status = System.Configuration.ConfigurationManager.AppSettings["OOSPrimarkFTPAgent_Status"];

            //Console.WriteLine("读取程序执行标志，判断是否可以执行.");

            //if (_SYS_Status.Trim().ToLower() != "ok_to_run")
            //{
            //    #region 中止运行提示

            //    _errMailSubject += "上一次此程序执行有错误，不可以执行.";
            //    sw.WriteLine("上一次程序执行出错误后将程序配置文件中的标志：OOSPrimarkFTPAgent_Status 置成了：Error, 请确保已找到上一次的执行出错原因并已解决后，将配置文件中的此标志改成：OK_to_Run 即可再次运行。 ");
            //    MailSendWebSvc.MailSend msErr = new MailSendWebSvc.MailSend();
            //    msErr.SendMessage(MailSendFrom, MailSendTo, _errMailSubject + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), false, sw.ToString());
            //    Console.WriteLine("上一次此程序执行有错误，不可以执行.");
            //    Console.WriteLine("因为上次程序执行出错后将程序配置文件中的标志：, ");
            //    Console.WriteLine("OOSPrimarkFTPAgent_Status 置成了：Error.");
            //    Console.WriteLine("请确保已找到上一次的执行出错原因并已解决。");
            //    Console.WriteLine("然后手动将配置文件中的此标志改成：OK_to_Run 即可再次运行。");
            //    Console.ReadLine();
            //    Console.ReadLine();
            //    #endregion

            //    return;
            //}

            string PATH = System.Configuration.ConfigurationManager.AppSettings["LogPath"] + DateTime.Now.ToString("yyyy-MM-dd") + @"\";
            string FTPserver = System.Configuration.ConfigurationManager.AppSettings["FTPserver"];
            string FTPuser = System.Configuration.ConfigurationManager.AppSettings["FTPuser"];
            string FTPpass = System.Configuration.ConfigurationManager.AppSettings["FTPpass"];
            string FTPFolder = System.Configuration.ConfigurationManager.AppSettings["SFTPFolder"];
            string WSServerlePath = System.Configuration.ConfigurationManager.AppSettings["WSServerlePath"];
            string _DownLoadRetryTimes = System.Configuration.ConfigurationManager.AppSettings["DownRetryTimes"];
            string _DownLoadRetryWaitSec = System.Configuration.ConfigurationManager.AppSettings["DownRetryWaitSec"];
            string DatfileExtName = ".dat";

            int _theAPPRetryTimes = 1;
            int _fileStatus = 0;   //值分别为： 0，1，2，3，4， 5
                                   //初值为0， 取服务器文列表成功为 1， 下载成功后赋 2； 将FTP服务器上的文件删除成功为 3； 插入DB成功后赋 4； 将文件移到Done目录中为 5 ；
                                   //因为 1，2，3 是 和远程服务器进行交互，所以最有可能连接失败，所以需要Retry， 所以把最有可能出错的步骤先做掉，以使得后面的动作出错的概率非常小。
                                   
            int _prevRetryFileStatus = 0;
            bool _HasDat_Download = false;
            bool _IsDataEmpty = false;

            SshConnectionInfo objInfo = new SshConnectionInfo();
            objInfo.User = FTPuser;
            objInfo.Host = FTPserver;
            //objInfo.IdentityFile = "password"; //有2中认证，一种基于PrivateKey,一种基于password
            objInfo.Pass = FTPpass; //基于密码
            SFTPHelper objSFTPHelper = new SFTPHelper(objInfo);
            string _ftpFile = "";
            string _newFtpFileName = "";

            #region 手动将文件订单写入DB 
            //WSServerlePath = @"E:\OOSWorkData\FTPDat\DatOfficialOrder\";
            //_newFtpFileName = "PrimarkT-0531.dat";

            //sw.WriteLine("step 3. 将文件写入临时DB." + DateTime.Now);
            //Console.WriteLine("step 3. 将文件写入临时DB." + DateTime.Now);
            //Console.WriteLine();
            //PrimarkFTPDatAgent _ftpDatAgent = new PrimarkFTPDatAgent();
            ////StringWriter sw11 = new StringWriter();
            //if (!_ftpDatAgent.RecivePrimarkOrigFile(WSServerlePath, _newFtpFileName, ref sw))
            //{
            //    throw new Exception("将文件[" + _newFtpFileName + "]写入临时DB 失败." + DateTime.Now);
            //}
            //else
            //{
            //    sw.WriteLine("将文件[" + _newFtpFileName + "]写入临时DB 成功." + DateTime.Now);
            //    Console.WriteLine("将文件[" + _newFtpFileName + "]写入临时DB 成功." + DateTime.Now);
            //    Console.WriteLine();
            //}
            //return;
            #endregion

            string filename = "";

            while (_theAPPRetryTimes <= int.Parse(_DownLoadRetryTimes))
            {
                filename = "_OOSPrimarkFTPAgent_";
                try
                {
                    sw.WriteLine("Start to Download from Primark SFTP.." + DateTime.Now);
                    sw.WriteLine();
                    #region 下载并处理

                    if (_fileStatus == 0)
                    {

                        #region 取服务器上文件列表

                        sw.WriteLine("取出SFTP服务器上的目录 ./users/checkpoint/from_ipms/ 下的所有文件列表." + DateTime.Now);
                        Console.WriteLine("取出SFTP服务器上的目录 ./users/checkpoint/from_ipms/ 下的所有文件列表：" + DateTime.Now);

                        ArrayList _sFtpLst = objSFTPHelper.GetFileList("/users/checkpoint/from_ipms/");
                        List<string> _sLstFTP = new List<string>();
                        foreach (string item in _sFtpLst)
                        {
                            if (item != "." & item != "..")
                            {
                                _sLstFTP.Add(item);
                            }
                        }

                        if (_sLstFTP.Count <= 0)
                        {
                            sw.WriteLine("远程服务器上没有文件需要下载。" + DateTime.Now);
                            Console.WriteLine("远程服务器上没有文件需要下载。" + DateTime.Now);
                            Console.WriteLine();
                            //return;
                            break;
                        }

                        #endregion

                        _HasDat_Download = true;
                        _ftpFile = _sLstFTP[0]; //处理第一个。
                        _fileStatus = 1;

                        sw.WriteLine("已取出了 1 个FTP服务器上的文件。[ " + _ftpFile + " ] " + DateTime.Now);
                        Console.WriteLine("已取出了 1 个FTP服务器上的文件。[ " + _ftpFile + " ] " + DateTime.Now);
                        Console.WriteLine();

                        sw.WriteLine("开始下载文件：" + DateTime.Now);
                        sw.WriteLine();
                        Console.WriteLine("开始下载文件:" + DateTime.Now);
                        Console.WriteLine();
                    }

                    if (_ftpFile.Trim() != "" && _fileStatus == 1)
                    {
                        #region step 1. 下载                    
                        sw.WriteLine("step 1. 当前下载的文件是：[" + _ftpFile + "].   " + DateTime.Now);
                        Console.WriteLine("step 1. 当前下载的文件是：[" + _ftpFile + "].  " + DateTime.Now);
                        Console.WriteLine();
                        _newFtpFileName = _ftpFile.Replace(DatfileExtName.ToUpper(), "").Replace(DatfileExtName, "") + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".dat";

                        bool _downStatus = objSFTPHelper.Download(FTPFolder + _ftpFile, WSServerlePath + _newFtpFileName);

                        if (!_downStatus)
                        {
                            #region Retry

                            int _RetryTimes = 1;
                            while (!_downStatus && _RetryTimes <= int.Parse(_DownLoadRetryTimes))
                            {
                                sw.WriteLine("第【" + _RetryTimes + "】次下载失败. " + DateTime.Now);
                                Console.WriteLine("第【" + _RetryTimes + "】次下载失败." + DateTime.Now);
                                Thread.Sleep(int.Parse(_DownLoadRetryWaitSec) * 1000);
                                sw.WriteLine("等待了【" + _DownLoadRetryWaitSec + "】分钟后重试。" + DateTime.Now);
                                Console.WriteLine("等待了【" + _DownLoadRetryWaitSec + "】分钟后重试。" + DateTime.Now);
                                _downStatus = objSFTPHelper.Download(FTPFolder + _ftpFile, WSServerlePath + _newFtpFileName);
                                _RetryTimes++;
                            }

                            if (_downStatus)
                            {
                                sw.WriteLine("共重试下载了【" + _RetryTimes + "】次后，文件下载成功。");
                                Console.WriteLine("共重试下载了【" + _RetryTimes + "】次后，文件下载成功。" + DateTime.Now);
                                Console.WriteLine();
                            }
                            else
                            {
                                //sw.WriteLine("共重试下载了【" + _RetryTimes + "】次后，文件仍下载失败。发送报错邮件。");
                                throw new Exception("共重试下载了【" + _RetryTimes + "】次后，文件仍下载失败。发送报错邮件。");
                            }

                            #endregion
                        }
                        else
                        {
                            sw.WriteLine("下载文件成功。" + DateTime.Now);
                            Console.WriteLine("下载文件成功。" + DateTime.Now);
                            Console.WriteLine();
                        }

                        sw.WriteLine("下载到本地的文件名为：[ " + _newFtpFileName + " ]. ");
                        Console.WriteLine("下载到本地的文件名为：[ " + _newFtpFileName + " ].");
                        Console.WriteLine();

                        #endregion

                        _fileStatus = 2;
                        Thread.Sleep(2000);

                        FileInfo fi = new FileInfo(WSServerlePath + _newFtpFileName);
                        sw.WriteLine("此文件的 Size 为：[ " + fi.Length + " ]. ");
                        sw.WriteLine();
                        Console.WriteLine("此文件的 Size 为：[ " + fi.Length + " ].");
                        Console.WriteLine();

                        if (fi.Length <= 0)
                        {
                            sw.WriteLine("此文件的 Size 为：[ 0 ]. 说明客人没有数据需要导入,退出");
                            sw.WriteLine();
                            Console.WriteLine("此文件的 Size 为：[ 0 ]. 说明客人没有数据需要导入,退出");
                            Console.WriteLine();
                            break;
                        }

                        sw.WriteLine("取得上一次下载的文件大小，判断若是相同，则认为是同一份，不下载。");
                        Console.WriteLine("取得上一次下载的文件大小，判断若是相同，则认为是同一份，不下载。");
                        OriginDatBase _DBhelp = new OriginDatBase();
                        string _sqlLastDat = "select top 1 FileName,FileSize_Bytes from dbo.OriginOrderHead order by Createtime desc ";
                        DataTable dtLast = _DBhelp.GetDBDataBySQL(_sqlLastDat);
                        sw.WriteLine("上一次的文件[ " + dtLast.Rows[0]["FileName"].ToString() + " ]的大小为：[ " + dtLast.Rows[0]["FileSize_Bytes"].ToString() + " ]。");
                        Console.WriteLine("上一次的文件[ " + dtLast.Rows[0]["FileName"].ToString() + " ]的大小为：[ " + dtLast.Rows[0]["FileSize_Bytes"].ToString() + " ]。");

                        if (fi.Length == long.Parse(dtLast.Rows[0]["FileSize_Bytes"].ToString()))
                        {
                            sw.WriteLine("与上一次的文件的大小相同，不导入。");
                            Console.WriteLine("与上一次的文件的大小相同，不导入。");
                            break;
                        }
                        else
                        {
                            sw.WriteLine("与上一次的文件的大小不相同，要导入。将 _fileStatus 置3，跳过删除步骤。");
                            Console.WriteLine("与上一次的文件的大小不相同，要导入。将 _fileStatus 置3，跳过删除步骤。");
                            _fileStatus = 3;
                            Thread.Sleep(3000);
                        }
                    }

                    if (_ftpFile.Trim() != "" && _fileStatus == 2)
                    {
                        #region step 2. 下载完删除远程服务器上的这个文件

                        //sw.WriteLine("step 2. 下载完删除远程服务上的这个文件：[ " + _ftpFile + " ]. " + DateTime.Now);
                        //Console.WriteLine("step 2.下载完删除远程服务上的这个文件：[ " + _ftpFile + " ]." + DateTime.Now);

                        //bool _DeleStatus = objSFTPHelper.Delete(FTPFolder + _ftpFile);

                        //if (!_DeleStatus)
                        //{
                        //    #region Retry

                        //    int _RetryTimes = 1;
                        //    while (!_DeleStatus && _RetryTimes <= int.Parse(_DownLoadRetryTimes))
                        //    {
                        //        sw.WriteLine("第【" + _RetryTimes + "】次删除失败. " + DateTime.Now);
                        //        Console.WriteLine("第【" + _RetryTimes + "】次删除失败." + DateTime.Now);
                        //        Thread.Sleep(int.Parse(_DownLoadRetryWaitSec) * 1000);

                        //        sw.WriteLine("等待了【" + _DownLoadRetryWaitSec + "】分钟后重试。" + DateTime.Now);
                        //        Console.WriteLine("等待了【" + _DownLoadRetryWaitSec + "】分钟后重试。" + DateTime.Now);

                        //        _DeleStatus = objSFTPHelper.Delete(FTPFolder + _ftpFile);

                        //        _RetryTimes++;
                        //    }

                        //    if (_DeleStatus)
                        //    {
                        //        sw.WriteLine("共重试了【" + _RetryTimes + "】次后，文件删除成功。");
                        //        sw.WriteLine();
                        //        Console.WriteLine("共重试了【" + _RetryTimes + "】次后，文件删除成功。" + DateTime.Now);
                        //        Console.WriteLine();
                        //    }
                        //    else
                        //    {
                        //        //sw.WriteLine("共重试下载了【" + _RetryTimes + "】次后，文件仍下载失败。发送报错邮件。");
                        //        throw new Exception("共重试下载了【" + _RetryTimes + "】次后，文件仍下载失败。发送报错邮件。");
                        //    }

                        //    #endregion
                        //}
                        //else
                        //{
                        //    sw.WriteLine("删除远程服务上的这个文件成功。 " + DateTime.Now);
                        //    sw.WriteLine();
                        //    Console.WriteLine("删除远程服务上的这个文件成功。" + DateTime.Now);
                        //    Console.WriteLine();

                        //}

                        #endregion

                        _fileStatus = 3;
                        Thread.Sleep(3000);
                    }

                    if (_newFtpFileName.Trim() != "" && _fileStatus == 3)
                    {
                        #region step 3. 将文件写入临时DB, 写入成功后，

                        sw.WriteLine("step 3. 将文件写入临时DB." + DateTime.Now);
                        Console.WriteLine("step 3. 将文件写入临时DB." + DateTime.Now);
                        Console.WriteLine();
                        PrimarkFTPDatAgent _ftpDatAgent = new PrimarkFTPDatAgent();
                        //StringWriter sw11 = new StringWriter();
                        string _errCode = "";
                        if (!_ftpDatAgent.RecivePrimarkOrigFile(WSServerlePath, _newFtpFileName, ref sw, ref _errCode))
                        {
                            if (_errCode.Trim().ToLower() == "no data")
                            {
                                sw.WriteLine("文件[" + _newFtpFileName + "] 没有数据不写临时DB." + DateTime.Now);
                                sw.WriteLine();
                                Console.WriteLine("文件[" + _newFtpFileName + "] 没有数据不写临时DB." + DateTime.Now);
                                Console.WriteLine();
                                _IsDataEmpty = true;
                                break;//退出循环 [2016-10-13] david.
                            }
                            else
                            {
                                throw new Exception("将文件[" + _newFtpFileName + "]写入临时DB 失败." + DateTime.Now);
                            }
                        }
                        else
                        {
                            sw.WriteLine("将文件[" + _newFtpFileName + "]写入临时DB 成功." + DateTime.Now);
                            sw.WriteLine();
                            Console.WriteLine("将文件[" + _newFtpFileName + "]写入临时DB 成功." + DateTime.Now);
                            Console.WriteLine();
                        }
                        #endregion

                        _fileStatus = 4;
                        Thread.Sleep(1000);
                    }

                    if (_newFtpFileName.Trim() != "" && _fileStatus == 4)
                    {
                        #region step 4. 执行成功，将文件移到Done文件夹。

                        sw.WriteLine("step 4. 执行成功，将文件移到Done文件夹。");
                        Console.WriteLine("step 4. 执行成功，将文件移到Done文件夹。" + DateTime.Now);
                        Console.WriteLine();
                        File.Move(WSServerlePath + _newFtpFileName, WSServerlePath + @"Done\" + _newFtpFileName);

                        #endregion

                        _fileStatus = 5;
                    }

                    #endregion

                    sw.WriteLine();
                    sw.WriteLine("FTP上的文件下载和导入成功，本次程序执行成功。" + DateTime.Now);
                    sw.WriteLine();
                    Console.WriteLine("FTP上的文件下载和导入成功，本次程序执行成功。" + DateTime.Now);
                    Console.WriteLine(); Console.WriteLine(); Console.WriteLine();

                }
                catch (Exception ex)
                {
                    sw.WriteLine("程序执行异常：" + ex.Message);
                    Console.WriteLine("程序执行异常：" + ex.Message + "  " + DateTime.Now);
                    filename += "-False";

                    #region 异常及Retry处理

                    if (_prevRetryFileStatus < _fileStatus)
                    {
                        _prevRetryFileStatus = _fileStatus;
                        _theAPPRetryTimes = 1;
                    }

                    if (_fileStatus < 2)
                    {
                        _newFtpFileName = "";
                    }
                    if (_fileStatus < 4 && _theAPPRetryTimes < int.Parse(_DownLoadRetryTimes))
                    {
                        sw.WriteLine("下载发生异常，并且重试次数为[ " + _theAPPRetryTimes + " ], 少于等于[ " + _DownLoadRetryTimes + " ] 次。 " + DateTime.Now);
                        Console.WriteLine("下载发生异常，并且重试次数[ " + _theAPPRetryTimes + " ], 少于等于[" + _DownLoadRetryTimes + "] 次。" + DateTime.Now);
                        Console.WriteLine();
                        _theAPPRetryTimes++;

                        Thread.Sleep(int.Parse(_DownLoadRetryWaitSec) * 1000);
                        sw.WriteLine("等待了【" + _DownLoadRetryWaitSec + "】分钟后重试。" + DateTime.Now);
                        Console.WriteLine("等待了【" + _DownLoadRetryWaitSec + "】分钟后重试。" + DateTime.Now);
                        sw.WriteLine("将跳转到循环起始处，然后根据 _fileStatus 值[ " + _fileStatus + " ]执行相应的步骤。");

                        continue;
                    }

                    if (_theAPPRetryTimes >= int.Parse(_DownLoadRetryTimes))
                    {
                        sw.WriteLine("下载发生异常，并且重试次数为[ " + _theAPPRetryTimes + " ], 大于[ " + _DownLoadRetryTimes + " ] 次，停止程序运行且发送报错邮件。 " + DateTime.Now);
                        Console.WriteLine("下载发生异常，并且重试次数[ " + _theAPPRetryTimes + " ], 大于[" + _DownLoadRetryTimes + "] 次，停止程序运行且发送报错邮件" + DateTime.Now);
                        Console.WriteLine();
                    }

                    //sw.WriteLine("将程序的App.config中的 OOSPrimarkFTPAgent_Status 置 Error. ");
                    //ConfigHelper.UpdateAppConfig("OOSPrimarkFTPAgent_Status", "Error");
                    //sw.WriteLine("将程序的App.config中的 OOSPrimarkFTPAgent_Status 置 Error 成功. ");
                    //Console.WriteLine("将程序的App.config中的 OOSPrimarkFTPAgent_Status 置 Error 成功. " + DateTime.Now);
                    if (!MailSendTo.Trim().ToLower().Contains("david.guo@maxim-group.com"))
                    {
                        MailSendTo += ";david.guo@maxim-group.com";
                    }
                    #region 执行出现异常，发送报错邮件
                    MailSendWebSvc.MailSend ms = new MailSendWebSvc.MailSend();
                    if (ms.SendMessage(MailSendFrom, MailSendTo, _errMailSubject + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"), false, sw.ToString()))
                    {
                        sw.WriteLine("发送报错邮件成功。");
                        Console.WriteLine("发送报错邮件成功。. " + DateTime.Now);
                    }
                    else
                    {
                        sw.WriteLine("发送报错邮件失败。等2分钟再试。");
                        Thread.Sleep(120000);
                        if (!ms.SendMessage(MailSendFrom, MailSendTo, _errMailSubject + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"), false, sw.ToString()))
                        {
                            sw.WriteLine("等了2分钟再试发送报错邮件 还是失败。");
                            Console.WriteLine("第1次发送报错邮件失败，等了2分钟后再次发送成功。. " + DateTime.Now);
                        }
                    }
                    #endregion

                    #endregion

                    break;
                }
                finally
                {
                    if (_fileStatus == 5)
                    {
                        _fileStatus = 0;

                        #region 执行成功，发送一封成功邮件，工作日未收到邮件说明有问题。
                        MailSendWebSvc.MailSend ms = new MailSendWebSvc.MailSend();
                        if (ms.SendMessage(MailSendFrom, MailSendTo, "Primark Dat订单下载及导入成功." + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"), false, sw.ToString()))
                        {
                            sw.WriteLine("Primark SFTP订单Dat下载并导入成功，发送成功邮件。");
                            Console.WriteLine("Primark SFTP订单Dat下载并导入成功，发送成功邮件。 " + DateTime.Now);
                        }
                        else
                        {
                            sw.WriteLine("发送报错邮件失败。等2分钟再试。");
                            Thread.Sleep(120000);
                            if (!ms.SendMessage(MailSendFrom, MailSendTo, "Primark Dat订单下载及导入成功." + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"), false, sw.ToString()))
                            {
                                sw.WriteLine("等了2分钟再试发送报错邮件 还是失败。");
                                Console.WriteLine("第1次发送报错邮件失败，等了2分钟后再次发送成功。. " + DateTime.Now);
                            }
                        }
                        #endregion
                    }
                }
            }

            sw.WriteLine();
            sw.WriteLine("释放 objSFTPHelper 对象。");

            objSFTPHelper.Close();

            sw.WriteLine("释放 objSFTPHelper 对象成功。");

            if (!_HasDat_Download && !filename.Contains("FtpNoFile"))
            {
                filename += "-FtpNoFile-";
            }
            if (_IsDataEmpty && !filename.Contains("FtpFileIsEmpty"))
            {
                filename += "-FtpFileIsEmpty-";
            }

            string file = filename + "--" + System.DateTime.Now.ToFileTime() + ".txt";
            if (!Directory.Exists(PATH))
            {
                Directory.CreateDirectory(PATH);
            }
            File.WriteAllText(PATH + file, sw.ToString(), System.Text.Encoding.UTF8);

            return;
        }
    }
}
