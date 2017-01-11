using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tamir.SharpSsh;

namespace OOSCommon
{
    public class SFTPHelper
    {
        private SshTransferProtocolBase m_sshCp;
        private SFTPHelper()
        {

        }
        public SFTPHelper(SshConnectionInfo connectionInfo)
        {
            m_sshCp = new Sftp(connectionInfo.Host, connectionInfo.User);

            if (connectionInfo.Pass != null)
            {
                m_sshCp.Password = connectionInfo.Pass;
            }

            if (connectionInfo.IdentityFile != null)
            {
                m_sshCp.AddIdentityFile(connectionInfo.IdentityFile);
            }
        }

        public bool Connected
        {
            get
            {
                return m_sshCp.Connected;
            }
        }
        public void Connect()
        {
            if (!m_sshCp.Connected)
            {
                m_sshCp.Connect();
            }
        }
        public void Close()
        {
            if (m_sshCp.Connected)
            {
                m_sshCp.Close();
            }
        }
        public bool Upload(string localPath, string remotePath)
        {
            try
            {
                if (!m_sshCp.Connected)
                {
                    m_sshCp.Connect();
                }
                m_sshCp.Put(localPath, remotePath);

                return true;
            }
            catch
            {
                return false;
            }

        }
        public bool Download(string remotePath, string localPath)
        {
            try
            {
                if (!m_sshCp.Connected)
                {
                    m_sshCp.Connect();
                }

                m_sshCp.Get(remotePath, localPath);

                return true;
            }
            catch(Exception ex)
            {
                throw new Exception(ex.Message, ex);
                //return false;
            }
        }
        public bool Delete(string remotePath)
        {
            try
            {
                if (!m_sshCp.Connected)
                {
                    m_sshCp.Connect();
                }
                ((Sftp)m_sshCp).Delete(remotePath);//刚刚新增的Delete方法

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
                //return false;
            }
        }

        public ArrayList GetFileList(string path)
        {
            try
            {
                if (!m_sshCp.Connected)
                {
                    m_sshCp.Connect();
                }
                return ((Sftp)m_sshCp).GetFileList(path);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
                //return null;
            }
        }
    }


    //创建一个辅助类
    public class SshConnectionInfo
    {
        public string IdentityFile { get; set; }
        public string Pass { get; set; }
        public string Host { get; set; }
        public string User { get; set; }
    }

    //上传文件
    //SshConnectionInfo objInfo = new SshConnectionInfo();
    //objInfo.User = "username";
    //        objInfo.Host = "host";
    //        objInfo.IdentityFile = "key"; //有2中认证，一种基于PrivateKey,一种基于password
    //        //objInfo.Pass = "password"; 基于密码
    //        SFTPHelper objSFTPHelper = new SFTPHelper(objInfo);
    //objSFTPHelper.Upload("localFile", "remoteFile");




    //下载文件
    //objSFTPHelper.Download("remoteFile", "localFile");

    //删除文件
    //objSFTPHelper.Delete("remoteFile");

    //遍历远程文件夹
    //ArrayList fileList = objSFTPHelper.GetFileList("remotePath");


}
