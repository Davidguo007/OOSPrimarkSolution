using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.IO;

namespace OOSCommon
{
    public class MailAgent
    {
        public static readonly string TestEmailReciver = "mis_oos@maxim-group.com.cn";
        public static void SendErrorMail(string _ErrorNumber, string _subject, string _User, string _ProgramPageName, string _UploadedFile, string _logfilename, string _logStr, ref StringWriter sw)
        {
            try
            {
                string _fromAddress = ConfigurationManager.AppSettings["FromAddress"];
                string _toAddress = ConfigurationManager.AppSettings["ErrorMailToAddress"];
                sw.WriteLine("Mail Sender: " + _fromAddress);
                sw.WriteLine("Mail ToAddress: " + _toAddress);
                sw.WriteLine("Mail Subject: OOS操作错误响应邮件: [错误编号：" + _ErrorNumber + "] - 错误描述： " + _subject);
                sw.WriteLine("错误触发人  ：" + _User);
                sw.WriteLine("错误触发时间：" + DateTime.Now.ToString());
                sw.WriteLine("错误发生页面：" + _ProgramPageName);

                List<string> _attchList = new List<string>();
                if (!string.IsNullOrEmpty(_UploadedFile))
                {
                    sw.WriteLine("错误时的上传文件：" + _UploadedFile);
                    _attchList.Add(_UploadedFile);
                }

                sw.WriteLine("记录错误的Log名：" + _logfilename);
                _attchList.Add(_logfilename);
                string[] _attachments = _attchList.ToArray();

                string _MailSubject = "OOS错误邮件: [错误编号：" + _ErrorNumber + "] - [User：" + _User + "] 描述： " + _subject;
                string _mailBody = ErrorMailBodyHtml(_ErrorNumber, _subject, _User, _ProgramPageName, _UploadedFile, _logfilename);

                //    WSMailSend.MailSend _ms = new OOSCRM.OrderPrint.WSMailSend.MailSend();
                //     _ms.SendMessageMultiCCAttach(_fromAddress, _toAddress, null, _MailSubject, true, _mailBody, _attachments);

                sw.WriteLine("错误通知邮件发送成功！");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }

        }

        public static string ErrorMailBodyHtml(string _ErrorNumber, string _ErrorDesc, string _User, string _ProgramPageName, string _UploadedFile, string _logfilename)
        {
            StringBuilder sbHtml = new StringBuilder();
            sbHtml.AppendFormat(@"
<p>
<b><strong><font size=3 >错误编号    ：</font></strong></b><font size=3 >{0}</font><br>
<b><strong><font size=3 >错误触发人  ：</font></strong></b><font size=3 >{1}</font><br>
<b><strong><font size=3 >错误触发时间：</font></strong></b><font size=3 >{2}</font><br>
<b><strong><font size=3 >错误发生页面：</font></strong></b><font size=3 >{3}</font>
</p><br><br>", _ErrorNumber, _User, DateTime.Now.ToString(), _ProgramPageName);

            sbHtml.Append(@"<p>");
            if (!string.IsNullOrEmpty(_UploadedFile))
            {
                sbHtml.AppendFormat(@"<b><strong><font size=3 >上传文件名: </font></strong></b><font size=3 >{0}</font><br>", _UploadedFile);
            }
            sbHtml.AppendFormat(@"<b><strong><font size=3 >程序执行Log文件名: </font></strong></b><font size=3 >{0}</font><br><br>", _logfilename);
            sbHtml.AppendFormat(@"<b><strong><font size=3 >错误概述: </font></strong></b><font size=3 >{0}</font>", _ErrorDesc);
            sbHtml.Append(@"</p><br><br>");
            sbHtml.Append(@"<p><b><strong><font size=3 >程序执行Log请查看附档。</font></strong></b></p>");

            return sbHtml.ToString();
        }




    }
}
