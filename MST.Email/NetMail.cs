using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace MST.Email
{
    public class NetMail
    {
        //string mailTo,string mailCC,string mailSubject,string htmlBody
        public static void SendMail(List<object> mailInfo)
        {
            try
            {
                //string mailTo = "ghma@futurmaster.com,zhen.zhang@futurmaster.com,king.jiang@futurmaster.com,qingfen.zheng@futurmaster.com,yeyi.sun@futurmaster.com";
                string mailTo = mailInfo[0].ToString();
                string mailCc = mailInfo[1].ToString();//"CN_Test@futurmaster.com"
                string mailSubject = mailInfo[2].ToString();//[New OLAP Server]Smoke Test Report For Daily Build               
                string htmlBody = mailInfo[3].ToString();
                string mailFrom = mailInfo[4].ToString();//CN_Test@futurmaster.com
                string mailFromDisplay = mailInfo[5].ToString();//CN Test Team
                string mailAttach = mailInfo[6].ToString();
                string mailSmtpClient = mailInfo[7].ToString();//"SXCH.FUTURMASTER.COM"

                WebClient client = new WebClient();
                //string strbody = ReplaceText(userName, htmlBody, myName);
                MailMessage message = new MailMessage();
                message.To.Add(mailTo);
                MailAddress copy = new MailAddress(mailCc);
                message.CC.Add(copy);
                message.Subject = mailSubject;
                message.From = new MailAddress(mailFrom, mailFromDisplay);
                message.Body = htmlBody;
                message.IsBodyHtml = true;
                if (File.Exists(mailAttach))
                {
                    Attachment data = new Attachment(mailAttach, MediaTypeNames.Application.Octet);
                    message.Attachments.Add(data);
                }
                SmtpClient smtp = new SmtpClient(mailSmtpClient);
                smtp.Send(message);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
