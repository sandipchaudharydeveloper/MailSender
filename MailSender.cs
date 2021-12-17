using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Web;

namespace EmailBulkingFromExcel
{
    public class MailSender
    {

        #region Mail Send Method
        /// <summary>
        /// Mail Send Method
        /// </summary>
        /// <param name="To">
        ///     You Wan't To Send Mail It's Mail Address Pass Here In List Of String
        /// </param>
        /// <param name="Subject">
        ///      You Wan't To Send Mail Subject It's Subject Pass Here In String
        /// </param>
        /// <param name="Body">
        ///      You Wan't To Send Mail Body It's Body Pass Here In String
        /// </param>
        /// <param name="LstFile">
        ///      You Wan't To Send Mail Attechments It's Attechments Pass Here In List Of HttpPostedFileFile
        /// </param>
        /// <returns></returns>
        public string MailSend(List<string> To, string Subject, string Body, List<HttpPostedFile> LstFile)
        {
            try
            {
                string from = "1402harshpatel@gmail.com"; //Enter Email Id Of Sender Here
                string password = "Patel@@123";           //Enter Email Id Password Sender Here

                foreach (string MailTo in To)
                {
                    Thread.Sleep(1000);
                    using (MailMessage mail = new MailMessage(from, MailTo))
                    {
                        mail.Subject = Subject;
                        mail.Body = Body;
                        foreach (HttpPostedFile uploadedFile in LstFile)
                        {
                            string fileName = Path.GetFileName(uploadedFile.FileName);
                            mail.Attachments.Add(new Attachment(uploadedFile.InputStream, fileName));
                        }
                        mail.IsBodyHtml = true;
                        SmtpClient smtp = new SmtpClient();
                        smtp.Host = "smtp.gmail.com";
                        smtp.EnableSsl = true;
                        NetworkCredential networkCredential = new NetworkCredential(from, password);
                        smtp.UseDefaultCredentials = true;
                        smtp.Credentials = networkCredential;
                        smtp.Port = 587;
                        smtp.Send(mail);
                    }
                }
            }
            catch (Exception ex)
            {

                return "fail !!! " + ex.Message;
            }
            return "success";
        }
        #endregion
    }
}