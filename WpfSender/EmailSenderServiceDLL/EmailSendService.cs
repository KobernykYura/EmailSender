using System;
using System.Linq;
using System.Net.Mail;
using System.Windows;
using System.Net;
using System.IO;
using CodePasswordDLL;
using System.Threading;
using Common;

namespace EmailSenderServiceDLL
{
    public class EmailSendService
    {
        #region vars
        private string strLogin; // email c которого будет рассылаться почта
        private string strPassword; // пароль к email с которого будет рассылаться почта
        private string strSmtp; // smtp - server
        private int iSmtpPort; // порт для smtp-server
        private string fileExceptionPath = @"../../files/TextException.txt"; // путь к файлу с описанием ошибки отправки
        private string strBody= File.ReadAllText(@"../../files/TextBody.txt"); // текст письма для отправки
        private string strSubject="Hello from WPF."; // тема письма для отправки
        #endregion
        public string Body
        {
            get { return strBody; }
            set { strBody = value; }
        }
        public string Subject
        {
            get { return strSubject; }
            set { strSubject = value; }
        }

        public EmailSendService(string sLogin, string sPassword, string smtp, int port)
        {
            this.strLogin = sLogin;
            this.strPassword = sPassword;
            this.strSmtp = smtp;
            this.iSmtpPort = port;
        }
        /// <summary>
        /// Method for sending an email to a specific recipient.
        /// Отправка email конкретному адресату.
        /// </summary>
        /// <param name="mail">Mail address.</param>
        /// <param name="smtp">Mail smtp.</param>
        private void SendMail(object objEmail)
        {
            Email obj = (Email)objEmail;
            string mail = obj.Name;
            string smtp = obj.Value;
            int port = obj.Port;

            using (MailMessage mm = new MailMessage(strLogin, mail))
            {
                mm.Subject = strSubject;
                mm.Body = strBody;
                mm.IsBodyHtml = false;
                using (SmtpClient sc = new SmtpClient(strSmtp, iSmtpPort))
                {
                    sc.EnableSsl = true;
                    sc.DeliveryFormat = SmtpDeliveryFormat.International;
                    sc.DeliveryMethod = SmtpDeliveryMethod.Network;
                    sc.UseDefaultCredentials = false;
                    sc.Credentials = new NetworkCredential(strLogin, CodePassword.getPassword(strPassword));
                    try
                    {
                        sc.Send(mm);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Невозможно отправить письмо " + ex.ToString());
                        File.AppendAllText(fileExceptionPath, DateTime.Now.ToLongDateString() + DateTime.Now.ToLongTimeString() + "\r\n" + ex.ToString() + "\r\n\r\n");
                    }
                }

            }
        }//private void SendMail(string mail, string name)
        /// <summary>
        /// Method for sending mail to emails from list.
        /// </summary>
        /// <param name="emails">List of emails.</param>
        public void SendMails(IQueryable<Email> emails)
        {
            foreach (Email email in emails)
            {
                //SendMail(email.Name, email.Value, email.Port);
                Thread thread = new Thread(new ParameterizedThreadStart(SendMail));
                thread.Start(email);
            }
        }
    }
}
