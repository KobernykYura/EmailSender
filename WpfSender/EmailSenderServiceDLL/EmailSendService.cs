using System;
using System.Linq;
using System.Net.Mail;
using System.Windows;
using System.Net;
using WPFData;

namespace EmailSenderServiceDLL
{
    public class EmailSendService
    {
        #region vars
        private string strLogin; // email c которого будет рассылаться почта
        private string strPassword; // пароль к email с которого будет рассылаться почта
        private string strSmtp = "smtp.mail.ru"; // smtp - server
        private int iSmtpPort = 25; // порт для smtp-server
        private string strBody; // текст письма для отправки
        private string strSubject; // тема письма для отправки
        #endregion

        public EmailSendService(string sLogin, string sPassword)
        {
            this.strLogin = sLogin;
            this.strPassword = sPassword;
        }
        /// <summary>
        /// Method for sending an email to a specific recipient.
        /// Отправка email конкретному адресату.
        /// </summary>
        /// <param name="mail">Mail address.</param>
        /// <param name="smtp">Mail smtp.</param>
        private void SendMail(string mail, string smtp, int port)
        {
            using (MailMessage mm = new MailMessage(strLogin, mail))
            {
                mm.Subject = strSubject;
                mm.Body = "Hello world!";
                mm.IsBodyHtml = false;
                SmtpClient sc = new SmtpClient(strSmtp, iSmtpPort);
                sc.EnableSsl = true;
                sc.DeliveryMethod = SmtpDeliveryMethod.Network;
                sc.UseDefaultCredentials = false;
                sc.Credentials = new NetworkCredential(strLogin, strPassword);
                try
                {
                    sc.Send(mm);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Невозможно отправить письмо " + ex.ToString());
                }

            }
        }//private void SendMail(string mail, string name)
        /// <summary>
        /// Method for sending mail to emails from list.
        /// </summary>
        /// <param name="emails">List of emails.</param>
        public void SendMails(IQueryable<Emails> emails)
        {
            foreach (Emails email in emails)
            {
                SendMail(email.Name, email.Value, email.Port);
            }
        }
    }
}
