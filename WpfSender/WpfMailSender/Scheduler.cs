using EmailSenderServiceDLL;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using WPFData;

namespace WpfMailSender
{
    /// <summary>
    /// Класс планировщик, который создает расписание, следит за его выполнением и напоминает о событиях
    /// Так же помогает автоматизировать рассылку писем в соответствии с расписанием
    /// </summary>
    class Scheduler
    {
        DispatcherTimer timer = new DispatcherTimer(); // таймер 
        EmailSendService emailSender; // экземпляр класса отвечающего за отправку писем
        DateTime dtSend; // дата и время отправки
        IQueryable<Emails> emails; // коллекция email'ов адресатов

        /// <summary>
        /// Методе который превращаем строку из текстбокса tbTimePicker в TimeSpan
        /// </summary>
        /// <param name="strSendTime"></param>
        /// <returns></returns>
        public TimeSpan GetSendTime(string strSendTime)
        {
            TimeSpan tsSendTime = new TimeSpan();
            try
            {
                tsSendTime = TimeSpan.Parse(strSendTime);
            }
            catch { }

            return tsSendTime;
        }
        ///
        /// <summary>
        /// Непостредственно отправка писем
        /// </summary>
        /// <param name="dtSend"></param>
        /// <param name="emailSender"></param>
        /// <param name="emails"></param>
        public void SendEmails(DateTime dtSend, EmailSendService emailSender, IQueryable<Emails> emails)
        {
            this.emailSender = emailSender; // Экземпляр класса отвечающего за отправку писем присваиваем 
            this.dtSend = dtSend;
            this.emails = emails;
            timer.Tick += Timer_Tick;
            timer.Interval = new TimeSpan(0, 0, 1);
            timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (dtSend.ToShortTimeString() == DateTime.Now.ToShortTimeString())
            {
                emailSender.SendMails(emails);
                timer.Stop();
                MessageBox.Show("Письма отправлены.");
            }
        }

    }
}
