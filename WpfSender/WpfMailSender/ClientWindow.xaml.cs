using System;
using System.Linq;
using System.Windows;
using WpfMailSender.AdditionalClasses;
using WPFData;
using EmailSenderServiceDLL;

namespace WpfMailSender
{
    /// <summary>
    /// Логика взаимодействия для ClientWindow.xaml
    /// </summary>
    public partial class ClientWindow : Window
    {
        public ClientWindow()
        {
            InitializeComponent();

            cbSenderSelect.ItemsSource = VariableClass.Senders;
            cbSenderSelect.DisplayMemberPath = "Key";
            cbSenderSelect.SelectedValuePath = "Value";

            cbSmtpSelect.ItemsSource = VariableClass.Addreses;
            cbSmtpSelect.DisplayMemberPath = "Key";
            cbSmtpSelect.SelectedValuePath = "Value";

            DataBaseClass db = new DataBaseClass();
            dgEmails.ItemsSource = db.Emails;

        }

        /// <summary>
        /// Form close item.
        /// </summary>
        /// <param name="sender">Object.</param>
        /// <param name="e">Event.</param>
        private void imClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        /// <summary>
        /// Button to send mail right now.
        /// </summary>
        /// <param name="sender">Object.</param>
        /// <param name="e">Event.</param>
        private void btnSendAtOnce_Click(object sender, RoutedEventArgs e)
        {
            string strLogin = cbSenderSelect.Text;
            string strPassword = cbSenderSelect.SelectedValue.ToString();
            if (string.IsNullOrEmpty(strLogin) || string.IsNullOrEmpty(strPassword))
            {
                MessageBox.Show("Выберите отправителя");
                return;
            }
            EmailSendService emailSender = new EmailSendService(strLogin, strPassword);
            emailSender.SendMails((IQueryable<Emails>)dgEmails.ItemsSource);

        }
        /// <summary>
        /// Button to send mail on date.
        /// </summary>
        /// <param name="sender">Object.</param>
        /// <param name="e">Event.</param>
        private void btnSendOnDate_Click(object sender, RoutedEventArgs e)
        {
            Scheduler sc = new Scheduler();
            TimeSpan tsSendTime = sc.GetSendTime(tbTimePicker.Text);
            if (tsSendTime == new TimeSpan())
            {
                MessageBox.Show("Некорректный формат даты");
                return;
            }
            DateTime dtSendDateTime = (cldSchedulDateTimes.SelectedDate ?? DateTime.Today).Add(tsSendTime);
            if (dtSendDateTime < DateTime.Now)
            {
                MessageBox.Show("Дата и время отправки писем не могут быть раньше, чем настоящее время");
                return;
            }
            EmailSendService emailSender = new EmailSendService(cbSenderSelect.Text, cbSenderSelect.SelectedValue.ToString());
            sc.SendEmails(dtSendDateTime, emailSender, (IQueryable<Emails>)dgEmails.ItemsSource);

        }
        /// <summary>
        /// Button to go on Scheduler tab.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnToSchedeuler_Click(object sender, RoutedEventArgs e)
        {
            tiFormation.IsSelected = false;
            tiScheduler.IsSelected = true;
        }
    }
}
