using System;
using System.Linq;
using System.Windows;
using Xceed.Wpf.Toolkit;
using WpfMailSender.AdditionalClasses;
using System.Windows.Controls;
using WPFData;
using EmailSenderServiceDLL;
using System.IO;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Data;
using Microsoft.Win32;

namespace WpfMailSender
{

    /// <summary>
    /// Логика взаимодействия для ClientWindow.xaml
    /// </summary>
    public partial class ClientWindow : Window
    {
        TextRange doc;
        string strLogin;
        string strPassword;
        string strSmtp;
        int port;
        
        public ClientWindow()
        {
            InitializeComponent();

            cbSenderSelect.ItemsSource = VariableClass.Senders;
            cbSenderSelect.DisplayMemberPath = "Key";
            cbSenderSelect.SelectedValuePath = "Value";

            cbSmtpSelect.ItemsSource = VariableClass.Addreses;
            cbSmtpSelect.DisplayMemberPath = "Key";
            cbSmtpSelect.SelectedValuePath = "Value";

            //LoadMail();

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
            try
            {
                strLogin = cbSenderSelect.Text;
                strPassword = cbSenderSelect.SelectedValue.ToString();
                strSmtp = cbSmtpSelect.Text;
                port = (int)cbSmtpSelect.SelectedValue;
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show("Incorrect sender input!\n Please, try again. ");
                return;
            }
            

            doc = new TextRange(rtbMailBody.Document.ContentStart, rtbMailBody.Document.ContentEnd);
            if (string.IsNullOrEmpty(strLogin) || string.IsNullOrEmpty(strPassword))
            {
                System.Windows.MessageBox.Show("Выберите отправителя");
                return;
            }
            if (string.IsNullOrEmpty(cbSmtpSelect.Text))
            {
                System.Windows.MessageBox.Show("Выберите smtp-сервер");
                return;
            }
            if (IsRichTextBoxEmpty(rtbMailBody))
            {
                System.Windows.MessageBox.Show("Не указан текст письма");
                tabMailBody.IsSelected = true;
                return;
            }
            EmailSendService emailSender = new EmailSendService(strLogin, strPassword, strSmtp, port);
            emailSender.SendMails((IQueryable<Emails>)dgEmails.ItemsSource);

        }
        /// <summary>
        /// Button to send mail on date.
        /// </summary>
        /// <param name="sender">Object.</param>
        /// <param name="e">Event.</param>
        private void btnSendOnDate_Click(object sender, RoutedEventArgs e)
        {
            strLogin = cbSenderSelect.Text;
            strPassword = cbSenderSelect.SelectedValue.ToString();
            strSmtp = cbSmtpSelect.Text;
            port = (int)cbSmtpSelect.SelectedValue;

            Scheduler sc = new Scheduler();
            TimeSpan tsSendTime = sc.GetSendTime(tbTimePicker.Text);
            if (tsSendTime == new TimeSpan())
            {
                System.Windows.MessageBox.Show("Некорректный формат даты");
                return;
            }
            DateTime dtSendDateTime = (cldSchedulDateTimes.SelectedDate ?? DateTime.Today).Add(tsSendTime);
            if (dtSendDateTime < DateTime.Now)
            {
                System.Windows.MessageBox.Show("Дата и время отправки писем не могут быть раньше, чем настоящее время");
                return;
            }
            EmailSendService emailSender = new EmailSendService(strLogin, strPassword, strSmtp, port);
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

        private void btnSaveMailText_Click(object sender, RoutedEventArgs e)
        {
            doc = new TextRange(rtbMailBody.Document.ContentStart, rtbMailBody.Document.ContentEnd);

            if (IsRichTextBoxEmpty(rtbMailBody))
            {
                System.Windows.MessageBox.Show("Не указан текст письма");
                tabMailBody.IsSelected = true;
                return;
            }

            File.WriteAllText(@"../../files/TextBody.txt",doc.Text);
            System.Windows.MessageBox.Show("Файл сохранен");
            #region Saving in file
            //SaveFileDialog sfd = new SaveFileDialog();
            //sfd.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|XAML Files (*.xaml)|*.xaml|All files (*.*)|*.*";
            //if (sfd.ShowDialog() == true)
            //{
            //doc = new TextRange(rtbMailBody.Document.ContentStart, rtbMailBody.Document.ContentEnd);
            //    using (FileStream fs = File.Create(sfd.FileName))
            //    {
            //        if (Path.GetExtension(sfd.FileName).ToLower() == ".rtf")
            //            doc.Save(fs, DataFormats.Rtf);
            //        else if (Path.GetExtension(sfd.FileName).ToLower() == ".txt")
            //            doc.Save(fs, DataFormats.Text);
            //        else
            //            doc.Save(fs, DataFormats.Xaml);
            //    }
            //}
            #endregion

        }
        /// <summary>
        /// Checking if RichTextBox is empty.
        /// </summary>
        /// <param name="rtb"></param>
        /// <returns></returns>
        public bool IsRichTextBoxEmpty(System.Windows.Controls.RichTextBox rtb)
        {
            if (rtb.Document.Blocks.Count == 0) return true;
            TextPointer startPointer = rtb.Document.ContentStart.GetNextInsertionPosition(LogicalDirection.Forward);
            TextPointer endPointer = rtb.Document.ContentEnd.GetNextInsertionPosition(LogicalDirection.Backward);
            return startPointer.CompareTo(endPointer) == 0;
        }
        private void LoadMail()
        {
            using (FileStream fs = File.Open(@"../../files/TextBody.txt", FileMode.Open))
            {
                FlowDocument document = XamlReader.Load(fs) as FlowDocument;
                if (document != null)
                    rtbMailBody.Document = document;
            }
        }

        private void TscTabSwitcher_btnNextClick(object sender, RoutedEventArgs e)
        {
            tabControl.SelectedIndex = tabControl.SelectedIndex + 1;
        }

        private void TscTabSwitcher_btnPreviousClick(object sender, RoutedEventArgs e)
        {
            //tscTabSwitcher.btnNextClick += TscTabSwitcher_btnNextClick;
            tabControl.SelectedIndex = tabControl.SelectedIndex - 1;
        }
    }
}
