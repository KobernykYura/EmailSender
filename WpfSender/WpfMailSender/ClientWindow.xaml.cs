using System;
using System.Linq;
using System.Windows;
using Xceed.Wpf.Toolkit;
using WpfMailSender.AdditionalClasses;
using System.Windows.Controls;
using CodePasswordDLL;
using EmailSenderServiceDLL;
using System.IO;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Data;
using Microsoft.Win32;
using WpfMailSender.AddDeleteEditWindows;
using EFData;
using Common;

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
        EFBase db;

        public ClientWindow()
        {
            InitializeComponent();

            cbSenderSelect.ItemsSource = VariableClass.Senders;
            cbSenderSelect.DisplayMemberPath = "Key";
            cbSenderSelect.SelectedValuePath = "Value";

            cbSmtpSelect.ItemsSource = VariableClass.Addreses;
            cbSmtpSelect.DisplayMemberPath = "Key";
            cbSmtpSelect.SelectedValuePath = "Value";

            LoadMail();

            db = new EFBase();
            dgEmails.ItemsSource = db.MyEmails;

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
                strLogin = cbSenderSelect.ComboBoxText;
                strPassword = cbSenderSelect.SelectedValue.ToString();
                strSmtp = cbSmtpSelect.ComboBoxText;
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
            emailSender.SendMails((IQueryable<Email>)dgEmails.ItemsSource);

        }
        /// <summary>
        /// Button to send mail on date.
        /// </summary>
        /// <param name="sender">Object.</param>
        /// <param name="e">Event.</param>
        private void btnSendOnDate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                strLogin = cbSenderSelect.ComboBoxText;
                strPassword = cbSenderSelect.SelectedValue.ToString();
                strSmtp = cbSmtpSelect.ComboBoxText;
                port = (int)cbSmtpSelect.SelectedValue;
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show("Incorrect sender input!\n Please, try again. ");
                return;
            }

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
            sc.SendEmails(dtSendDateTime, emailSender, (IQueryable<Email>)dgEmails.ItemsSource);

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
        /// <summary>
        /// Saving mail in .txt file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
            tabControl.SelectedIndex = tabControl.SelectedIndex - 1;

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
        /// <summary>
        /// Preloading text from .txt file in RichTextBox.
        /// </summary>
        private void LoadMail()
        {
            string fileName = @"../../files/TextBody.txt";
            TextRange range;

            using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate))
            {
                range = new TextRange(rtbMailBody.Document.ContentStart, rtbMailBody.Document.ContentEnd);
                if (range != null)
                    range.Load(fs, DataFormats.Text);
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

        /// <summary>
        /// Add new sender.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSenderSelect_btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (cbSenderSelect.ComboBoxText == "") { System.Windows.MessageBox.Show("Email address is empty."); return; }
            AddWindow childWindow = new AddWindow();
            childWindow.Owner = this;
            childWindow.ShowDialog();

            try
            {
                VariableClass.Senders.Add(cbSenderSelect.ComboBoxText, CodePassword.getCodPassword(childWindow.Password));
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show("User with such address is already exist.");
                return;
            }
            cbSenderSelect.UpdateComboBox();
        }
        /// <summary>
        /// Delete this sender.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSenderSelect_btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (cbSenderSelect.ComboBoxText == "") { System.Windows.MessageBox.Show("The object to remove is not selected."); return; }

            MessageBoxResult result =
            System.Windows.MessageBox.Show($"Are you shure to delete {cbSenderSelect.ComboBoxText}?", $"Удаление: {cbSenderSelect.ComboBoxText}", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (result.ToString() == "Yes") { VariableClass.Senders.Remove(cbSenderSelect.ComboBoxText); cbSenderSelect.UpdateComboBox(); }

        }
        /// <summary>
        /// Change this sender.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSenderSelect_btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (cbSenderSelect.ComboBoxText == "") { System.Windows.MessageBox.Show("The object to edit is not selected."); return; }
            
            EditWindow childWindow = new EditWindow();
            childWindow.Owner = this;
            childWindow.InitializeContent(cbSenderSelect.ComboBoxText, CodePassword.getPassword((string)cbSenderSelect.SelectedValue));
            VariableClass.Senders.Remove(cbSenderSelect.ComboBoxText);

            childWindow.ShowDialog();

            VariableClass.Senders.Add(childWindow.Address, CodePassword.getCodPassword(childWindow.Password));
            cbSenderSelect.UpdateComboBox();

        }

        /// <summary>
        /// Add new smtp.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSmtpSelect_btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (cbSmtpSelect.ComboBoxText == "") { System.Windows.MessageBox.Show("SMTP address is empty."); return; }
            AddSMTPWindow childWindow = new AddSMTPWindow();
            childWindow.Owner = this;
            childWindow.ShowDialog();

            try
            {
                VariableClass.Addreses.Add(cbSmtpSelect.ComboBoxText, childWindow.Port);
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show("SMTP with such address is already exist.");
                return;
            }
            cbSmtpSelect.UpdateComboBox();
        }
        /// <summary>
        /// Delete this smtp.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSmtpSelect_btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (cbSmtpSelect.ComboBoxText == "") { System.Windows.MessageBox.Show("The object to remove is not selected."); return; }

            MessageBoxResult result =
            System.Windows.MessageBox.Show($"Are you shure to delete {cbSmtpSelect.ComboBoxText}?", $"Удаление: {cbSmtpSelect.ComboBoxText}", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (result.ToString() == "Yes") { VariableClass.Addreses.Remove(cbSmtpSelect.ComboBoxText); cbSmtpSelect.UpdateComboBox(); }
        }
        /// <summary>
        /// Change this smtp.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSmtpSelect_btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (cbSmtpSelect.ComboBoxText == "") { System.Windows.MessageBox.Show("The object to edit is not selected."); return; }

            EditSMTPWindow childWindow = new EditSMTPWindow();
            childWindow.Owner = this;
            childWindow.InitializeContent(cbSmtpSelect.ComboBoxText, (int)cbSmtpSelect.SelectedValue);
            VariableClass.Addreses.Remove(cbSmtpSelect.ComboBoxText);

            childWindow.ShowDialog();

            VariableClass.Addreses.Add(childWindow.Smtp, childWindow.Port);
            cbSmtpSelect.UpdateComboBox();
        }

        private void btnAddAddress_Click(object sender, RoutedEventArgs e)
        {
            AddAddressWindow childWindow = new AddAddressWindow();
            childWindow.Owner = this;
            childWindow.ShowDialog();

            db.MyEmails.Add(childWindow.result);
            dgEmails.UpdateLayout();
        }

        private void btnEditAddress_Сlick(object sender, RoutedEventArgs e)
        {
            AddAddressWindow childWindow = new AddAddressWindow();
            childWindow.Owner = this;
            childWindow.ShowDialog();

            //db.MyEmails.Add(childWindow.result);
        }

        private void btnDeleteAddress_Click(object sender, RoutedEventArgs e)
        {
            db.MyEmails.Remove((Email)dgEmails.SelectedItem);
        }
    }
}
