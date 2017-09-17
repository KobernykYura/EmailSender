using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfMailSender.SchedulerListView
{
    /// <summary>
    /// Логика взаимодействия для UserListControl.xaml
    /// </summary>
    public partial class UserListControl : UserControl
    {
        public event RoutedEventHandler btnMinusClick;
        public event RoutedEventHandler btnEditClick;

        private string mailText = "";

        private string tbTimePickerText = "";
        public string TbTimePickerText
        {
            get { return tbTimePicker.Text; }
            set
            {
                tbTimePickerText = value;
                tbTimePicker.Text = value;
            }
        }

        public string MailText
        {
            get
            {
                return mailText;
            }

            set
            {
                mailText = value;
            }
        }

        public UserListControl()
        {
            InitializeComponent();
        }

        private void btnEditSender_Click(object sender, RoutedEventArgs e)
        {
            btnMinusClick?.Invoke(sender, e);
        }

        private void btnDeleteSender_Click(object sender, RoutedEventArgs e)
        {
            btnEditClick?.Invoke(sender, e);
        }

        private void tbTimePicker_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {

        }

        //public TimeSpan GetSendTime(string strSendTime)
        //{
        //    TimeSpan tsSendTime = new TimeSpan();
        //    try
        //    {
        //        tsSendTime = TimeSpan.Parse(strSendTime);
        //    }
        //    catch { }

        //    return tsSendTime;
        //}
    }
}
