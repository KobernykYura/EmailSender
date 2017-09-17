using Common;
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
    /// Логика взаимодействия для ListSchedulerView.xaml
    /// </summary>
    public partial class ListSchedulerView : UserControl
    {
        private DateTime dateTime;

        public event RoutedEventHandler btnAdd;
        public event RoutedEventHandler btnEdit;
        public event RoutedEventHandler btnDelete;

        /// <summary>
        /// DateTime Property
        /// </summary>
        public DateTime TimeData
        {
            get { return TimeData; }
        }
        /// <summary>
        /// Combo box ItemsSource.
        /// </summary>
        public System.Collections.IEnumerable ItemsSource
        {
            get { return lvListField.ItemsSource; }
            set { lvListField.ItemsSource = value; }
        }
        /// <summary>
        /// Combo box DisplayMemberPath
        /// </summary>
        public string DisplayMemberPath
        {
            get { return lvListField.DisplayMemberPath; }
            set { lvListField.DisplayMemberPath = value; }
        }
        /// <summary>
        /// Combo box SelectedValuePath
        /// </summary>
        public string SelectedValuePath
        {
            get { return lvListField.SelectedValuePath; }
            set { lvListField.SelectedValuePath = value; }
        }
        /// <summary>
        /// Return selected object value key.
        /// </summary>
        public object SelectedValue
        {
            get { return lvListField.SelectedValue; }
            //set { cbSelect.SelectedValue = value; }
        }

        public ListSchedulerView()
        {
            InitializeComponent();
        }

        private void btnAddDate_Click(object sender, RoutedEventArgs e)
        {
            lvListField.Items.Add(new UserListControl());
            UserListControl l = (UserListControl)lvListField.Items[lvListField.Items.Count - 1];
            l.btnMinusClick += (a, b) => { lvListField.Items.Remove(l); };
            l.btnEditClick += (a, b) => {
                TextWindow w = new TextWindow();
                w.Loaded += (a2, b2) => Data.MailText = l.MailText;
                w.Closed += (a1, b1) => l.MailText = Data.MailText;
                w.ShowDialog();
            };
        }

        private void btnDeleteDate_Click(object sender, RoutedEventArgs e)
        {
            btnDelete?.Invoke(sender, e);
        }

        private void btnEditMail_Click(object sender, RoutedEventArgs e)
        {
            btnEdit?.Invoke(sender, e);
        }
    }
}
