using System.Windows;
using System.Windows.Controls;

namespace WpfMailSender
{
    /// <summary>
    /// Логика взаимодействия для ComboBoxControll.xaml
    /// </summary>
    public partial class ComboBoxControl : UserControl
    {
        
        public event RoutedEventHandler btnAdd;
        public event RoutedEventHandler btnEdit; 
        public event RoutedEventHandler btnDelete;

        /// <summary>
        /// Label content. 
        /// </summary>
        public string Text
        {
            get { return lSender.Content.ToString(); }
            set { lSender.Content = value; }
        }
        /// <summary>
        /// Combo box Content.
        /// </summary>
        public string TextContent
        {
            get { return cbSelect.Text; }
            //set { cbSelect.Text = value; }
        }
        /// <summary>
        /// Combo box ItemsSource.
        /// </summary>
        public System.Collections.IEnumerable ItemsSource
        {
            get { return cbSelect.ItemsSource; }
            set { cbSelect.ItemsSource = value; }
        }
        /// <summary>
        /// Combo box DisplayMemberPath
        /// </summary>
        public string DisplayMemberPath
        {
            get { return cbSelect.DisplayMemberPath; }
            set { cbSelect.DisplayMemberPath = value; }
        }
        /// <summary>
        /// Combo box SelectedValuePath
        /// </summary>
        public string SelectedValuePath
        {
            get { return cbSelect.SelectedValuePath; }
            set { cbSelect.SelectedValuePath = value; }
        }
        /// <summary>
        /// Return selected object value key.
        /// </summary>
        public object SelectedValue
        {
            get { return cbSelect.SelectedValue; }
            //set { cbSelect.SelectedValue = value; }
        }
        /// <summary>
        /// Return selected object text.
        /// </summary>
        public string ComboBoxText
        { get { return cbSelect.Text; } }

        public ComboBoxControl()
        {
            InitializeComponent();
        }

        private void btnAddSender_Click(object sender, RoutedEventArgs e)
        {
            btnAdd?.Invoke(sender, e);
        }

        private void btnEditSender_Click(object sender, RoutedEventArgs e)
        {
            btnEdit?.Invoke(sender, e);
        }

        private void btnDeleteSender_Click(object sender, RoutedEventArgs e)
        {
            btnDelete?.Invoke(sender, e);
        }
        /// <summary>
        /// Update Combo box method.
        /// </summary>
         public void UpdateComboBox()
        {
            cbSelect.BeginInit();
            cbSelect.Text = "";
            cbSelect.EndInit();
        }
    }
}
