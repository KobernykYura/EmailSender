using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using WpfMailSender.AdditionalClasses;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfMailSender.AddDeleteEditWindows
{
    /// <summary>
    /// Логика взаимодействия для EditWindow.xaml
    /// </summary>
    public partial class EditWindow : Window
    {
        private string password;
        private string address;
        private int port;

        /// <summary>
        /// Property of password.
        /// </summary>
        public string Password { get { return password; } }
        /// <summary>
        /// Property of address.
        /// </summary>
        public string Address { get { return address; } }
        /// <summary>
        /// Property of port.
        /// </summary>
        public int Port { get { return port; } }

        public EditWindow()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            if (addressInput.Text == "" || pwInput.Text == "")
            {
                MessageBox.Show("Check address or password!\n It could't be empty.");
                return;  
            }
            else
            {
                password = pwInput.Text;
                address = addressInput.Text;
                this.Close();
            }

        }

        /// <summary>
        /// Initializing password and address for changing.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public void InitializeContent(string key, string value)
        {
            this.password = value;
            this.address = key;

            addressInput.Text = key;
            pwInput.Text = value;
        }
    }
}
