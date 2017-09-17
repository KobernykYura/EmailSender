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
using System.Windows.Shapes;
using WpfMailSender.AdditionalClasses;

namespace WpfMailSender.AddDeleteEditWindows
{
    /// <summary>
    /// Логика взаимодействия для EditSMTPWindow.xaml
    /// </summary>
    public partial class EditSMTPWindow : Window
    {
        private int port;
        private string smtp; // Поля

        public int Port { get { return port; } } // Свойства
        public string Smtp { get { return smtp; } }

        public EditSMTPWindow()
        {
            InitializeComponent();
            portInput.ItemsSource = VariableClass.Ports;

        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            if (smtpInput.Text != "" || portInput.Text != "")
            {
                try
                {
                    port = Convert.ToInt32(portInput.Text);
                    smtp = smtpInput.Text;
                }
                catch (Exception)
                {
                    MessageBox.Show("Invalid port input!", "Warning!", MessageBoxButton.OK);
                    return;
                }
                this.Close();
            }
            else MessageBox.Show("Check smtp!\n It could be incorrect.");
        }

        public void InitializeContent(string smtp, int port)
        {
            this.port = port;
            this.smtp = smtp;

            smtpInput.Text = smtp;
            portInput.Text = port.ToString();
        }
    }
}
