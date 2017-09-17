using System;
using System.Windows;
using WpfMailSender.AdditionalClasses;

namespace WpfMailSender.AddDeleteEditWindows
{
    /// <summary>
    /// Логика взаимодействия для AddSMTPWindow.xaml
    /// </summary>
    public partial class AddSMTPWindow : Window
    {
        private int port;

        public int Port { get { return port; } }

        public AddSMTPWindow()
        {
            InitializeComponent();
            smtpInput.ItemsSource = VariableClass.Ports;
        }
        
        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            if (smtpInput.Text != "")
            {
                try
                {
                    port = Convert.ToInt32(smtpInput.Text);
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
    }
}
