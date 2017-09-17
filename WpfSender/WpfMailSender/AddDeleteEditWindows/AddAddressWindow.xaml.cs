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
using System.Windows.Shapes;

namespace WpfMailSender.AddDeleteEditWindows
{
    /// <summary>
    /// Логика взаимодействия для AddAddressWindow.xaml
    /// </summary>
    public partial class AddAddressWindow : Window
    {
        public Email result;

        public AddAddressWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                result.Name = tbAddress.Text;
                result.Value = tbSmtp.Text;
                result.Port = Convert.ToInt32(tbPort.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("Incorrect input. Please, try again.");
                return;
            }

            this.Close();
        }
    }
}
