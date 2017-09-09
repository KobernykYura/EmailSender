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

namespace WpfMailSender
{
    /// <summary>
    /// Логика взаимодействия для TabSwitcher.xaml
    /// </summary>
    public partial class TabSwitcherControl : UserControl
    {
        public event RoutedEventHandler btnNextClick;
        public event RoutedEventHandler btnPreviousClick;

        public TabSwitcherControl()
        {
            InitializeComponent();
        }


        #region properties
        private bool bHideBtnPrevious = false; // поле соответствующее будет ли скрыта кнопка Пердыдущий
        private bool bHideBtnNext = false; // поле соответствующее будет ли скрыта кнопка Следующий
        
        /// <summary>
        /// Свойство соответствующее будет ли скрыта кнопка Предыдущий. 
        /// Что бы Свойство отразилось на PropertiesGrid у нашего контрола, его нужно представить именно свойством, а не полем
        /// </summary>


        public bool IsHideBtnPrevious
        {
            get { return bHideBtnPrevious; }
            set
            {
                bHideBtnPrevious = value;
                SetButtons(); // метод, который отвечает на отрисовку кнопок
            }
        }// IsHideBtnPrevios

        public bool IsHideBtnNext
        {
            get { return bHideBtnNext; }
            set
            {
                bHideBtnNext = value;
                SetButtons(); // метод который отвечает за отрисовку кнопок
            }
        }// IsHideBtnPrevios
        private void BtnNextTrueBtnPreviosFalse()
        {
            btnNext.Visibility = Visibility.Hidden;
            btnPrevious.Visibility = Visibility.Visible;
            btnPrevious.Width = 229;
            btnNext.Width = 0;
            btnPrevious.HorizontalAlignment = HorizontalAlignment.Stretch;
        }
        private void BtnPreviosTrueBtnNextFalse()
        {
            btnPrevious.Visibility = Visibility.Hidden;
            btnNext.Visibility = Visibility.Visible;
            btnNext.Width = 229;
            btnPrevious.Width = 0;
            btnNext.HorizontalAlignment = HorizontalAlignment.Stretch;
        }
        private void BtnPreviosFalseBtnNextFalse()
        {
            btnNext.Visibility = Visibility.Visible;
            btnPrevious.Visibility = Visibility.Visible;
            btnNext.Width = 115;
            btnPrevious.Width = 114;
        }
        private void BtnPreviosTrueBtnNextTrue()
        {
            btnPrevious.Visibility = Visibility.Hidden;
            btnNext.Visibility = Visibility.Hidden;
        }
        /// <summary>
        /// Метод который отвечает за отрисовку кнопок.
        /// Есть четыре варианта, когда обе кнопки доступны. Доступна одна и не доступна вторая. И обе кнопки недоступны
        /// </summary>
        private void SetButtons()
        {
            if (bHideBtnPrevious && bHideBtnNext) BtnPreviosTrueBtnNextTrue();
            else if (!bHideBtnNext && !bHideBtnPrevious) BtnPreviosFalseBtnNextFalse();
            else if (bHideBtnNext && !bHideBtnPrevious) BtnNextTrueBtnPreviosFalse();
            else if (!bHideBtnNext && bHideBtnPrevious) BtnPreviosTrueBtnNextFalse();

        }


        #endregion

        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            btnPreviousClick?.Invoke(sender, e);
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            btnNextClick?.Invoke(sender, e);
        }
    }
}
