using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfMailSender.DataBase;

namespace WpfMailSender
{       
    /// <summary>
    /// Класс который отвечает за работу с базой данных
    /// </summary>
    class DataBaseClass
    {
        private EmailsDataClassesDataContext emails = new EmailsDataClassesDataContext();
        public IQueryable<Emails> Emails
        {
            get
            {
                return from c in emails.Emails select c;
            }
        }
    }
}
