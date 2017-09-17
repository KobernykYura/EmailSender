using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CodePasswordDLL;

namespace WpfMailSender.AdditionalClasses
{
    public static class VariableClass
    {
        public static Dictionary<string, string> Senders
        {
            get { return dicSenders; }
        }
        private static Dictionary<string, string> dicSenders = new Dictionary<string, string>()
        {
            { "kobernicyri@mail.ru",CodePassword.getCodPassword("studentyura2014") },
            { "sokol777@mail.ru",CodePassword.getCodPassword("squerty") },
            { "sok74@yandex.ru",CodePassword.getCodPassword(";liq34tjk") }
        };

        public static Dictionary<string, int> Addreses
        {
            get { return dicAddresses; }
        }
        private static Dictionary<string, int> dicAddresses = new Dictionary<string, int>()
        {
            { "smtp.mail.ru", 25 },
            { "smtp.yandex.ru", 25 },
            { "smtp.gmail.com",465 }
        };

        public static List<int> Ports
        {
            get { return listPorts; }
        }
        private static List<int> listPorts = new List<int>()
        {
            { 25 },
            { 465 },
            { 81 },
            { 443 }

        };

    }
}
