using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TravellerDataToMO
{
    class Traveller
    {
        public string Фамилия { get; set; }
        public string Имя { get; set; }
        public string Отчество { get; set; }
        public string ДатаРождения { get; set; }
        public string НР { get; set; }
        public string КодМедОрганизации { get; set; }


        public PropertyInfo[] ReturnType()
        {
            return this.GetType().GetProperties();
        }
    }
}
