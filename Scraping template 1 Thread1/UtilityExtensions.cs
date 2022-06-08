using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scraping_template_1_Thread1
{
    public static class UtilityExtensions
    {
        public static void Save<T>(this List<T> items)
        {
            var name = typeof(T).Name;
            File.WriteAllText(name, JsonConvert.SerializeObject(items));
        }
    }
}
