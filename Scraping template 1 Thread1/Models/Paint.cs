using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scraping_template_1_Thread1.Models
{
    public class Paint
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public string ArtistName { get; set; }
        public string ArtistUrl { get; set; }
        public string Copyright { get; set; }
        public string Tags { get; set; }
        public string Url { get; set; }
        public string UrlImages { get; set; }
        public Dictionary<string, string> Details { get; set; }
        public string Location { get; set; }
        public string Dates { get; set; }
        public string Dimensions { get; set; }
        public string Medium { get; set; }
        public string Enteredby { get; set; }
        public string Artistage { get; set; }




    }
}
