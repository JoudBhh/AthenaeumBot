using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scraping_template_1_Thread1.Models
{
    public class Artist
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string DateBirth { get; set; }
        public string Nationality { get; set; }
        public string ArtworksNum { get; set; }
        public string ArtworksUrl { get; set; }
        public List<string> AllArrtworksUrls { get; set; }
        public string Url { get; set; }
  

    }
}
