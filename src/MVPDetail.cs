using Ganss.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MVPSpider
{
    public class MVPDetail
    {
        [Column(8)]
        public string Url { get; set; }
        [Column(9)]
        public string PhotoUrl { get; set; }
        [Column(3)]
        public string Category { get; set; }
        [Column(4)]
        public string TechFocus { get; set; }
        [Column(5)]
        [Column("Year in Program")]
        public string YearInProgram { get; set; }
        [Column(14)]
        public string Biography { get; set; }
        [Column(7)]
        public string Country { get; set; }

        [Column(1)]
        public string Name_En { get; set; }

        [Column(2)]
        public string Name_Cn { get; set; }
        [Column(6)]
        public string Gender { get; set; }
        [Column(10)]
        public string Social_Linkedin { get; set; }
        [Column(11)]
        public string Social_Twitter { get; set; }
        [Column(12)]
        public string Social_Github { get; set; }
        [Column(15)]
        public string Social_Blog { get; set; }
        [Column(13)]
        public string Social_Facebook { get; set; }
        [Column(14)]
        public string Social_Youtube { get; set; }
        [Column(16)]
        public string Social_Bilibili { get; set; }

    }
}
