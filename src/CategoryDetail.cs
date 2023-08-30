using Ganss.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MVPSpider
{
    public class CategoryDetail
    {
        public CategoryDetail() { 
            this.Names = new List<string>(); 
        }
        public void Increase() {
            Count++;
        }
        [Column(1)]
        public string Category { get; set; }
        [Column(2)]
        public int Count { get; set; }
        [Ganss.Excel.Ignore]
        public List<string> Names { get; set; }

        public override string ToString()
        {
            return String.Join(", ", Names.ToArray()); 
        }
    }
}
