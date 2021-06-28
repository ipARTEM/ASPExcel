using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASPExcel.Models
{
    public class ImportExcel
    {

        public ImportExcel()
        {
            PBs = new List<PhoneBrand>();

        }

        public List<PhoneBrand> PBs { get;  set; }

        public int ErrorsTotal { get; set; }
    }
}
