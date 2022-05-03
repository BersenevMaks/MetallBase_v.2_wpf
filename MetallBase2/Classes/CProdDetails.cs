using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetallBase2.Classes
{
    public class CProdDetails
    {
        public string OrgName { get; set; }
        public string City { get; set; }
        public string Telephone { get; set; }
        public string Email { get; set; }
        public DataTable Managers { get; set; }
    }
}
