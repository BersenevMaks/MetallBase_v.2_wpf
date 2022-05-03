using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetallBase2.Classes
{
    public class OrganizationsViewModel
    {
        public List<string> Organizations { get; set; }
        public string OrgID { get; set; }
        public string OrgName { get; set; }
        public string OrgAddress { get; set; }
        public string OrgTel { get; set; }
        public string OrgEmail { get; set; }
        public string OrgSite { get; set; }
        public string OrgINN { get; set; }
        public string OrgRSchet { get; set; }
        public string OrgKorSchet { get; set; }
        public string OrgBIK { get; set; }
        public string OrgDatePrice { get; set; }

        public string OrgsCount { get; set; }

        public bool IsEnabledDelButton { get; set; } = false;
        public bool IsEnabledSaveButton { get; set; } = false;

    }
}
