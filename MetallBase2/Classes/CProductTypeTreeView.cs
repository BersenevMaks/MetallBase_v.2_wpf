using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetallBase2.Classes
{
    class CProductTypeTreeView
    {
        public string TypeName { get; set; }
        public string ParentName { get; set; }
        public ObservableCollection<CProductTreeView> Products { get; set; }
        public CProductTypeTreeView()
        {
            Products = new ObservableCollection<CProductTreeView>();
        }
    }
}
