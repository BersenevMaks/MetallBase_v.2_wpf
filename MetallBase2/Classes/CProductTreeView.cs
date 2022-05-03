using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetallBase2.Classes
{
    class CProductTreeView
    {
        public string Name { get; set; }
        public string ParentName { get; set; }
        public ObservableCollection<CProductTypeTreeView> Types { get; set; }
        public CProductTreeView()
        {
            Types = new ObservableCollection<CProductTypeTreeView>();
        }
    }
}
