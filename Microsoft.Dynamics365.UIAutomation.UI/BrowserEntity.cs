using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.UI
{
    class BrowserEntity
    {
        public string id
        {
            get;
            set;
        }
        public string BrowserName
        {
            get;
            set;
        }
        public bool IsChecked
        {
            get { return _isChecked; }
            set { _isChecked = value; }
        }
        private bool _isChecked;

    }
}
