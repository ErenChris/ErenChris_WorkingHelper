using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkingHelper.Models
{
    class InformationForWebBrowser
    {
        public InformationForWebBrowser(string url)
        {
            this.Url = url;
        }

        public string Url { get; set; }
    }
}
