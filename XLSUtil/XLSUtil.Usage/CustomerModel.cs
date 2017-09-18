using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSUtil.Usage
{
    public class CustomerModel
    {
        public string Name { get; set; }
        public int Category { get; set; }
        public DateTime LastAccess { get; set; }
        public bool Active { get; set; }
    }
}
