using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSUtil.Library;

namespace XLSUtil.Usage
{
    class Program
    {
        static void Main(string[] args)
        {
            var customerImport = new XLS2Model<CustomerModel>(@"C:\temp\test\customer.xlsx", s => s.Name, s => s.Category, s => s.LastAccess, s => s.Active);

            List<CustomerModel> customerList = customerImport.Get();

            if(customerList != null)
            {
                //Do stuff...
            }

        }
    }
}
