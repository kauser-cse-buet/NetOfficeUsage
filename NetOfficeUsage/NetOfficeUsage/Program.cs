using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOfficeUsage
{
    class Program
    {
        static void Main(string[] args)
        {
            MSWordManager ms = new MSWordManager();
            ms.CreateNewDoc("NewDoc" + DateTime.Now.Ticks);
        }
    }
}
