using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.outlook.exporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var explorer = new OutlookExplorer();
            explorer.ReadAllMailItems();
        }
    }
}
