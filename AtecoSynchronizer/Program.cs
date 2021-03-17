using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;

namespace AtecoSynchronizer
{
    class Program
    {
        static void Main()
        {
            DirectoryInfo logs = new DirectoryInfo(Directory.GetCurrentDirectory()+"\\logs");
            if (!Directory.Exists(logs.ToString()))
                logs.Create();
            string IRP_path = Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["IRP"], AIRAP_path = Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["AIRAP"];
        }
    }
}
