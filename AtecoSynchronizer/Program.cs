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
        static void Main(string[] args)
        {
            string[] IRP = Directory.GetFiles(ConfigurationManager.AppSettings["IRP"]), AIRAP = Directory.GetFiles(ConfigurationManager.AppSettings["AIRAP"]);
        }
    }
}
