using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkingHelper.Handler
{
    class ETextReader
    {
        public string filePath { get; set; }
        public string TextResult { get; set; }
        public StreamReader SR { get; set; }

        public ETextReader(string Path)
        {
            filePath = Path;
            try
            {
                SR = new StreamReader(filePath);
            }
            catch
            {
                Console.WriteLine("No such file!");
            }
        }

        public string GatTextFile()
        {
            while(!SR.EndOfStream)
            {
                string str = SR.ReadLine().Trim();
                TextResult += str;
            }

            return TextResult;
        }
    }
}
