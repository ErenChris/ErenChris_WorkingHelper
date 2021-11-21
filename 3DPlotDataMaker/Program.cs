using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace _3DPlotDataMaker
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFilePath = @".\Input.txt";
            string outputFilePath = @".\Output.txt";
            string line = null;
            string[] temp = null;
            string result = null;

            try
            {
                using (StreamReader sr = new StreamReader(inputFilePath))
                {
                    while ((line = sr.ReadLine()) != null)
                    {
                        temp = Regex.Split(line, ": ");
                        if (temp[1].Contains("["))
                        {
                            result = result + "\n" + temp[1];
                        }
                        else
                        {
                            result += temp[1];
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            try
            {
                using(StreamWriter sw = new StreamWriter(outputFilePath))
                {
                    sw.Write(result);
                }
            }
            catch (Exception e)
            {

                throw;
            }

            Console.WriteLine("Done!");
            Console.ReadLine();
        }
    }
}
