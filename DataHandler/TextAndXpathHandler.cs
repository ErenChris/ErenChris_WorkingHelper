using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkingHelper.Handler
{
    public static class TextAndXpathHandler
    {
        public static string CharArrayToString(char[] cha, int len)
        {
            string str = "";

            for (int i = 0; i < len; i++)
            {
                str += string.Format("{0}", cha[i]);
            }

            return str;
        }

        public static string ChangeToChildXPath(string Xpath, int index)
        {
            char[] tempCharArray = new char[] { };
            tempCharArray = Xpath.ToCharArray();
            tempCharArray[Xpath.Length - 2] = char.Parse(index.ToString());
            string result = CharArrayToString(tempCharArray, Xpath.Length);

            return result;
        }
    }
}
