using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WorkingHelper.Handler
{
    class WebHandler
    {
        public void webClientMethod1()
        {
            WebClient wc = new WebClient();
            wc.Encoding = Encoding.UTF8;
            string html = wc.DownloadString("https://www.baidu.com/");

            MatchCollection matches = Regex.Matches(html, "<a.*>(.*)</a>");
            foreach (Match item in matches)
            {
                Console.WriteLine(item.Groups[1].Value);
            }
            Console.ReadKey();
        }

        public string SendRequest()
        {
            string url = "https://www.baidu.com/";
            Uri httpURL = new Uri(url);

            HttpWebRequest httpReq = (HttpWebRequest)WebRequest.Create(httpURL);

            HttpWebResponse httpResp = (HttpWebResponse)httpReq.GetResponse();

            System.IO.Stream respStream = httpResp.GetResponseStream();

            System.IO.StreamReader respStreamReader = new System.IO.StreamReader(respStream, Encoding.UTF8);
            string strBuff = respStreamReader.ReadToEnd();

            respStreamReader.Close();
            respStream.Close();
            return strBuff;
        }

        public string GetHtmlStr(string url, string encoding)
        {
            string htmlStr = "";
            try
            {
                if (!String.IsNullOrEmpty(url))
                {
                    WebRequest request = WebRequest.Create(url);
                    WebResponse response = request.GetResponse();
                    Stream datastream = response.GetResponseStream();
                    Encoding ec = Encoding.Default;
                    if (encoding == "UTF8")
                    {
                        ec = Encoding.UTF8;
                    }
                    else if (encoding == "Default")
                    {
                        ec = Encoding.Default;
                    }
                    StreamReader reader = new StreamReader(datastream, ec);
                    htmlStr = reader.ReadToEnd();
                    reader.Close();
                    datastream.Close();
                    response.Close();
                }
            }
            catch { }
            return htmlStr;
        }
    }
}
