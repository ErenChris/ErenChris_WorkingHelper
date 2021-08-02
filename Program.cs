using OpenQA.Selenium;
using System;
using WorkingHelper.Handler;
using WorkingHelper.Models;

namespace WorkingHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            //using (IWebDriver driver = new OpenQA.Selenium.Chrome.ChromeDriver())
            //{
            //    driver.Navigate().GoToUrl("http://www.baidu.com");  //driver.Url = "http://www.baidu.com"是一样的

            //    var source = driver.PageSource;

            //    Console.WriteLine(source);

            //    var byClassName = driver.FindElements(By.ClassName("text-color"));
            //    Console.WriteLine(byClassName);

            //    Console.ReadLine();
            //}
            string result = "";
            InformationForWebBrowser informationForWebBrowser = new InformationForWebBrowser("https://www.baidu.com");
            WebHandler webHandler = new WebHandler();
            result = webHandler.GetHtmlStr(informationForWebBrowser.Url, "UTF-8");

            Console.WriteLine(result);
            Console.ReadLine();
        }
    }
}
