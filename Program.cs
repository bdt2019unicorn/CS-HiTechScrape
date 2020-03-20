using System;
using System.Collections.Generic; 
using HtmlAgilityPack; 

namespace CS_HiTechScrape
{
    class Program
    {
        static List<string> LoadYear()
        {
            List<string> years_url = new List<string>(); 
            string url = "https://www.hitech.org.nz/awards/finalists/"; 
            HtmlWeb web = new HtmlWeb();
            HtmlDocument years_dom = web.Load(url); 
            HtmlNode main_nodes = years_dom.DocumentNode.SelectSingleNode("//*[@id='app']/main");
            HtmlNodeCollection year_nodes = main_nodes.SelectNodes("//*[@id='app']/main/div/div[@class='columns']/a");
            foreach (HtmlNode a in year_nodes)
            {
                string href = a.Attributes.
            }
            return years_url; 
        }


        static void Main(string[] args)
        {
            List<string> years_urls = LoadYear(); 

            Console.ReadLine(); 
        }
    }
}
