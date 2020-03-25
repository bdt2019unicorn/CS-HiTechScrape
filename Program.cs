using System;
using System.Collections.Generic;
using System.Linq;
using HtmlAgilityPack;
using System.Data;
using ClosedXML.Excel;

namespace CS_HiTechScrape
{
    class Program
    {

        private static HtmlDocument GetDocument(string url)
        {
            HtmlWeb web = new HtmlWeb();
            HtmlDocument document = web.Load(url);
            return document; 
        }

        private static HtmlNodeCollection MainNodeSelection(HtmlDocument dom, string extra_path)
        {
            HtmlNode main_nodes = dom.DocumentNode.SelectSingleNode("//*[@id='app']/main");
            HtmlNodeCollection collection = main_nodes.SelectNodes("//*[@id='app']/main/div/div[@class='columns']/"+extra_path);
            return collection; 
        }
        static List<string> LoadYear()
        {
            List<string> years_url = new List<string>(); 
            string url = "https://www.hitech.org.nz/awards/finalists/";
            HtmlDocument years_dom = GetDocument(url);
            HtmlNodeCollection year_nodes = MainNodeSelection(years_dom, "a"); 
            foreach (HtmlNode a in year_nodes)
            {
                string href = a.Attributes["href"].Value;
                years_url.Add("https://www.hitech.org.nz" + href); 
            }
            return years_url; 
        }

        private static bool CheckNonExclude(string[]excludes, string value)
        {
            for (int i = 0; i < excludes.Length; i++)
            {
                if(value.IndexOf(excludes[i])>=0)
                {
                    return false; 
                }
            }

            return true; 
        }

        static HtmlNodeCollection LoadCategories(string url)
        {
            string[] excludes = { "Achiever", "Individual" }; 
            HtmlDocument year_dom = GetDocument(url);

            HtmlNodeCollection categories = MainNodeSelection(year_dom, "div[@class='winners-item']");
            List<HtmlNode> categories_removal = new List<HtmlNode>();
            foreach (HtmlNode category in categories)
            {
                HtmlNode tittle_div = category.SelectSingleNode("./div[1]/h4");
                if (!CheckNonExclude(excludes, tittle_div.InnerText))
                {
                    categories_removal.Add(category); 
                }
            }

            foreach (HtmlNode removal_item in categories_removal)
            {
                categories.Remove(removal_item); 
            }
            
            return categories; 
        }

        private static string CompanyLink(HtmlNode company)
        {
            string company_link = ""; 
            HtmlNodeCollection a = company.SelectNodes(".//a");
            try
            {
                foreach (HtmlNode link in a)
                {
                    string href = link.Attributes["href"].Value;
                    if (href.IndexOf("http") == 0)
                    {
                        company_link = href;
                        break;
                    }
                }
            }
            catch { }

            return company_link; 
        }

        private static void CombineDictionary(ref Dictionary<string,string>combine, Dictionary<string,string>new_values)
        {
            foreach (KeyValuePair<string,string> item in new_values)
            {
                if(!combine.Keys.Contains(item.Key))
                {
                    combine[item.Key] = item.Value; 
                }
            }
        }

        static Dictionary<string,string>NamesAndLink(HtmlNode category)
        {
            string[] excludes = { "University", "Datacom", "Trade Me", "Xero","Serko","Vodafone","IBM", "MYOB","Livestock Improvement Corporation", "Sharesies", "Fergus","Vend","Pushpay","EROAD","Movio" };

            System.Diagnostics.Debug.WriteLine(category.InnerText);

            HtmlNodeCollection all_potential = category.SelectNodes("./div[@class='editor-content']/p");
            if(all_potential==null)
            {
                all_potential = category.SelectNodes("./div[@class='editor-content']/a"); 
            }
            Dictionary<string, string> companies = new Dictionary<string, string>(); 
            try
            {
                foreach (HtmlNode company in all_potential)
                {
                    if (CheckNonExclude(excludes, company.InnerText))
                    {
                        string link = CompanyLink(company);
                        if (link != "")
                        {
                            companies[company.InnerText] = CompanyLink(company);
                        }
                    }
                }
            }
            catch { }


            return companies; 
        }

        static DataTable CompanyLinkTable(Dictionary<string,string>all_companies)
        {
            DataTable companies = new DataTable();
            companies.Columns.Add("Name");
            companies.Columns.Add("Link");
            foreach (KeyValuePair<string,string> company in all_companies)
            {
                DataRow company_table_row = companies.NewRow();
                company_table_row["Name"] = company.Key;
                company_table_row["Link"] = company.Value;
                companies.Rows.Add(company_table_row); 
            }
            return companies; 
        }

        static void ExportToExcel(DataTable all_companies_table)
        {
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(all_companies_table, "Companies");
            wb.SaveAs("Hi Tech Award.xlsx");
        }

        static void Main(string[] args)
        {
            List<string> years_urls = LoadYear();
            Dictionary<string, string> all_companies = new Dictionary<string, string>();
            foreach (string year_url in years_urls)
            {
                System.Diagnostics.Debug.WriteLine(year_url); 
                HtmlNodeCollection categories = LoadCategories(year_url);
                foreach (HtmlNode category in categories)
                {
                    Dictionary<string, string> companies = NamesAndLink(category);
                    CombineDictionary(ref all_companies, companies);
                }
            }

            DataTable all_companies_table = CompanyLinkTable(all_companies);
            ExportToExcel(all_companies_table);
            System.Diagnostics.Debug.WriteLine("The code has been done, please check it now");

            Console.WriteLine("The code has been done, press any key to completer"); 

            Console.ReadLine(); 
        }
    }
}
