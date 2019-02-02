using MG2Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Client
{
    
    class Program
    {
        static private string user = ConfigurationManager.AppSettings["userName"];
        static private string passWord = ConfigurationManager.AppSettings["passWord"];
        static private string Token = ConfigurationManager.AppSettings["Token"];
        static private string websiteURL = "http://magento-211349-651516.cloudwaysapps.com";
        //static private string websiteURL = "http://www.joinseattle.com";

        static void Main(string[] args)
        {
           


           

        }

        static void GetToken(string userName,string passWord)
        {
            var m2 = new Magento(websiteURL);
            Token = m2.GetAdminToken(userName, passWord);
            Token  = Token.Replace("\"", "");
        }

        static void GetSku(string skuName,string token)
        {
            var magento = new Magento(websiteURL);
            magento.GetSku(token, skuName);
        }
        static void GetProductList(string token)
        {
            var magento = new Magento(websiteURL);
            magento.GetProductlist(token);
        }
        static void GetCategoryList(string token)
        {
            var magento = new Magento(websiteURL);
            //magento.GetCategoryList(token);
        }
        static void CreateCategory(string name, string token)
        {
            var magento = new Magento(websiteURL, token);
            magento.CreateCategory(name);
        }
    }
}
