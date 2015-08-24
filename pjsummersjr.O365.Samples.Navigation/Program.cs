using System;
using System.Configuration;
using System.Security;

using Microsoft.SharePoint.Client;

namespace pjsummersjr.O365.Samples.Navigation
{
    class Program
    {
        static void Main(string[] args)
        {
            string username = ConfigurationManager.AppSettings["O365User"];
            string password = ConfigurationManager.AppSettings["O365Password"];
            string siteUrl = ConfigurationManager.AppSettings["O365Site"];

            NavigationManager navMan = new NavigationManager(siteUrl, username, GetStringAsSecureString(password));

            Console.WriteLine("***************** Current Navigation **********************");
            NavigationNodeCollection navNodes = navMan.GetTopNavigation();
            foreach (var node in navNodes)
            {
                Console.WriteLine(node.Title);
            }
            Console.WriteLine("\n\n");

            Console.WriteLine("***************** Adding node - Title: \"Paul's Subsite\" *****************");
            navMan.AddNodeToNavigation("Paul's Subsite", "http://www.microsoft.com");
            Console.WriteLine("Added node successfully.");
            Console.WriteLine("\n\n");

            Console.WriteLine("***************** Print Updated Navigation **********************");
            NavigationNodeCollection newNavNodes = navMan.GetTopNavigation();
            foreach (var node in navNodes)
            {
                Console.WriteLine(node.Title);
            }
            Console.WriteLine("\n\n");

            Console.WriteLine("***************** Removing nodes containing - Title: \"Paul's Subsite\" *****************");
            navMan.RemoveNodeByTitle("Paul's Subsite");
            Console.WriteLine("Removed nodes successfully.");
            Console.WriteLine("\n\n");

            Console.WriteLine("***************** Print Updated Navigation **********************");
            newNavNodes = navMan.GetTopNavigation();
            foreach (var node in navNodes)
            {
                Console.WriteLine(node.Title);
            }
            Console.WriteLine("\n\n");

            Console.WriteLine("Demo complete. Press any key to quit...");
            Console.ReadLine();


        }

        private static SecureString GetStringAsSecureString(string str)
        {
            SecureString secStr = new SecureString();
            foreach (char c in str)
            {
                secStr.AppendChar(c);
            }

            return secStr;
        }
    }
}
