using System.Linq;

using Microsoft.SharePoint.Client;

namespace pjsummersjr.O365.Samples.Navigation
{
    class NavigationManager
    {
        private ClientContext spClient;
        private SharePointOnlineCredentials creds;
        public string SiteUrl { get; set; }
        public string UserName { get; set; }
        public System.Security.SecureString Password { get; set; }

        public NavigationManager(string siteUrl, string userName, System.Security.SecureString password)
        {
            SiteUrl = siteUrl;
            UserName = userName;
            Password = password;
        }

        public NavigationNodeCollection GetTopNavigation()
        {
            using (var ctx = GetSPContext())
            {
                NavigationNodeCollection navNodes = ctx.Web.Navigation.TopNavigationBar;
                ctx.Load(navNodes);
                ctx.ExecuteQuery();
                return navNodes;
            }
        }

        public NavigationNodeCollection AddNodeToNavigation(NavigationNodeCreationInformation node)
        {
            NavigationNodeCollection navNodes = GetTopNavigation();
            navNodes.Add(node);
            using (var ctx = GetSPContext())
            {
                ctx.Load(navNodes);
                ctx.ExecuteQuery();
            }
            return navNodes;
        }

        public NavigationNodeCollection RemoveNodeByTitle(string nodeTitle)
        {
            NavigationNodeCollection navNodes = GetTopNavigation();
            var deleteNodes = from node in navNodes where node.Title.Equals(nodeTitle) select node;
            var nodeToDelete = deleteNodes.First();
            nodeToDelete.DeleteObject();

            using (var ctx = GetSPContext())
            {
                ctx.Load(navNodes);
                ctx.ExecuteQuery();
            }

            return navNodes;
        }

        public NavigationNodeCollection AddNodeToNavigation(string Title, string link)
        {
            NavigationNodeCreationInformation navNodeDescriptor = new NavigationNodeCreationInformation();
            navNodeDescriptor.Title = Title;
            navNodeDescriptor.Url = link;
            return AddNodeToNavigation(navNodeDescriptor);
        }

        private SharePointOnlineCredentials GetCredentials()
        {
            if (creds == null)
            {
                creds = new SharePointOnlineCredentials(UserName, Password);
            }
            return creds;
        }

        private ClientContext GetSPContext()
        {
            if (spClient == null)
            {
                spClient = new ClientContext(SiteUrl);
                spClient.Credentials = GetCredentials();
            }

            return spClient;
        }
    }
}
