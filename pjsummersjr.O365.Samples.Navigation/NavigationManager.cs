using System.Linq;

using Microsoft.SharePoint.Client;

namespace pjsummersjr.O365.Samples.Navigation
{
    /// <summary>
    /// This class abstracts many of the operations required for managing navigation in a SharePoint Online site
    /// </summary>
    class NavigationManager
    {
        private ClientContext spClient;
        private SharePointOnlineCredentials creds;
        /// <summary>
        /// The SharePoint site where the navigation will be managed
        /// </summary>
        public string SiteUrl { get; set; }
        /// <summary>
        /// The username of a user with appropriate permissions for managing the navigation
        /// </summary>
        public string UserName { get; set; }
        /// <summary>
        /// The password of the above user
        /// </summary>
        public System.Security.SecureString Password { get; set; }

        public NavigationManager(string siteUrl, string userName, System.Security.SecureString password)
        {
            SiteUrl = siteUrl;
            UserName = userName;
            Password = password;
        }
        /// <summary>
        /// Returns the NavigationNodeCollection for the top navigation, also known as the Global Navigation
        /// </summary>
        /// <returns></returns>
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
        /// <summary>
        /// Adds a node to the navigation
        /// </summary>
        /// <param name="node">The parameters for the new navigation node encapsulated in a NavigationNodeCreationInformation object</param>
        /// <returns></returns>
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
        /// <summary>
        /// Adds a node to the navigation. This only requires the display value (Title) and link that the NavigationNode will
        /// follow.
        /// </summary>
        /// <param name="Title">Display value for the NavigationNode</param>
        /// <param name="link">URL to which the NavigationNode will link</param>
        /// <returns></returns>
        public NavigationNodeCollection AddNodeToNavigation(string Title, string link)
        {
            NavigationNodeCreationInformation navNodeDescriptor = new NavigationNodeCreationInformation();
            navNodeDescriptor.Title = Title;
            navNodeDescriptor.Url = link;
            return AddNodeToNavigation(navNodeDescriptor);
        }
        /// <summary>
        /// Removes a node from the navigation with the specified title. If multiple nodes have the same title, only the first will
        /// be removed so the command will have to be run multiple times. Requires an exact match.
        /// </summary>
        /// <param name="nodeTitle"></param>
        /// <returns></returns>
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
