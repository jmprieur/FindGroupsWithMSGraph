using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Linq;

namespace FindGroupsWithMSGraph
{
    static class Program
    {
        // Application Id as obtained by creating an application from https://apps.dev.microsoft.com
        // See also the guided setup:https://docs.microsoft.com/en-us/azure/active-directory/develop/guidedsetups/active-directory-windesktop
        private const string clientId = "[your_application_id_here]";


        /// <summary>
        /// Displays groups for a user and its organization
        /// WARNING: As of today, this application needs consent of an Azure AD administrator
        /// See https://developer.microsoft.com/en-us/graph/docs/concepts/known_issues
        /// </summary>
        /// <param name="args">unused</param>
        static void Main(string[] args)
        {
            // Get an access token
            PublicClientApplication app = new PublicClientApplication(clientId);
            string[] scopes = { "User.Read", "Group.Read.All", "Directory.Read.All", "Directory.AccessAsUser.All" };

            // Instanciate the Microsoft Graph, and provides the way to acquire the token.
            GraphServiceClient graph = new GraphServiceClient(new DelegateAuthenticationProvider(
             (requestMessage) =>
             {
                 AuthenticationResult result = null;
                 var u = app.Users.FirstOrDefault();

                 // If a user has already signed-in, we try first to acquire the token silently, and then if this fails
                 // we try to acquire it with a user interaction.
                 if (u != null)
                 {
                     try
                     {
                         result = app.AcquireTokenSilentAsync(scopes, app.Users.FirstOrDefault()).Result;
                     }
                     catch (MsalClientException ex)
                     {
                         if (ex.ErrorCode == "interaction_required")
                         {
                             result = app.AcquireTokenAsync(scopes).Result;
                         }
                     }
                 }
                 else
                 {
                     result = app.AcquireTokenAsync(scopes).Result;
                 }
                 requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", result.AccessToken);
                 return Task.FromResult(0);
             }));

            // Display the group Ids for all the group the signed-in user is part of.
            DisplayGroupIdsForGroupTheUserIsAMemberOf(graph).Wait();

            // Display all the groups in the organization of the signed-in user
            DisplayGroupIdsForAllGroupsInMyOrg(graph).Wait();

            // Display the group Ids for all the groups for a give user (here the signed-in user again)
            DisplayGroupIdsForUser(graph, graph.Me.Request().GetAsync().Result.UserPrincipalName).Wait();
        }

        /// <summary>
        /// For the signed-in user, displays on the standard output the groups the user is a member of
        /// </summary>
        /// <param name="graph">Graph</param>
        private static async Task DisplayGroupIdsForAllGroupsInMyOrg(IGraphServiceClient graph)
        {
            // All the groups in my organization
            var allGroupsRequest = graph.Groups.Request();
            while (allGroupsRequest != null)
            {
                IGraphServiceGroupsCollectionPage allGroups = await allGroupsRequest.GetAsync();
                foreach (Group group in allGroups)
                {
                    Console.WriteLine(group.Id);
                }
                allGroupsRequest = allGroups.NextPageRequest;
            }
        }

        /// <summary>
        /// Displays on the standard output all the groups in the organization of the signed-in user
        /// </summary>
        /// <param name="graph">Graph</param>
        private static async Task DisplayGroupIdsForGroupTheUserIsAMemberOf(IGraphServiceClient graph)
        {
            var myGroupsRequest = graph.Me.GetMemberGroups(false).Request();
            while (myGroupsRequest != null)
            {
                IDirectoryObjectGetMemberGroupsCollectionPage myGroups = await myGroupsRequest.PostAsync();
                foreach (string groupId in myGroups)
                {
                    Console.WriteLine(groupId);
                }
                myGroupsRequest = myGroups.NextPageRequest;
            }
        }

        /// <summary>
        /// Displays on the standard output all the groups in the organization of the signed-in user
        /// </summary>
        /// <param name="graph">Graph</param>
        private static async Task DisplayGroupIdsForUser(IGraphServiceClient graph, string userPrincipalName)
        {
            var allUserGroupRequest = graph.Users[userPrincipalName].MemberOf.Request();
            while (allUserGroupRequest != null)
            {
                var groupsPage = await allUserGroupRequest.GetAsync();
                foreach(var group in groupsPage)
                {
                    Console.WriteLine(group.Id);
                }
                allUserGroupRequest = groupsPage.NextPageRequest;
            }
        }
    }
}
