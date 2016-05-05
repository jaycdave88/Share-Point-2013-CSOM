using System;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace SharPoint_Testing_POC
{
    class Program
    {
        //documentListName - is the Share Point site where the data is kept (i.e https://intranet.claritycon.com/Sites/DefaultCollection2013/Exiger/Insight%20to%20SharePoint%20Testing/Forms/AllItems.aspx)
        static string documentListName = "Insight to SharePoint Testing";

        //siteUrl - is the url for Share Point *NOTE: DO NOT TRY TO INPUT AN ENTIRE URL - it will crash with an error of no specific URL
        static string siteURL = "https://intranet.claritycon.com/Sites/DefaultCollection2013/Exiger/";

        //userNameFixed - is the account user name. (i.e jdave)
        static string userNameFixed = "";

        //passwordFixed - is the password for your username 
        static string passwordFixed = "";


        /// <summary>
        /// Main method to run all code, just uncomment the methods and update the data to alter the Share Point Site
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            
            //DeleteAFile("Test5", documentListName);
            ListAllFolders();
            //UploadFile(documentListName, "Test1.txt");

        }


        /// <summary>
        /// Will list all the Folders inside of a Site
        /// </summary>
        private static void ListAllFolders()
        {
            //Certificate Validation
            ServicePointManager.ServerCertificateValidationCallback +=
                delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };

            //Windows Authentication
            NetworkCredential networkCreds = new NetworkCredential(userNameFixed, passwordFixed, "clarityinternal");

            try
            {
                using (ClientContext clientContext = new ClientContext(siteURL))
                {
                    if (clientContext != null)
                    {
                        //Passing networkCredentials to create session
                        clientContext.Credentials = networkCreds;

                        //Indicats whether the client library needs to validate the method parameters on the client side.
                        clientContext.ValidateOnClient = true;

                        //Gets the website that is associated with the client context.
                        Web site = clientContext.Web;

                        //Represents a Folder on SharePoint Web Site
                        FolderCollection collFolder = site.Folders;

                        //Load the Folder
                        clientContext.Load(collFolder);

                        //Execute the Query aka. go fetch data.
                        clientContext.ExecuteQuery(); 

                        Console.WriteLine("The current site contains the following folders:\n\n");

                        foreach (Folder myFolder in collFolder)
                        {
                            Console.WriteLine("Name: {0} \n ", myFolder.Name);
                        }
                    }
                }
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e);
               
            }

        }

        /// <summary>
        /// Uploads the specified file to a SharePoint Site
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="listTitle"></param>
        /// <param name="fileName"></param>
        private static void UploadFile(string listTitle, string fileName)
        {
            //Certificate Validation
            ServicePointManager.ServerCertificateValidationCallback +=
                delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };

            //Windows Authentication
            NetworkCredential networkCreds = new NetworkCredential(userNameFixed, passwordFixed, "clarityinternal");

            try
            {
                using (ClientContext clientContext = new ClientContext(siteURL))
                {
                    if (clientContext != null)
                    {
                        //Passing networkCredentials to create session
                        clientContext.Credentials = networkCreds;

                        //Indicats whether the client library needs to validate the method parameters on the client side.
                        clientContext.ValidateOnClient = true;

                        //Opening FileStream to read the contents of a .txt and save it to SharePoint
                        using (var fileStream = new FileStream(fileName, FileMode.Open))
                        {
                            var fileInfo = new FileInfo(fileName);
                            var list = clientContext.Web.Lists.GetByTitle(listTitle);
                            clientContext.Load(list.RootFolder);
                            clientContext.ExecuteQuery();
                            var fileUrl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, fileInfo.Name);

                            Microsoft.SharePoint.Client.File.SaveBinaryDirect(clientContext, fileUrl, fileStream, true);
                        }
                    }
                }
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e);

            }
        }


        /// <summary>
        /// Will list out all the items within a Site, conduct a search and delete the item when found.
        /// </summary>
        /// <param name="sFileName"></param>
        /// <param name="sFldrLoc"></param>
        private static void DeleteAFile(string sFileName, string sFldrLoc)
        {
            //Certificate Validation
            ServicePointManager.ServerCertificateValidationCallback +=
                delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };

            //Windows Authentication
            NetworkCredential networkCreds = new NetworkCredential(userNameFixed, passwordFixed, "clarityinternal");

            try
            {
                using (ClientContext clientContext = new ClientContext(siteURL))
                {
                    if (clientContext != null)
                    {
                        //Passing networkCredentials to create session
                        clientContext.Credentials = networkCreds;

                        //Indicats whether the client library needs to validate the method parameters on the client side.
                        clientContext.ValidateOnClient = true;

                        //Gets the website that is associated with the client context.
                        Web web = clientContext.Web;

                        //Represents a collection of List objects that are the client context
                        ListCollection collList = web.Lists;

                        //Creating a List and passing in the documentListName var
                        List oList = collList.GetByTitle(sFldrLoc);

                        //Creating a new query
                        CamlQuery query = new CamlQuery();

                        //executing a query to list all objects within a SharePoint site.
                        query.ViewXml = "<View><Query><Where><Leq>" +
                             "<FieldRef Name='ID'/><Value Type='Number'>100</Value>" +
                             "</Leq></Where></Query><RowLimit>50</RowLimit></View>";

                        //Represents a collection of ListItems with the results from query
                        ListItemCollection collListItem = oList.GetItems(query);

                        clientContext.Load(collListItem,
                            items => items.IncludeWithDefaultProperties(
                                item => item.DisplayName));

                        //Execute 
                        clientContext.ExecuteQuery();

                        //Iterate over the collection
                        foreach (ListItem listitem in collListItem)
                        {
                            //conditional - find the listitem name that matchs file name.
                            if (listitem.DisplayName.Equals(sFileName))
                            {
                                //Delete object/file
                                listitem.DeleteObject();

                                //Execute and update
                                clientContext.ExecuteQuery();

                                Console.WriteLine("{0}, has been deleted sucessfully!", listitem.DisplayName);
                            }
                        }
                    }
                }
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e);
            }
        }
    }
}

