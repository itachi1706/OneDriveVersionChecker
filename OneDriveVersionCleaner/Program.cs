using System;
using System.Linq;
using System.Net;
using System.Security;
using System.Collections;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.IO;

namespace OneDriveVersionCleaner
{
    internal class Program
    {

        public static ArrayList versionedFiles = new ArrayList();

        private static void Main(string[] args)
        {

            // ------------------------------------------------------------------------------------------------------------
            // IMPORTANT: This is not production code, and it deletes files from your cloud storage. Use at your own risk!
            // ------------------------------------------------------------------------------------------------------------

            ExeConfigurationFileMap customConfigFileMap = new ExeConfigurationFileMap
            {
                ExeConfigFilename = "configuration.xml"
            };

            Configuration customConfig = ConfigurationManager.OpenMappedExeConfiguration(customConfigFileMap, ConfigurationUserLevel.None);

            AppSettingsSection appSettings = (customConfig.GetSection("appSettings") as AppSettingsSection);

            string webFullUrl = appSettings.Settings["WebUrl"].Value;
            string username = appSettings.Settings["Email"].Value;
            string pwstr = appSettings.Settings["Password"].Value;
            string rootRelativeUrl = appSettings.Settings["Root"].Value;
            string outFile = $@"{appSettings.Settings["OutputDir"].Value}";

            SecureString password = new NetworkCredential("", pwstr).SecurePassword;  // HACK :)

            WriteMessage($"Logging in as {username} to {webFullUrl} to look at {rootRelativeUrl} and output versioned files to {outFile}", ConsoleColor.White);

            versionedFiles.Add("Name,Path,Count,File Size");

            ClientContext context = new ClientContext(webFullUrl)
            {
                Credentials = new SharePointOnlineCredentials(username, password)
            };
            context.Load(context.Web);
            context.Load(context.Web.Lists);
            context.Load(context.Web, web => web.ServerRelativeUrl);
            context.ExecuteQuery();
            WriteMessage($"Connected to: {context.Web.ServerRelativeUrl}", ConsoleColor.Green);

            List list = context.Web.Lists.Single(l => l.Title == "Documents");
            context.Load(list);
            context.ExecuteQuery();
            WriteMessage($"Number of files: {list.ItemCount}", ConsoleColor.Green);

            WriteMessage($"\r\nFolder: {rootRelativeUrl}", ConsoleColor.Green);
            ProcessFolder(context, list, rootRelativeUrl);

            WriteMessage($"\r\nWriting list of versioned files to {outFile}", ConsoleColor.Green);
            System.IO.File.WriteAllLines(outFile, versionedFiles.Cast<string>());

            WriteMessage("\r\nDone.", ConsoleColor.Green);
        }

        private static void ProcessFolder(ClientContext context, List list, string rootRelativeUrl)
        {
            const int pageSize = 1000;

            Folder folder = context.Web.GetFolderByServerRelativeUrl(context.Web.ServerRelativeUrl + rootRelativeUrl);
            context.Load(folder);
            context.ExecuteQuery();

            CamlQuery query = new CamlQuery
            {
                ViewXml = $@"<View>
                       <RowLimit>{pageSize}</RowLimit>
                       <Query>
                         <Where>
                           <Eq>
                             <FieldRef Name='ContentType'/>
                             <Value Type='Computed'>Document</Value>
                           </Eq>
                         </Where>
                       </Query>
                     </View>",
                FolderServerRelativeUrl = folder.ServerRelativeUrl
            };

            bool hasMoreRecords = false;
            int pageCount = 1;

            do
            {
                WriteMessage($"\r\nPage: {pageCount}", ConsoleColor.White);
                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();
                WriteMessage($"File Count: {items.Count,30}", ConsoleColor.Cyan);

                ProcessItems(context, items, rootRelativeUrl);

                hasMoreRecords = items.ListItemCollectionPosition != null;
                query.ListItemCollectionPosition = items.ListItemCollectionPosition;

                pageCount++;
            } while (hasMoreRecords);

            //WriteMessage($"\r\nProcessing Subfolders...", ConsoleColor.White);
            ProcessSubFolders(context, list, folder, rootRelativeUrl);
        }

        // EDIT: process subfolders as well
        private static void ProcessSubFolders(ClientContext context, List list, Folder folder, string rootPath)
        {
            int pageSize = 1000;
            CamlQuery query = new CamlQuery
            {
                ViewXml = $@"<View>
                       <RowLimit>{pageSize}</RowLimit>
                       <Query>
                         <Where>
                           <Eq>
                             <FieldRef Name='ContentType'/>
                             <Value Type='Computed'>Folder</Value>
                           </Eq>
                         </Where>
                       </Query>
                     </View>",
                FolderServerRelativeUrl = folder.ServerRelativeUrl
            };

            bool hasMoreRecords = false;
            int pageCount = 1;

            do
            {
                //WriteMessage($"\r\nPage: {pageCount}", ConsoleColor.White);
                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();
                //WriteMessage($"Folder Count: {items.Count,30}", ConsoleColor.Cyan);

                foreach (ListItem item in items)
                {
                    context.Load(item);
                    context.ExecuteQuery();

                    Folder itemFolder = item.Folder;
                    if (itemFolder == null) continue;

                    context.Load(itemFolder);
                    context.ExecuteQuery();

                    string newPath = rootPath + "/" + itemFolder.Name;

                    WriteMessage($"Processing {newPath}", ConsoleColor.Yellow);

                    ProcessFolder(context, list, newPath);

                    // DEBUG
                    //ProcessSubFolders(context, list, itemFolder, newPath);
                }

                hasMoreRecords = items.ListItemCollectionPosition != null;
                query.ListItemCollectionPosition = items.ListItemCollectionPosition;

                pageCount++;
            } while (hasMoreRecords);
        }

        private static void ProcessItems(ClientContext context, ListItemCollection items, string parentPath)
        {
            foreach (ListItem item in items)
            {
                context.Load(item);
                context.ExecuteQuery();
                ProcessFile(context, item, parentPath);
            }
        }

        private static void ProcessFile(ClientContext context, ListItem item, string parentPath)
        {
            Microsoft.SharePoint.Client.File file = item.File;

            if (file != null)
            {
                context.Load(file);
                context.Load(file.Versions);
                context.ExecuteQuery();
                long fileSize = file.Length;
                int versionCount = file.Versions.Count;

                if (versionCount > 0)
                {
                    WriteMessage($"File: {file.Name,30}, Version count: {versionCount,5}, Size: {fileSize,8}, Deleting versions...", ConsoleColor.Gray);
                    versionedFiles.Add($"{file.Name},{parentPath}/{file.Name},{versionCount},{fileSize}");
                    // TODO: file.Versions.DeleteAll();
                }
                else
                {
                    // TODO: Uncomment for not empty
                    WriteMessage($"File: {file.Name,30}, Version count: {versionCount,5}", ConsoleColor.DarkGray);
                }
            }
        }

        private static void WriteMessage(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
