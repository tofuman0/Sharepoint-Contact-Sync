using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PnP.Framework;
using System.DirectoryServices;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Windows.Forms;
using System.Text.Json;
using System.DirectoryServices.ActiveDirectory;
using System.Text.Encodings.Web;
using System.Text.Unicode;

namespace Sharepoint_Contact_Sync
{
    class Program
    {
        private static readonly byte[] key = { 204, 51, 158, 250, 255, 127, 27, 22, 228, 160, 17, 163, 242, 202, 121, 32, 195, 173, 57, 180, 241, 177, 140, 21, 128, 111, 226, 95, 215, 217, 22, 168 };

        private class SharepointFieldNames
        {
            public string? FirstName { get; set; }
            public string? LastName { get; set; }
            public string? JobTitle { get; set; }
            public string? OfficeNumber { get; set; }
            public string? MobileNumber { get; set; }
            public string? Extension { get; set; }
            public string? Department { get; set; }
        }
        private class JsonConfig
        {
            public string? SiteURL { get; set; }
            public string? ListName { get; set; }
            public string? ApplicationID { get; set; }
            public string? CertificatePath { get; set; }
            public string? CertificatePassword { get; set; }
            public string? TenantURL { get; set; }
            public string? DomainController { get; set; }
            public string? LDAPPath { get; set; }
            public string? LDAPQuery { get; set; }
            public string? CSVPath { get; set; }
            public string? DataFetchType { get; set; }
            public Int32? ExecuteLimit { get; set; }
            public Int32? RequestTimeout { get; set; }
            public bool? ClearListFirst { get; set; }
#nullable enable
            public SharepointFieldNames? SharepointFieldNames { get; set; }
#nullable disable
        }

        static private JsonConfig config = null;

        private struct UserDetails
        {
            public string FirstName;
            public string LastName;
            public string JobTitle;
            public string OfficeNumber;
            public string MobileNumber;
            public string Extension;
            public string Department;
            public UserDetails(string firstName, string lastName, string jobTitle, string officeNumber, string mobileNumber, string extension, string department)
            {
                FirstName = firstName;
                LastName = lastName;
                JobTitle = jobTitle;
                OfficeNumber = officeNumber;
                MobileNumber = mobileNumber;
                Extension = extension;
                Department = department;
            }
        };

        static Int32 Main(string[] args)
        {
            try
            {
                CheckConfig("contactsync.json");
                config = ReadConfig("contactsync.json");

                if(args.Count() > 0)
                {
                    for(Int32 i = 0; i < args.Count(); i++)
                    {
                        if (args[i].ToLower() == "-getfieldnames" || args[i].ToLower() == "/getfieldnames")
                        {
                            GetFieldNames();
                            return 0;
                        }
                        else if (args[i].ToLower() == "-cert" || args[i].ToLower() == "/cert")
                        {
                            if(i + 2 > args.Count())
                            {
                                ShowHelp();
                                return 1;
                            }
                            SetCert(args[i + 1], args[i + 2]);
                            Console.WriteLine("Updated configuration with certificate path and password.");
                            i += 2;
                        }
                        else if (args[i].ToLower() == "-defaultconfig" || args[i].ToLower() == "/defaultconfig")
                        {
                            GenerateDefaultConfig("contactsync.json");
                            Console.WriteLine("Default configuration created.");
                            return 0;
                        }
                        else if (args[i].ToLower() == "-?" || args[i].ToLower() == "/?")
                        {
                            ShowHelp();
                            return 0;
                        }
                        else
                        {
                            ShowHelp();
                            return 1;
                        }
                    }
                    return 0;
                }
                else
                {
                    AuthenticationManager authManager = new(config.ApplicationID, config.CertificatePath, UnprotectString(config.CertificatePassword), config.TenantURL);
                    using (var cc = authManager.GetContext(config.SiteURL))
                    {
                        Microsoft.SharePoint.Client.List oList = cc.Web.Lists.GetByTitle(config.ListName);
                        cc.RequestTimeout = Convert.ToInt32(config.RequestTimeout) * 1000;
                        cc.Load(oList);

                        if (config.ClearListFirst == true)
                        {
                            // Remove Rows
                            Microsoft.SharePoint.Client.ListItemCollection collListItem = oList.GetItems(Microsoft.SharePoint.Client.CamlQuery.CreateAllItemsQuery());
                            cc.Load(collListItem);
                            cc.ExecuteQuery();

                            if (RemoveAllContacts(collListItem) == true)
                            {
                                cc.ExecuteQuery();
                            }
                        }

                        // Add Rows
                        cc.Load(oList);
                        List<UserDetails> Users = GetUserDetails();
                        Int32 ExecuteCount = 0;
                        foreach (var User in Users)
                        {
                            AddContact(oList, User.FirstName, User.LastName, User.JobTitle, User.OfficeNumber, User.MobileNumber, User.Extension, User.Department);
                            ExecuteCount++;
                            if (ExecuteCount > config.ExecuteLimit)
                            {
                                cc.ExecuteQuery();
                                ExecuteCount = 0;
                            }
                        }
                        if (ExecuteCount > 0)
                            cc.ExecuteQuery();
                        return 0;
                    };
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
                return 1;
            }
        }

        private static void SetCert(string path, string password)
        {
            config.CertificatePath = path;
            config.CertificatePassword = ProtectString(password);
            WriteConfig(config, "contactsync.json");
        }

        private static void GetFieldNames()
        {
            AuthenticationManager authManager = new(config.ApplicationID, config.CertificatePath, UnprotectString(config.CertificatePassword), config.TenantURL);
            using (var cc = authManager.GetContext(config.SiteURL))
            {
                Microsoft.SharePoint.Client.List oList = cc.Web.Lists.GetByTitle(config.ListName);
                cc.RequestTimeout = Convert.ToInt32(config.RequestTimeout) * 1000;
                cc.Load(oList);

                // Query Rows
                FieldCollection oFieldCollection = oList.Fields;
                cc.Load(oFieldCollection);
                cc.ExecuteQuery();

                Int32 propertyLength = 0;
                // Get longest property name
                foreach (Field field in oFieldCollection)
                {
                    if (field.EntityPropertyName.Length > propertyLength)
                        propertyLength = field.EntityPropertyName.Length;
                }

                propertyLength = ((propertyLength / 4) + 2) * 4;

                {
                    string padding = "";
                    while (padding.Length + 4 < propertyLength)
                        padding += " ";
                    Console.WriteLine("Name" + padding + "Description");
                    Console.WriteLine("----" + padding + "-----------");
                }

                foreach (Field field in oFieldCollection)
                {
                    string padding = "";
                    while (padding.Length + field.EntityPropertyName.Length < propertyLength)
                        padding += " ";
                    Console.WriteLine(field.EntityPropertyName + padding + field.Title);
                }
            }
        }

        static bool RemoveAllContacts(Microsoft.SharePoint.Client.ListItemCollection list)
        {
            Console.WriteLine("Removing all contacts from list.");
            int listItemCount = list.Count;
            for (int i = 0; i < listItemCount; i++)
            {
                Console.Write("\rRemoving list line " + (i + 1) + "/" + listItemCount);
                list[0].DeleteObject();
            }
            Console.WriteLine("\nComplete.");
            return (listItemCount > 0);
        }

        static void AddContact(Microsoft.SharePoint.Client.List List, string FirstName, string LastName, string JobTitle, string OfficeNumber, string MobileNumber, string Extension, string Department)
        {
            Microsoft.SharePoint.Client.ListItemCreationInformation ListCreationInfo = new Microsoft.SharePoint.Client.ListItemCreationInformation();
            Microsoft.SharePoint.Client.ListItem Item = List.AddItem(ListCreationInfo);
            Item[config.SharepointFieldNames.FirstName] = FirstName;
            Item[config.SharepointFieldNames.LastName] = LastName;
            Item[config.SharepointFieldNames.JobTitle] = JobTitle;
            Item[config.SharepointFieldNames.OfficeNumber] = OfficeNumber;
            Item[config.SharepointFieldNames.MobileNumber] = MobileNumber;
            Item[config.SharepointFieldNames.Extension] = Extension;
            Item[config.SharepointFieldNames.Department] = Department;
            Item.Update();
        }

        static List<UserDetails> GetUserDetails()
        {
            List<UserDetails> userDetails = new List<UserDetails>();
            if (config.DataFetchType.ToLower() == "ldap")
            {
                Console.WriteLine("Retrieving user list from LDAP: " + config.DomainController + " - " + config.LDAPPath);
                DirectoryEntry ldapConnection = new DirectoryEntry(config.DomainController);
                ldapConnection.Path = "LDAP://" + config.LDAPPath;
                ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

                DirectorySearcher search = new DirectorySearcher(ldapConnection);
                search.Filter = config.LDAPQuery;

                // create an array of properties that we would like and  
                // add them to the search object  

                string[] requiredProperties = new string[] { "givenName", "sn", "title", "homePhone", "mobile", "ipPhone", "Department" };

                foreach (String property in requiredProperties)
                    search.PropertiesToLoad.Add(property);

                search.ServerTimeLimit = new TimeSpan(60);
                search.PageSize = 500;

                SearchResultCollection results = search.FindAll();
                Console.WriteLine("Number of LDAP results: " + results.Count);
                Console.WriteLine("Users without first and second names, without homephone, mobile and ipphone are ignored.");
                foreach (SearchResult result in results)
                {
                    DirectoryEntry user = result.GetDirectoryEntry();
                    userDetails.Add(new UserDetails(
                            Convert.ToString(user.Properties["givenName"].Value),
                            Convert.ToString(user.Properties["sn"].Value),
                            Convert.ToString(user.Properties["title"].Value),
                            Convert.ToString(user.Properties["homePhone"].Value),
                            Convert.ToString(user.Properties["mobile"].Value),
                            Convert.ToString(user.Properties["ipPhone"].Value),
                            Convert.ToString(user.Properties["Department"].Value)
                        )
                    );
                    Console.Write("\r" + userDetails.Count + "/" + results.Count);
                }
                Console.WriteLine("");
                ldapConnection.Dispose();
            }
            else
            {
                Console.WriteLine("Retrieving user list from CSV: " + config.CSVPath);
                var csvData = LoadCSV(config.CSVPath, ",");

                Console.WriteLine("Number of CSV results: " + csvData["Full Name"].Count);
                for (Int32 i = 0; i < csvData["Full Name"].Count; i++)
                {
                    var name = csvData["Full Name"][i].Split(" ");
                    if (name.Count() != 2)
                        continue;
                    userDetails.Add(new UserDetails(
                            name[0],
                            name[1],
                            csvData["Job Title"][i],
                            csvData["Phone No."][i],
                            csvData["Mobile No."][i],
                            csvData["Extension"][i],
                            csvData["Division"][i]
                        )
                    );
                }
            }
            return userDetails;
        }

        static Dictionary<string, List<string>> LoadCSV(string path, string delimiter)
        {
            Dictionary<string, List<string>> results = new Dictionary<string, List<string>>();

            using(var reader = new StreamReader(path))
            {
                bool valueNamesObtained = false;
                List<string> valueNames = new List<string>();
                Int32 rowCount = 0;
                while (!reader.EndOfStream)
                {
                    if(valueNamesObtained == false)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(delimiter);
                        foreach(var value in values)
                        {
                            valueNames.Add(value);
                            results.Add(value, new List<string>());
                        }
                        valueNamesObtained = true;
                    }
                    else
                    {
                        rowCount++;
                        var line = reader.ReadLine();
                        var values = line.Split(delimiter);
                        if(values.Length != valueNames.Count())
                        {
                            throw new Exception($"Values in {path} CSV row {rowCount} does not equal header count {valueNames.Count()}");
                        }
                        for(Int32 i = 0; i< values.Length; i++)
                        {
                            results[valueNames[i]].Add(values[i]);
                        }
                    }
                }
            }
            return results;
        }
        static void ShowHelp()
        {
            Console.WriteLine($"\nUsage: {Environment.ProcessPath.Split("\\").Last()} [-getfieldvalues] [-cert cert_path cert_password]\n");
            Console.WriteLine($"Options:\n");
            Console.WriteLine($"    -getfieldvalues     Returns with field values in sharepoint list.");
            Console.WriteLine($"    -cert               Sets the certificate to use and its password in configuration file.");
            Console.WriteLine($"    -defaultconfig      Replaces/Creates with default configuration file.\n");
            Console.WriteLine($"Note: Populate the SharepointFieldNames values in the configuration file from the values retrieved from -getfieldvalues.\n");
        }

        static string ProtectString(string input)
        {
            string result = "";
            List<byte> protectBuf = new List<byte>();
            Int32 keyOffset = 0;
            foreach(var character in input)
            {
                protectBuf.Add((byte)(Convert.ToByte((~Convert.ToByte(character & 0x0F) ^ 0xA7) & 0xFF) ^ key[keyOffset]));
                keyOffset = (keyOffset + 1) % key.Length;
                protectBuf.Add((byte)(Convert.ToByte((~Convert.ToByte(character & 0xF0) ^ 0xA7) & 0xFF) ^ key[keyOffset]));
                keyOffset = (keyOffset + 1) % key.Length;
            }
            result = Convert.ToHexString(protectBuf.ToArray());
            return result;
        }

        static string UnprotectString(string input)
        {
            string result = "";
            List<byte> unprotectBuf = new List<byte>();
            Int32 keyOffset = 0;
            for (Int32 i = 0; i < input.Length / 4; i++)
            {
                byte value1 = Convert.ToByte((~(Convert.ToByte(Convert.ToUInt32("0x" + input.Substring(i * 4, 2), 16) ^ 0xA7)) ^ key[keyOffset]) & 0xFF);
                keyOffset = (keyOffset + 1) % key.Length;
                byte value2 = Convert.ToByte((~(Convert.ToByte(Convert.ToUInt32("0x" + input.Substring((i * 4) + 2, 2), 16) ^ 0xA7)) ^ key[keyOffset]) & 0xFF);
                keyOffset = (keyOffset + 1) % key.Length;
                unprotectBuf.Add((byte)(value1 | value2));
            }
            result = System.Text.Encoding.Default.GetString(unprotectBuf.ToArray());
            return result;
        }

        static private void CheckConfig(string path)
        {
            if (System.IO.File.Exists(path) == true)
            {
                bool writeConfig = false;
                var jsonConfig = ReadConfig(path);
                if (jsonConfig.TenantURL == null)
                {
                    jsonConfig.TenantURL = "tenant.onmicrosoft.com";
                    writeConfig = true;
                }
                if (jsonConfig.ExecuteLimit == null || jsonConfig.ExecuteLimit == 0)
                {
                    jsonConfig.ExecuteLimit = 200;
                    writeConfig = true;
                }
                if (jsonConfig.ListName == null)
                {
                    jsonConfig.ListName = "List Name";
                    writeConfig = true;
                }
                if (jsonConfig.SiteURL == null)
                {
                    jsonConfig.SiteURL = "https://tenant.sharepoint.com/Sites/Site";
                    writeConfig = true;
                }
                if (jsonConfig.ApplicationID == null)
                {
                    jsonConfig.ApplicationID = "Obtain from 365 Azure portal";
                    writeConfig = true;
                }
                if (jsonConfig.CertificatePath == null)
                {
                    jsonConfig.CertificatePath = ".\\cert.pfx";
                    writeConfig = true;
                }
                if (jsonConfig.CertificatePassword == null)
                {
                    jsonConfig.CertificatePassword = "Run application with -cert switch to set certificate. Passwords are stored in reversable mechanisms so correct precautions to protect the configuration file is needed.";
                    writeConfig = true;
                }
                if (jsonConfig.CSVPath == null)
                {
                    jsonConfig.CSVPath = ".\\data.csv";
                    writeConfig = true;
                }
                if (jsonConfig.DataFetchType == null)
                {
                    jsonConfig.DataFetchType = "csv";
                    writeConfig = true;
                }
                if (jsonConfig.DomainController == null)
                {
                    jsonConfig.DomainController = "dc.domain.local";
                    writeConfig = true;
                }
                if (jsonConfig.LDAPPath == null)
                {
                    jsonConfig.LDAPPath = "CN=Users,DC=domain,DC=local";
                    writeConfig = true;
                }
                if(jsonConfig.LDAPQuery == null)
                {
                    jsonConfig.LDAPQuery = "(&(objectClass=User)(objectCategory=Person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(givenName=*)(sn=*)(|(homePhone=*)(mobile=*)(ipPhone=*)))";
                    writeConfig = true;
                }
                if (jsonConfig.RequestTimeout == null || jsonConfig.RequestTimeout == 0)
                {
                    jsonConfig.RequestTimeout = 60;
                    writeConfig = true;
                }
                if (jsonConfig.ClearListFirst == null)
                {
                    jsonConfig.ClearListFirst = true;
                    writeConfig = true;
                }
                if (jsonConfig.SharepointFieldNames == null)
                {
                    jsonConfig.SharepointFieldNames = new SharepointFieldNames { FirstName = "FirstName", LastName = "LastName", JobTitle = "JobTitle", OfficeNumber = "OfficeNumber", MobileNumber = "MobileNumber", Extension = "Extension", Department = "Department" };
                    writeConfig = true;
                }
                else
                {
                    if (jsonConfig.SharepointFieldNames.FirstName == null)
                    {
                        jsonConfig.SharepointFieldNames.FirstName = "FirstName";
                        writeConfig = true;
                    }
                    if (jsonConfig.SharepointFieldNames.LastName == null)
                    {
                        jsonConfig.SharepointFieldNames.LastName = "LastName";
                        writeConfig = true;
                    }
                    if (jsonConfig.SharepointFieldNames.OfficeNumber == null)
                    {
                        jsonConfig.SharepointFieldNames.OfficeNumber = "OfficeNumber";
                        writeConfig = true;
                    }
                    if (jsonConfig.SharepointFieldNames.MobileNumber == null)
                    {
                        jsonConfig.SharepointFieldNames.MobileNumber = "MobileNumber";
                        writeConfig = true;
                    }
                    if (jsonConfig.SharepointFieldNames.Extension == null)
                    {
                        jsonConfig.SharepointFieldNames.Extension = "Extension";
                        writeConfig = true;
                    }
                    if (jsonConfig.SharepointFieldNames.Department == null)
                    {
                        jsonConfig.SharepointFieldNames.Department = "Department";
                        writeConfig = true;
                    }
                    if (jsonConfig.SharepointFieldNames.JobTitle == null)
                    {
                        jsonConfig.SharepointFieldNames.JobTitle = "JobTitle";
                        writeConfig = true;
                    }
                }

                if (writeConfig == true)
                    WriteConfig(jsonConfig, path);
                return;
            }

            GenerateDefaultConfig(path);
        }

        static private void WriteConfig(JsonConfig config, string path)
        {
            try
            {
                string jsonDefault = JsonSerializer.Serialize(config, new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, WriteIndented = true });
                System.IO.File.WriteAllText(path, jsonDefault);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        static private JsonConfig ReadConfig(string path)
        {
            string jsonFile = System.IO.File.ReadAllText(path);
#nullable enable
            JsonConfig? jsonConfig = JsonSerializer.Deserialize<JsonConfig>(jsonFile);
#nullable disable

            return jsonConfig;
        }

        static private void GenerateDefaultConfig(string path)
        {
            var defaultConfig = new JsonConfig
            {
                SiteURL = "https://tenant.sharepoint.com/Sites/Site",
                ListName = "List Name",
                ApplicationID = "Obtain from 365 Azure portal",
                CertificatePath = ".\\cert.pfx",
                CertificatePassword = "Run application with -cert switch to set certificate. Passwords are stored in reversable mechanisms so correct precautions to protect the configuration file is needed.",
                TenantURL = "tenant.onmicrosoft.com",
                DomainController = "dc.domain.local",
                LDAPPath = "CN=Users,DC=Domain,DC=local",
                LDAPQuery = "(&(objectClass = User)(objectCategory = Person)(!userAccountControl:1.2.840.113556.1.4.803:= 2)(givenName = *)(sn = *)(| (homePhone = *)(mobile = *)(ipPhone = *)))",
                CSVPath = ".\\data.csv",
                DataFetchType = "csv",
                ExecuteLimit = 200,
                RequestTimeout = 60,
                ClearListFirst = true,
                SharepointFieldNames = new SharepointFieldNames { FirstName = "FirstName", LastName = "LastName", JobTitle = "JobTitle", OfficeNumber = "OfficeNumber", MobileNumber = "MobileNumber", Extension = "Extension", Department = "Department" }
            };

            WriteConfig(defaultConfig, path);
        }
    }
}
