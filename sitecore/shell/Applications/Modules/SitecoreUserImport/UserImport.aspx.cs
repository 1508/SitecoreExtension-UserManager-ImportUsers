using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using Sitecore.Configuration;
using Sitecore.Install;
using Sitecore.Install.Security;
using Sitecore.Security.Accounts;
using Sitecore.Security.Domains;

namespace SitecoreExtension.SitecoreUserImporter
{
    public partial class UserImport : System.Web.UI.Page
    {
        private SortedList<string, UserData> usersByEmail = new SortedList<string, UserData>();
        private SortedList<string, UserData> usersByFullName = new SortedList<string, UserData>();
        private string userDomain
        {
            get
            {
                return txtUserDomain.Text;
            }
        }
        public string userImportFilePath
        {
            get 
            {
                return string.Format("{0}\\App_Data\\SitecoreUserImport\\SitecoreUserImport.csv", AppDomain.CurrentDomain.BaseDirectory);
            }
        }
        private string userExportFilePath
        {
            get 
            {
                return string.Format("{0}\\App_Data\\SitecoreUserImport\\SitecoreUserExport-{1}-package.zip", AppDomain.CurrentDomain.BaseDirectory, DateTime.Now.ToString("yyyy-M-d-hhmmss"));
            }
        }
        private string mailTemplateFile
        {
            get
            {
                return string.Format("{0}\\App_Data\\SitecoreUserImport\\MailTemplate.html", AppDomain.CurrentDomain.BaseDirectory);
            }
        }
        private bool replaceEmails
        {
            get { return this.chkReplaceEmails.Checked; }
        }
        private int maxImports
        {
            get
            {
                int i = 0;
                if (int.TryParse(txtMaxImports.Text, out i))
                {
                    return i;
                }

                return 0;
            }
        }

        private bool HasAccess
        {
            get
            {
                if (Sitecore.Context.User != null && Sitecore.Context.User.IsAuthenticated && Sitecore.Context.User.IsAdministrator)
                {
                    return true;
                }

                return false;
            }
        }

        private object CachedUsers
        {
            get { return ViewState["UserImport.UserList"]; }

            set { ViewState["UserImport.UserList"] = value; }
        }

        protected int UploadCount
        {
            get
            {
                if(CachedUsers != null && CachedUsers is List<UserData>)
                {
                    List<UserData> users = (List<UserData>)this.CachedUsers;
                    return users.Count;
                }
                return 0;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!this.HasAccess)
            {
                HttpContext.Current.Response.Redirect("/", true);                   
            }

            txtNotificationMailTemplate.TextMode = TextBoxMode.MultiLine;
            if(!this.IsPostBack)
            {
                FileInfo fi = new FileInfo(mailTemplateFile);
                if(fi.Exists)
                {
                    FileStream fs = fi.OpenRead();
                    int length = (int)fs.Length;
                    byte[] buffer = new byte[length];
                    fs.Read(buffer, 0, length);
                    fs.Close();

                    string mailTemplateHtml = Encoding.UTF8.GetString(buffer);
                    this.txtNotificationMailTemplate.Text = mailTemplateHtml;
                }

                chkReplaceEmails.Checked = true;
                txtMaxImports.Text = "20000";
            }
        }

        [Serializable]
        private class UserData
        {
            public string Name { get; set; }
            public string Email { get; set; }
            public string Login { get; set; }
            public string Password { get; set; }
            public string Website { get; set; }
            public string Description { get; set; }
            public List<string> Roles { get; set; }
            public string RawLine { get; set; }

            public UserData()
            {
                this.Name = this.Email = this.Login = this.RawLine = string.Empty;
                this.Roles = new List<string>();
            }

            public UserData(User user) : base()
            {
                this.Name = this.Email = this.Login = this.RawLine = string.Empty;
                this.Roles = new List<string>();

                this.Name = user.Profile.FullName;
                this.Email = user.Profile.Email;
                this.Login = user.LocalName;
                foreach(Role role in user.Roles)
                {
                    this.Roles.Add(role.LocalName);
                }
            }
        }

        protected void btnValidateRoles_Click(object sender, EventArgs e)
        {
            BuildUserList();

            List<string> rolesUsed = new List<string>();
            foreach (string key in usersByEmail.Keys)
            {
                UserData userData = usersByEmail[key];
                foreach(string role in userData.Roles)
                {
                    if(!rolesUsed.Contains(role))
                        rolesUsed.Add(role);
                }
            }

            rolesUsed.Sort();

            StringBuilder debug = new StringBuilder();
            debug.Append("<h1>Missing Roles</h1><ul>");
            foreach(string role in rolesUsed)
            {
                string fullRolename = GetFullName(role);
                try
                {
                    if (!Role.Exists(fullRolename))
                    {
                        debug.AppendFormat("<li>{0}</li>", fullRolename);
                    }
                }
                catch (Exception ex)
                {
                    debug.AppendFormat("<li>Invalid RoleName: {0} ({1})</li>", role, ex.Message);
                }
            }
            debug.Append("</ul>");

            divDebug.InnerHtml = debug.ToString();
        }

        private string GetFullName(string localName)
        {
            if(localName.StartsWith(this.userDomain+"\\"))
                return localName;

            return string.Format("{0}\\{1}", this.userDomain, localName);
        }

        private string GetShortName(string localName)
        {
            if (localName.StartsWith(this.userDomain + "\\"))
                return localName.Replace(this.userDomain + "\\", string.Empty);

            return localName;
        }

        private void BuildUserList(StringBuilder errors = null)
        {
            if(CachedUsers != null && CachedUsers is List<UserData>)
            {
                List<UserData> users = (List<UserData>) this.CachedUsers;
                foreach(UserData user in users)
                {
                    if(!this.usersByEmail.ContainsKey(user.Email))
                    {
                        this.usersByEmail.Add(user.Email, user);
                    }
                    string fullName = this.GetFullName(user.Login);
                    if(!this.usersByFullName.ContainsKey(fullName))
                    {
                        this.usersByFullName.Add(fullName, user);
                    }
                }

                return;
            }

            FileInfo fi = new FileInfo(userImportFilePath);
            FileStream fs = fi.OpenRead();
            int length = (int)fs.Length;
            byte[] buffer = new byte[length];
            fs.Read(buffer, 0, length);
            fs.Close();

            // 1252 = ANSI encoding
            string str = Encoding.GetEncoding(1252).GetString(buffer);

            BuildUserListFromCSV(str, errors);
        }

        private void BuildUserListFromCSV(string str, StringBuilder errors)
        {
            string[] lines = str.Split('\n');
            int lineCount = 0;
            foreach (string line in lines)
            {
                lineCount++;
                if (string.IsNullOrEmpty(line))
                {
                    if(errors != null)
                        errors.AppendFormat("Line {0}: Line was empty<br/>", lineCount);
                    continue;
                }

                string[] values = ParseLine(line);
                if (values == null || values.Count() < 5)
                {
                    if(errors != null)
                        errors.AppendFormat("Line {0}: Line was not in correct format [{1}]<br/>", lineCount, line);
                    continue;
                }

                string name = values[0];
                string email = values[1];
                string description = values[2];
                string roles = values[3];
                string website = values[4];
                string login = string.Empty;

                if (values.Count() > 5 && !string.IsNullOrWhiteSpace(values[5]))
                {
                    login = values[5];
                }
                else if(!string.IsNullOrEmpty(email) && email.Contains("@"))
                {
                    login = email.Substring(0, email.IndexOf("@"));
                }

                if (string.IsNullOrEmpty(email))
                {
                    if (errors != null)
                        errors.AppendFormat("Line: {0}: missing email (skipped): {1} - {2}</br>", lineCount, login, line);
                }

                if (!string.IsNullOrEmpty(email) && !string.IsNullOrEmpty(login))
                {
                    UserData userData = new UserData()
                                            {
                                                Name = name,
                                                Email = email,
                                                Login = login,
                                                Description = description,
                                                Roles = GetCsvList(roles),
                                                Website = website,
                                                RawLine = line
                                            };

                    if (!usersByEmail.ContainsKey(userData.Email))
                    {
                        string fullName = this.GetFullName(userData.Login);
                        if(!usersByFullName.ContainsKey(fullName))
                        {
                            usersByEmail.Add(userData.Email, userData);
                            usersByFullName.Add(this.GetFullName(userData.Login), userData);
                        }
                        else
                        {
                            if(errors != null)
                                errors.AppendFormat("Line: {0}: duplicate login {1} (login:{2}) line: {3}</br>", lineCount, fullName, login, line);
                        }
                    }
                    else
                    {
                        if (errors != null)
                            errors.AppendFormat("Line: {0}: duplicate email: {1} - {2}</br>", lineCount, login, line);
                    }
                }
            }
        }

        private List<string> GetCsvList(string csvList)
        {
            if(string.IsNullOrEmpty(csvList))
            {
                return new List<string>();
            }

            List<string> returnList = new List<string>();
            string[] values = csvList.Split(',');
            foreach(string value in values)
            {
                if(!string.IsNullOrEmpty(value.Trim()))
                    returnList.Add(value.Trim());
            }

            return returnList;
        }

        protected void btnTestUserImport_Click(object sender, EventArgs e)
        {
            BuildUserList();
            StringBuilder debug = new StringBuilder();
            debug.Append("<h1>Test-Importing Users</h1><ul>");
            int userCount = 1;
            foreach (string key in usersByEmail.Keys)
            {
                if (userCount > maxImports)
                {
                    break;
                }

                UserData userData = usersByEmail[key];
                ImportUser(userData, debug, userCount, true);
                userCount++;
            }
            debug.Append("</ul>");

            this.divDebug.InnerHtml = debug.ToString();
        }

        protected void btnShowExistingUsers_Click(object sender, EventArgs e)
        {
            StringBuilder debug = new StringBuilder();
            debug.Append("<h1>Existing Users</h1><ul>");
            var users = Domain.GetDomain(this.userDomain).GetUsers();
            foreach (User user in users)
            {
                UserData userData = new UserData(user);
                debug.Append(this.GetUserInfoListItem(userData, true));
            }
            debug.Append("</ul>");

            this.divDebug.InnerHtml = debug.ToString();
        }

        protected void btnImportUsers_Click(object sender, EventArgs e)
        {
            BuildUserList();
            StringBuilder debug = new StringBuilder();
            debug.Append("<h1>Importing Users</h1><ul>");
            int userCount = 1;
            foreach (string key in usersByEmail.Keys)
            {
                if (userCount > maxImports)
                {
                    break;
                }

                UserData userData = usersByEmail[key];
                ImportUser(userData, debug, userCount, false);
                userCount++;
            }
            debug.Append("</ul>");

            this.divDebug.InnerHtml = debug.ToString();
        }

        private void ImportUser(UserData userData, StringBuilder debug, int userCount, bool testrun)
        {
            string realEmail = userData.Email;
            if (replaceEmails)
            {
                userData.Email = "test" + userCount + "@1508.dk";
            }

            string userName = this.GetFullName(userData.Login);
            User user = null;
            bool userExists = false;
            if(!Sitecore.Security.Accounts.User.Exists(userName))
            {
                if(!testrun)
                {
                    // If the user does not exist, then create it
                    userData.Password = GeneratePassword() + "$";
                    user = global::Sitecore.Security.Accounts.User.Create(userName, userData.Password);
                    user.Profile.FullName = userData.Name;
                    user.Profile.Email = userData.Email;
                    user.Profile.Comment = userData.Description;
                    user.Profile.Save();
                }
            }
            else
            {
                user = global::Sitecore.Security.Accounts.User.FromName(userName, true);
                user.Profile.FullName = userData.Name;
                user.Profile.Email = userData.Email;
                user.Profile.Comment = userData.Description;
                user.Profile.Save();
                userExists = true;
            }

            if(!testrun)
            {
                // Attach all roles
                foreach(string role in userData.Roles)
                {
                    string fullRolename = GetFullName(role);
                    if(!Roles.IsUserInRole(user.Name, fullRolename) && Roles.RoleExists(fullRolename))
                    {
                        Roles.AddUserToRole(user.Name, fullRolename);
                    }
                }

                // Remove unknown roles
                foreach (string role in Roles.GetRolesForUser(user.Name))
                {
                    if (!userData.Roles.Contains(GetShortName(role)))
                    {
                        Roles.RemoveUserFromRole(user.Name, role);
                    }
                }
            }

            List<string> roles = new List<string>();
            roles.AddRange(userData.Roles.ToArray());
            debug.Append(this.GetUserInfoListItem(userData, userExists));

            userData.Email = realEmail;
        }

        protected void btnExportUsersInList_Click(object sender, EventArgs e)
        {
            this.BuildUserList();
            Sitecore.Data.Database db = Sitecore.Configuration.Factory.GetDatabase("master");
            Sitecore.Install.PackageProject document = new PackageProject();
            document.Metadata.PackageName = "SitecoreUserExport-" + DateTime.Now.ToString("yyyy-MM-dd-hhmmss");
            document.Metadata.Author = "Autogenerated by SitecoreExtension.SitecoreUserImporter";

            Sitecore.Install.Security.SecuritySource securitySource = new SecuritySource();
            securitySource.Name = "Users To Export";

            int count = 1;
            int packageCount = 0;
            foreach (string key in usersByEmail.Keys)
            {
                UserData userData = usersByEmail[key];

                string userName = GetFullName(userData.Login);

                if (Sitecore.Security.Accounts.User.Exists(userName))
                {
                    securitySource.AddAccount(new AccountString(userName, AccountType.User));
                    packageCount++;
                }

                count++;
            }

            document.Sources.Add(securitySource);
            document.SaveProject = true;

            if (packageCount > 0)
            {
                //path where the zip file package is saved
                using (Sitecore.Install.Zip.PackageWriter writer = new Sitecore.Install.Zip.PackageWriter(this.userExportFilePath))
                {
                    Sitecore.Context.SetActiveSite("shell");
                    writer.Initialize(Sitecore.Install.Installer.CreateInstallationContext());
                    Sitecore.Install.PackageGenerator.GeneratePackage(document, writer);
                    Sitecore.Context.SetActiveSite("website");
                }

                divDebug.InnerHtml = string.Format("<h1>{0} Users packaged, file location: </h1>{1}", packageCount, this.userExportFilePath);
            }
            else
            {
                divDebug.InnerHtml = string.Format("<h1>Could not find any users</h1>");
            }
        }

        protected void btnSendNotifications_Click(object sender, EventArgs e)
        {
            string html = this.txtNotificationMailTemplate.Text;
            if (string.IsNullOrEmpty(html))
            {
                divDebug.InnerHtml = "<h1>Error, no HTML Template defined</h1>";
                return;
            }

            string subject = RetrieveMetaValue(html, "MailSubject", "Account has been created");
            string mailSender = RetrieveMetaValue(html, "MailSender", string.Empty);

            if (string.IsNullOrEmpty(subject))
            {
                divDebug.InnerHtml = "<h1>Error, no mail subject has been defined check the necessary metatag has been added</h1>";
                return;
            }

            if (string.IsNullOrEmpty(mailSender))
            {
                divDebug.InnerHtml = "<h1>Error, no mail sender has been defined check the necessary metatag has been added</h1>";
                return;
            }

            BuildUserList();
            StringBuilder debug = new StringBuilder();
            StringBuilder errors = new StringBuilder();
            errors.Append("<h1>Failed to create password or send notification to</h1><ul>");
            debug.Append("<h1>Creating passwords and sending notifications to</h1></ul>");
            int userCount = 1;
            foreach (string key in usersByEmail.Keys)
            {
                if(userCount > maxImports)
                {
                    break;
                }

                UserData userData = usersByEmail[key];
                string fullName = GetFullName(userData.Login);
                if (Sitecore.Security.Accounts.User.Exists(fullName))
                {
                    User user = Sitecore.Security.Accounts.User.FromName(fullName, true);
                    if(!string.IsNullOrEmpty(user.Profile.Email))
                    {
                        MembershipUser membershipUser = Membership.GetUser(fullName, true);
                        if (membershipUser != null)
                        {
                            string newPassword = membershipUser.ResetPassword();
                            bool success = Membership.ValidateUser(fullName, newPassword);

                            string mailHtml = html.Replace("$name", user.Profile.FullName).Replace("$username", user.LocalName).Replace("$password", newPassword).Replace("$website", userData.Website);

                            try
                            {
                                SendMail(mailSender, user.Profile.Email, subject, mailHtml);
                                debug.Append(this.GetUserInfoListItem(userData));
                            }
                            catch (Exception)
                            {
                                errors.AppendFormat("<li>Error sending e-mail: {0}<ul>{1}</ul> </li>", user.Profile.Email, this.GetUserInfoListItem(userData));
                            }
                        }
                        else
                        {
                            errors.AppendFormat("<li>User lookup on {1} returned null: <ul>{0}</ul> </li>", this.GetUserInfoListItem(userData), fullName);        
                        }
                    }
                    else
                    {
                        errors.AppendFormat("<li>User does not have an e-mail address defined : <ul>{0}</ul> </li>", this.GetUserInfoListItem(userData));    
                    }
                }
                else
                {
                    errors.AppendFormat("<li>User does not exist: <ul>{0}</ul> </li>", this.GetUserInfoListItem(userData));    
                }

                userCount++;
            }

            errors.Append("</ul>");
            debug.Append("</ul>");

            divDebug.InnerHtml = string.Concat(errors.ToString(), debug.ToString());
        }

        public static void SendMail(string from, string to, string subject, string message)
        {
            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(from);
            mailMessage.To.Add(new MailAddress(to));
            mailMessage.Body = message;
            mailMessage.Subject = subject;
            mailMessage.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.Send(mailMessage);
            
            client.Dispose();
            mailMessage.Dispose();
        }


        private string RetrieveMetaValue(string html, string metadataName, string defaultValue)
        {
            Regex metaGrabber = new Regex(string.Format("<meta name=\"{0}\" content=\"([^\"]*)\" />", metadataName));
            Match match = metaGrabber.Match(html);
            if(match.Success && match.Groups.Count > 1)
            {
                return match.Groups[1].Value;
            }
            return defaultValue;
        }

        private string GetUserInfoListItem(UserData userData, bool? userExists = null)
        {
            List<string> roles = new List<string>();
            roles.AddRange(userData.Roles.ToArray());
            string output = string.Empty;
            if(userExists == null)
            {
                output = string.Format("<li>Email: {0} Name: {1} Login: {2} Password: {3} Description: {4} Roles: {5} </li>", userData.Email, userData.Name, userData.Login, userData.Password, userData.Description, string.Join(",", roles.ToArray()));
            }
            else
            {
                output = string.Format("<li>Email: {0} Name: {1} Login: {2} Password: {3} Description: {4} Roles: {5} User Exists: {6}</li>", userData.Email, userData.Name, userData.Login, userData.Password, userData.Description, string.Join(",", roles.ToArray()), userExists);
            }

            return output;
        }

        protected void btnBuildUserList_Click(object sender, EventArgs e)
        {
            StringBuilder errors = new StringBuilder();
            BuildUserList(errors);

            StringBuilder debug = new StringBuilder();
            debug.Append(errors.ToString());
            debug.Append("<h1>Users to import</h1><br/><ul>");
            foreach (string key in usersByEmail.Keys)
            {
                UserData userData = usersByEmail[key];
                List<string> roles = new List<string>();
                roles.AddRange(userData.Roles.ToArray());
                debug.Append(this.GetUserInfoListItem(userData));
            }
            debug.Append("</ul>");

            divDebug.InnerHtml = debug.ToString();            
        }

        protected void btnDeleteUsersInList_Click(object sender, EventArgs e)
        {
            BuildUserList();

            StringBuilder debug = new StringBuilder();
            debug.Append("<h1>Deleting users:</h1>");
            debug.Append("<ul>");

            int count = 1;
            foreach (string key in usersByEmail.Keys)
            {
                UserData userData = usersByEmail[key];
                if (count > maxImports)
                {
                    break;
                }

                string userName = GetFullName(userData.Login);

                if (Sitecore.Security.Accounts.User.Exists(userName))
                {
                    User user = Sitecore.Security.Accounts.User.FromName(userName, true);
                    if (user != null)
                    {
                        debug.Append(this.GetUserInfoListItem(userData));
                        user.Delete();
                    }
                }

                count++;
            }

            debug.Append("</ul>");

            divDebug.InnerHtml = debug.ToString();
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            StringBuilder errors = new StringBuilder();
            StringBuilder debug = new StringBuilder();
            if(this.fileImportFile.HasFile)
            {
                byte[] buffer = this.fileImportFile.FileBytes;

                // 1252 = ANSI encoding
                string str = Encoding.GetEncoding(1252).GetString(buffer);

                this.BuildUserListFromCSV(str, errors);
            }

            List<UserData> users = new List<UserData>();
            debug.Append(errors.ToString());
            debug.Append("<h1>Users uploaded</h1><br/><ul>");
            foreach (string key in usersByEmail.Keys)
            {
                UserData userData = usersByEmail[key];
                List<string> roles = new List<string>();
                roles.AddRange(userData.Roles.ToArray());
                debug.Append(this.GetUserInfoListItem(userData));
                users.Add(userData);
            }
            debug.Append("</ul>");

            this.CachedUsers = users;
            divDebug.InnerHtml = debug.ToString();            
        }



        private static string GeneratePassword(int lengt = 8, string validChars = "")
        {

            if (string.IsNullOrEmpty(validChars))
            {
                validChars = "abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ0123456789";
            }

            byte[] randomBytes = new byte[4];

            // Generate 4 random bytes.
            var rng = new RNGCryptoServiceProvider();
            rng.GetBytes(randomBytes);

            // Convert 4 bytes into a 32-bit integer value.
            int seed = (randomBytes[0] & 0x7f) << 24 |
                        randomBytes[1] << 16 |
                        randomBytes[2] << 8 |
                        randomBytes[3];

            var rnd = new Random(seed);

            var passwordAarray = new char[lengt];

            for (int i = 0; i < lengt; i++)
            {
                passwordAarray[i] = validChars[rnd.Next(0, validChars.Length)];
            }

            return string.Join("", passwordAarray);

        }

        private static string[] ParseLine(string line)
        {
            var values =
            new List<string>();

            string value = string.Empty;

            bool inQuote = false;
            for (int i = 0; i < line.Length; i++)
            {
                switch (line[i])
                {
                    case '"':
                        inQuote = !inQuote;
                        break;
                    case ',':
                        if (inQuote)
                        {
                            value += line[i];
                            break;
                        }
                        else
                        {
                            inQuote = false;
                            values.Add(value);
                            value = string.Empty;
                        }
                        break;
                    case ';':
                        if (inQuote)
                        {
                            value += line[i];
                            break;
                        }
                        else
                        {
                            inQuote = false;
                            values.Add(value);
                            value = string.Empty;
                        }
                        break;
                    default:
                        value += line[i];
                        break;
                }
            }
            values.Add(value);


            return values.ToArray();
        }
    }
}