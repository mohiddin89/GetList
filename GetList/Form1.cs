using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System.Globalization;
using System.Collections;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.Publishing;
using System.Xml;

namespace GetList
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int FoldCount = 0;
        int FileCount = 0;
        int ListFoldCount = 0;
        int ListFileCount = 0;
        public static Web web;

        public static string siteTitle;

        Dictionary<string, int> BuiltinGroups = new Dictionary<string, int>();
        Dictionary<string, int> ADGroups = new Dictionary<string, int>();
        List<string> lstADGroupsColl = new List<string>();
        private void button1_Click(object sender, EventArgs e)
        {

            #region AD Groups CSV Reading

            //lstADGroupsColl.Clear();

            //if (!string.IsNullOrEmpty(textBox3.Text))
            //{
            //    StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox3.Text));

            //    while (!sr.EndOfStream)
            //    {
            //        try
            //        {
            //            lstADGroupsColl.Add(sr.ReadLine().Trim().ToLower());
            //        }
            //        catch
            //        {
            //            continue;
            //        }
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please browse the path for ADGroups.csv");
            //}

            #endregion

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            //StreamWriter excelWriterScoringMatrixNew = null;

            //excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            //excelWriterScoringMatrixNew.WriteLine("Filename" + "," + "URL" + "," + "Owners" + "," + "Built-in-Groups" + "," + "AD Groups" + "," + "Start Time" + "," + "End Date" + "," + "Remarks");
            //excelWriterScoringMatrixNew.Flush();

            //List<string> ListNames = new List<string>();
            //ListNames.Add("Site Assets");
            //ListNames.Add("2_Documents and Pages");
            //ListNames.Add("1_Uploaded Files");
            //ListNames.Add("Discussions");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        foreach (List list in _Lists)
                        {
                            clientcontext.Load(list);
                            clientcontext.ExecuteQuery();

                            string listName = list.Title;

                            try
                            {
                                //bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, listName));

                                //if (_dListExist)
                                {
                                    if (listName == "Status")
                                    {
                                        // try
                                        {
                                            List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                            clientcontext.Load(Pagelist);
                                            clientcontext.ExecuteQuery();

                                            ViewCollection ViewColl = Pagelist.Views;
                                            clientcontext.Load(ViewColl);
                                            clientcontext.ExecuteQuery();

                                            Microsoft.SharePoint.Client.View v = ViewColl[0];
                                            clientcontext.Load(v);
                                            clientcontext.ExecuteQuery();

                                            //v.DeleteObject();
                                            //clientcontext.ExecuteQuery();

                                            v.ViewFields.RemoveAll();
                                            v.Update();
                                            clientcontext.ExecuteQuery();

                                            v.ViewFields.Add("StatusDescription");
                                            v.Update();
                                            clientcontext.ExecuteQuery();
                                        }
                                    }

                                    if ((list.BaseTemplate.ToString() == "109") && (listName != "Photos" || listName == "Images"))
                                    {
                                        #region Commented

                                        FieldCollection FldColl = list.Fields;
                                        clientcontext.Load(FldColl);
                                        clientcontext.ExecuteQuery();
                                        bool TagCateExist = false;

                                        foreach (Field tagField in FldColl)
                                        {
                                            clientcontext.Load(tagField);
                                            clientcontext.ExecuteQuery();

                                            if (tagField.Title.ToLower() == "tags" || tagField.Title.ToLower() == "categorization")
                                            {
                                                TagCateExist = true;
                                                break;
                                            }
                                        }

                                        #endregion

                                        if (TagCateExist)
                                        {
                                            List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                            clientcontext.Load(Pagelist);
                                            clientcontext.ExecuteQuery();

                                            ViewCollection ViewColl = Pagelist.Views;
                                            clientcontext.Load(ViewColl);
                                            clientcontext.ExecuteQuery();

                                            Microsoft.SharePoint.Client.View v = ViewColl[0];
                                            clientcontext.Load(v);
                                            clientcontext.ExecuteQuery();

                                            v.ViewFields.RemoveAll();
                                            v.Update();
                                            clientcontext.ExecuteQuery();

                                            v.ViewFields.Add("DocIcon");
                                            v.ViewFields.Add("Title");
                                            v.ViewFields.Add("LinkFilename");
                                            v.ViewFields.Add("Created");
                                            v.ViewFields.Add("Created By");
                                            v.ViewFields.Add("Modified");
                                            v.ViewFields.Add("Modified By");
                                            v.ViewFields.Add("Tags");
                                            v.ViewFields.Add("Categorization");
                                            v.Update();
                                            clientcontext.ExecuteQuery();
                                        }
                                    }

                                    #region Commented

                                    //else
                                    //{
                                    //    List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                    //    clientcontext.Load(Pagelist);
                                    //    clientcontext.ExecuteQuery();

                                    //    ViewCollection ViewColl = Pagelist.Views;
                                    //    clientcontext.Load(ViewColl);
                                    //    clientcontext.ExecuteQuery();

                                    //    Microsoft.SharePoint.Client.View v = ViewColl[0];
                                    //    clientcontext.Load(v);
                                    //    clientcontext.ExecuteQuery();

                                    //    v.ViewFields.RemoveAll();
                                    //    v.Update();
                                    //    clientcontext.ExecuteQuery();

                                    //    v.ViewFields.Add("DocIcon");
                                    //    v.ViewFields.Add("Title");
                                    //    v.ViewFields.Add("LinkFilename");
                                    //    v.ViewFields.Add("Created");
                                    //    v.ViewFields.Add("Created By");
                                    //    v.ViewFields.Add("Modified");
                                    //    v.ViewFields.Add("Modified By");
                                    //    v.ViewFields.Add("Tags");
                                    //    v.ViewFields.Add("Categorization");
                                    //    v.Update();
                                    //    clientcontext.ExecuteQuery();
                                    //}

                                    //Pagelist.ContentTypesEnabled = true;
                                    //Pagelist.Update();
                                    //clientcontext.ExecuteQuery(); 

                                    #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
                #region OLD
                //    this.Text = lstSiteColl[j] + "  Processing...";

                //    string startingTime = DateTime.Now.ToString();

                //    try
                //    {
                //        siteTitle = string.Empty;
                //        AuthenticationManager authManager = new AuthenticationManager();

                //        List<string> SPoExist = new List<string>();

                //        SPoExist.Add("spo.admin.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin2.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin3.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin4.verinon@agilent.onmicrosoft.com");
                //        SPoExist.Add("spo.admin5.verinon@agilent.onmicrosoft.com");

                //        string actualSPO = string.Empty;

                //        foreach (string sp in SPoExist)
                //        {
                //            try
                //            {
                //                using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), sp, "Lot62215"))
                //                {
                //                    clientcontext.Load(clientcontext.Web);
                //                    clientcontext.ExecuteQuery();

                //                    actualSPO = sp;

                //                    break;
                //                }
                //            }
                //            catch (Exception ex)
                //            {
                //                continue;
                //            }
                //        }

                //        //using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "adam.a@VerinonTechnology.onmicrosoft.com", "Lot62215##"))
                //        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), actualSPO, "Lot62215"))
                //        {


                //            ListCollection _Lists = clientcontext.Web.Lists;
                //            clientcontext.Load(_Lists);
                //            clientcontext.ExecuteQuery();

                //            bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(list => string.Equals(list.Title, "Site Assets"));

                //            if (_dListExist)
                //            {
                //                List Pagelist = clientcontext.Web.Lists.GetByTitle("Site Assets");
                //                clientcontext.Load(Pagelist);
                //                clientcontext.Load(Pagelist.RootFolder);
                //                clientcontext.ExecuteQuery();

                //                Pagelist.ContentTypesEnabled = true;
                //                Pagelist.Update();
                //                clientcontext.ExecuteQuery();
                //            }

                //            string admins = string.Empty;
                //            List<UserEntity> adminsColl = clientcontext.Site.RootWeb.GetAdministrators();

                //            foreach (UserEntity admin in adminsColl)
                //            {//SPO Admin 

                //                //User adUser = clientcontext.Site.RootWeb.SiteUsers.GetByLoginName(admin.LoginName);
                //                //adUser.is

                //                if (admin.Title != "FUN-SPO-SITECOLL-ADMINS" && (!admin.Title.ToLower().Contains("spo admin")) && (!admin.Email.ToLower().Contains("spo.admin@agilent.onmicrosoft.com")))
                //                {
                //                    if (!string.IsNullOrEmpty(admin.Email))
                //                    {
                //                        admins += admin.Email + ";";
                //                    }
                //                    else
                //                    {
                //                        admins += admin.Title + ";";
                //                    }
                //                }
                //            }

                //            Web oWebcurr = clientcontext.Site.RootWeb;
                //            clientcontext.Load(oWebcurr);
                //            clientcontext.ExecuteQuery();

                //            BuiltinGroups.Clear();
                //            ADGroups.Clear();

                //            siteTitle = oWebcurr.Title;

                //            string siteCollName = siteTitle.Replace(" ", "_");

                //            siteCollName = siteCollName.Replace("//", "_");

                //            string siteCollNameFileName = string.Empty;

                //            StreamWriter excelWriterScoringNew = null;

                //            if (!string.IsNullOrEmpty(siteTitle))
                //            {
                //                siteCollNameFileName = siteCollName;
                //                excelWriterScoringNew = System.IO.File.CreateText(textBox2.Text + "\\" + siteCollName + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
                //            }
                //            else
                //            {
                //                string[] siteCollNameFileNameXX = lstSiteColl[j].ToString().Trim().Split(new char[] { '/' });

                //                string actName = siteCollNameFileNameXX[siteCollNameFileNameXX.Length - 1];

                //                actName = actName.Replace(" ", "_");
                //                actName = actName.Replace("\\", "_");

                //                siteCollNameFileName = actName;
                //                excelWriterScoringNew = System.IO.File.CreateText(textBox2.Text + "\\" + actName + "_Report_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
                //            }

                //            excelWriterScoringNew.WriteLine("Site Coll Owners" + "," + admins + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "");
                //            excelWriterScoringNew.Flush();


                //            excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "Group" + "," + "Given though" + "," + "Folders" + "," + "Files" + "," + "Design" + "," + "Contribute" + "," + "Read" + "," + "Full Control" + "," + "Edit" + "," + "View Only" + "," + "Approve" + "," + "Contribute Limited" + "," + "OtherPermissions");
                //            excelWriterScoringNew.Flush();

                //            //////excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "SiteCollection Owners" + "," + "AD Group/Everyone granted directly" + "," + "Granted directly/added inside SP-Group" + "," + "Total number of Folders" + "," + "Total number of Files" + "," + "Design" + "," + "Contribute" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design" + "," + "Design");
                //            //excelWriterScoringNew.WriteLine("Object Type" + "," + "URL" + "," + "SiteCollection Owners" + "," + "AD Group/Everyone granted directly" + "," + "Granted directly/added inside SP-Group" + "," + "Total number of Folders" + "," + "Total number of Files" + "," + "Design" + "," + "Contribute" + "," + "Read" + "," + "Full Control" + "," + "Edit" + "," + "View Only" + "," + "Approve" + "," + "Contribute Limited" + "," + "OtherPermissions");
                //            //excelWriterScoringNew.Flush();

                //            #region Site Coll

                //            RoleAssignmentCollection webRoleAssignments = null;
                //            GroupCollection webGroups = null;

                //            try
                //            {
                //                webRoleAssignments = clientcontext.Web.RoleAssignments;
                //                clientcontext.Load(webRoleAssignments);
                //                clientcontext.ExecuteQuery();

                //                clientcontext.Load(clientcontext.Web);
                //                clientcontext.ExecuteQuery();

                //                webGroups = clientcontext.Web.SiteGroups;
                //                clientcontext.Load(webGroups);
                //                clientcontext.ExecuteQuery();

                //                bool foundatSiteLevel = false;

                //                string AdGroupsinGroup = string.Empty;
                //                string AdGroupsatSite = string.Empty;

                //                foreach (RoleAssignment member1 in webRoleAssignments)
                //                { //c:0u.c|tenant|                             

                //                    try
                //                    {
                //                        //if (!foundatSiteLevel)
                //                        //{
                //                        clientcontext.Load(member1.Member);
                //                        clientcontext.ExecuteQuery();

                //                        if (member1.Member.Title.Contains("c:0u.c|tenant|"))
                //                        {
                //                            continue;
                //                        }

                //                        #region Role Definations

                //                        RoleDefinitionBindingCollection rdefColl = member1.RoleDefinitionBindings;
                //                        clientcontext.Load(rdefColl);
                //                        clientcontext.ExecuteQuery();

                //                        string Design = string.Empty;
                //                        string Contribute = string.Empty;
                //                        string Read = string.Empty;
                //                        string FullControl = string.Empty;
                //                        string Edit = string.Empty;
                //                        string ViewOnly = string.Empty;
                //                        string Approve = string.Empty;
                //                        string ContributeLimited = string.Empty;
                //                        string OtherPermissions = string.Empty;

                //                        foreach (RoleDefinition rdef in rdefColl)
                //                        {
                //                            clientcontext.Load(rdef);
                //                            clientcontext.ExecuteQuery();

                //                            switch (rdef.Name)
                //                            {
                //                                case "Design":
                //                                    Design = "Yes";
                //                                    break;

                //                                case "Contribute":
                //                                    Contribute = "Yes";
                //                                    break;

                //                                case "Read":
                //                                    Read = "Yes";
                //                                    break;

                //                                case "Full Control":
                //                                    FullControl = "Yes";
                //                                    break;

                //                                case "Edit":
                //                                    Edit = "Yes";
                //                                    break;

                //                                case "View Only":
                //                                    ViewOnly = "Yes";
                //                                    break;

                //                                case "Contribute Limited":
                //                                    ContributeLimited = "Yes";
                //                                    break;

                //                                case "Approve":
                //                                    Approve = "Yes";
                //                                    break;

                //                                default:
                //                                    OtherPermissions = rdef.Name;
                //                                    break;
                //                            }
                //                        }

                //                        #endregion

                //                        if (member1.Member.PrincipalType == PrincipalType.SharePointGroup)
                //                        {
                //                            Group ouserGroup = (Group)member1.Member.TypedObject;
                //                            clientcontext.Load(ouserGroup);
                //                            clientcontext.ExecuteQuery();

                //                            UserCollection userColl = ouserGroup.Users;
                //                            clientcontext.Load(userColl);
                //                            clientcontext.ExecuteQuery();

                //                            foreach (User xUser in userColl)
                //                            {
                //                                if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                //                                {
                //                                    //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "--" + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                                    //AdGroupsinGroup += ouserGroup.Title + "; ";
                //                                    //break;
                //                                    if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                //                                    {
                //                                        if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                        {
                //                                            if (xUser.Title.ToString().ToLower().Contains("everyone"))
                //                                            {
                //                                                //if (!BuiltinGroups.Contains(xUser.Title))
                //                                                //{
                //                                                //    BuiltinGroups.Add(xUser.Title);
                //                                                //}

                //                                                if (BuiltinGroups.ContainsKey(xUser.Title))
                //                                                {
                //                                                    BuiltinGroups[xUser.Title]++;
                //                                                }
                //                                                else
                //                                                {
                //                                                    BuiltinGroups.Add(xUser.Title, 1);
                //                                                }
                //                                            }
                //                                            else
                //                                            {
                //                                                //if (!ADGroups.Contains(xUser.Title))
                //                                                //{
                //                                                //    ADGroups.Add(xUser.Title);
                //                                                //}

                //                                                if (ADGroups.ContainsKey(xUser.Title))
                //                                                {
                //                                                    ADGroups[xUser.Title]++;
                //                                                }
                //                                                else
                //                                                {
                //                                                    ADGroups.Add(xUser.Title, 1);
                //                                                }
                //                                            }

                //                                            excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                            excelWriterScoringNew.Flush();
                //                                        }
                //                                    }
                //                                    //foundatSiteLevel = true;
                //                                    //break;
                //                                }

                //                                //if (xUser.Title == "Everyone except external users")
                //                                //{
                //                                //    excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");

                //                                //    foundatSiteLevel = true;
                //                                //    break;
                //                                //}
                //                            }
                //                        }
                //                        if (member1.Member.PrincipalType == PrincipalType.SecurityGroup)
                //                        {
                //                            //if (member1.Member.Title == "Everyone except external users")
                //                            //{
                //                            //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + member1.Member.Title + "\"" + "," + "\"" + "--" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                            //AdGroupsatSite += member1.Member.Title + "; ";

                //                            if (lstADGroupsColl.Contains(member1.Member.Title.ToString().Trim().ToLower()))
                //                            {
                //                                if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                {
                //                                    if (member1.Member.Title.ToString().ToLower().Contains("everyone"))
                //                                    {
                //                                        //if (!BuiltinGroups.Contains(member1.Member.Title))
                //                                        //{
                //                                        //    BuiltinGroups.Add(member1.Member.Title);
                //                                        //}

                //                                        if (BuiltinGroups.ContainsKey(member1.Member.Title))
                //                                        {
                //                                            BuiltinGroups[member1.Member.Title]++;
                //                                        }
                //                                        else
                //                                        {
                //                                            BuiltinGroups.Add(member1.Member.Title, 1);
                //                                        }
                //                                    }
                //                                    else
                //                                    {
                //                                        //if (!ADGroups.Contains(member1.Member.Title))
                //                                        //{
                //                                        //    ADGroups.Add(member1.Member.Title);
                //                                        //}
                //                                        if (ADGroups.ContainsKey(member1.Member.Title))
                //                                        {
                //                                            ADGroups[member1.Member.Title]++;
                //                                        }
                //                                        else
                //                                        {
                //                                            ADGroups.Add(member1.Member.Title, 1);
                //                                        }
                //                                    }

                //                                    excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + member1.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                    excelWriterScoringNew.Flush();
                //                                }
                //                            }
                //                            //foundatSiteLevel = true;
                //                            //break;
                //                            //}
                //                        }

                //                        #region Commented Is Uesr

                //                        //if (member1.Member.PrincipalType == PrincipalType.User)
                //                        //{
                //                        //    if (member1.Member.Title == "Everyone except external users")
                //                        //    {
                //                        //        excelWriterScoringNew.WriteLine("\"" + "Site" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                        //        foundatSiteLevel = true;
                //                        //        break;
                //                        //    }
                //                        //} 

                //                        #endregion
                //                        //}
                //                        //else
                //                        //{
                //                        //    break;
                //                        //}
                //                    }
                //                    catch (Exception ex)
                //                    {
                //                        continue;
                //                    }
                //                }
                //                //excelWriterScoringNew.WriteLine("\"" + "SiteCollection" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSite + "\"" + "," + "\"" + AdGroupsinGroup + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                //                //excelWriterScoringNew.Flush();
                //            }
                //            catch (Exception ex)
                //            {
                //            }

                //            #endregion

                //            #region Lists

                //            ListCollection olistColl = clientcontext.Web.Lists;
                //            clientcontext.Load(olistColl);
                //            clientcontext.ExecuteQuery();

                //            foreach (List oList in olistColl)
                //            {
                //                bool foundatListLevel = false;

                //                clientcontext.Load(oList);
                //                clientcontext.Load(oList, li => li.HasUniqueRoleAssignments);
                //                clientcontext.ExecuteQuery();

                //                if (oList.BaseType == BaseType.DocumentLibrary)
                //                {
                //                    if (oList.Title == "Documents")
                //                    {
                //                        bool foXXXXundatListLevel = false;
                //                    }

                //                    if ((oList.Title != "Form Templates" && oList.Title != "Site Assets" && oList.Title != "SitePages" && oList.Title != "Style Library" && oList.Hidden == false && oList.IsCatalog == false && oList.BaseTemplate == 101) || oList.BaseTemplate == 700)
                //                    {
                //                        string UniqueRoles = string.Empty;

                //                        #region Commented Test

                //                        //if (oList.Title == "Documents")
                //                        //{
                //                        //    clientcontext.Load(oList.RootFolder);
                //                        //    clientcontext.ExecuteQuery();

                //                        //    clientcontext.Load(clientcontext.Web);
                //                        //    clientcontext.ExecuteQuery();

                //                        //    GetCounts(oList.RootFolder, clientcontext);
                //                        //}

                //                        #endregion

                //                        //if (oList.HasUniqueRoleAssignments)
                //                        //{
                //                        //    UniqueRoles = "Unique Permissions";
                //                        //}
                //                        //else
                //                        //{
                //                        //    UniqueRoles = "Inherit from Parent";
                //                        //}

                //                        if (oList.HasUniqueRoleAssignments)
                //                        {
                //                            ListFoldCount = 0;
                //                            ListFileCount = 0;

                //                            GetCountsatListLevel(oList.RootFolder, clientcontext);

                //                            string AdGroupsinSileCollListGroup = string.Empty;
                //                            string AdGroupsatSileCollListSite = string.Empty;

                //                            #region SiteColl Lists Permission check

                //                            RoleAssignmentCollection roles = oList.RoleAssignments;
                //                            clientcontext.Load(roles);
                //                            clientcontext.ExecuteQuery();

                //                            Web oWebx = clientcontext.Web;
                //                            clientcontext.Load(oWebx);
                //                            clientcontext.ExecuteQuery();

                //                            foreach (RoleAssignment rAssignment in roles)
                //                            {


                //                                #region Role Definations

                //                                RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                //                                clientcontext.Load(rdefColl);
                //                                clientcontext.ExecuteQuery();

                //                                string Design = string.Empty;
                //                                string Contribute = string.Empty;
                //                                string Read = string.Empty;
                //                                string FullControl = string.Empty;
                //                                string Edit = string.Empty;
                //                                string ViewOnly = string.Empty;
                //                                string Approve = string.Empty;
                //                                string ContributeLimited = string.Empty;
                //                                string OtherPermissions = string.Empty;

                //                                foreach (RoleDefinition rdef in rdefColl)
                //                                {
                //                                    clientcontext.Load(rdef);
                //                                    clientcontext.ExecuteQuery();

                //                                    switch (rdef.Name)
                //                                    {
                //                                        case "Design":
                //                                            Design = "Yes";
                //                                            break;

                //                                        case "Contribute":
                //                                            Contribute = "Yes";
                //                                            break;

                //                                        case "Read":
                //                                            Read = "Yes";
                //                                            break;

                //                                        case "Full Control":
                //                                            FullControl = "Yes";
                //                                            break;

                //                                        case "Edit":
                //                                            Edit = "Yes";
                //                                            break;

                //                                        case "View Only":
                //                                            ViewOnly = "Yes";
                //                                            break;

                //                                        case "Contribute Limited":
                //                                            ContributeLimited = "Yes";
                //                                            break;

                //                                        case "Approve":
                //                                            Approve = "Yes";
                //                                            break;

                //                                        default:
                //                                            OtherPermissions = rdef.Name;
                //                                            break;
                //                                    }
                //                                }

                //                                #endregion

                //                                try
                //                                {
                //                                    //if (!foundatListLevel)
                //                                    //{
                //                                    clientcontext.Load(rAssignment.Member);
                //                                    clientcontext.ExecuteQuery();

                //                                    if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                //                                    {
                //                                        continue;
                //                                    }

                //                                    if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                //                                    {
                //                                        Group ouserGroup = (Group)rAssignment.Member.TypedObject;
                //                                        clientcontext.Load(ouserGroup);
                //                                        clientcontext.ExecuteQuery();

                //                                        UserCollection userColl = ouserGroup.Users;
                //                                        clientcontext.Load(userColl);
                //                                        clientcontext.ExecuteQuery();

                //                                        foreach (User xUser in userColl)
                //                                        {
                //                                            if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                //                                            {

                //                                                //if (xUser.Title == "Everyone except external users")
                //                                                //{
                //                                                //clientcontext.Load(oList.RootFolder);
                //                                                //clientcontext.ExecuteQuery();   

                //                                                //AdGroupsinSileCollListGroup += ouserGroup.Title + ";";
                //                                                //foundatListLevel = true;
                //                                                //break;

                //                                                if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                //                                                {
                //                                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                                    {
                //                                                        if (xUser.Title.ToString().ToLower().Contains("everyone"))
                //                                                        {
                //                                                            //if (!BuiltinGroups.Contains(xUser.Title))
                //                                                            //{
                //                                                            //    BuiltinGroups.Add(xUser.Title);
                //                                                            //}

                //                                                            if (BuiltinGroups.ContainsKey(xUser.Title))
                //                                                            {
                //                                                                BuiltinGroups[xUser.Title]++;
                //                                                            }
                //                                                            else
                //                                                            {
                //                                                                BuiltinGroups.Add(xUser.Title, 1);
                //                                                            }
                //                                                        }
                //                                                        else
                //                                                        {
                //                                                            //if (!ADGroups.Contains(xUser.Title))
                //                                                            //{
                //                                                            //    ADGroups.Add(xUser.Title);
                //                                                            //}
                //                                                            if (ADGroups.ContainsKey(xUser.Title))
                //                                                            {
                //                                                                ADGroups[xUser.Title]++;
                //                                                            }
                //                                                            else
                //                                                            {
                //                                                                ADGroups.Add(xUser.Title, 1);
                //                                                            }
                //                                                        }

                //                                                        excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                                        excelWriterScoringNew.Flush();
                //                                                        //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "--" + "\"" + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + UniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");
                //                                                    }
                //                                                }

                //                                                //break;
                //                                            }
                //                                        }
                //                                    }
                //                                    if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                //                                    {
                //                                        //if (rAssignment.Member.Title == "Everyone except external users")
                //                                        //{
                //                                        //clientcontext.Load(oList.RootFolder);
                //                                        //clientcontext.ExecuteQuery();                                                  

                //                                        //AdGroupsatSileCollListSite += rAssignment.Member.Title + ";";
                //                                        //foundatListLevel = true;
                //                                        if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                //                                        {


                //                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                //                                            {
                //                                                if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                //                                                {
                //                                                    //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                //                                                    //{
                //                                                    //    BuiltinGroups.Add(rAssignment.Member.Title);
                //                                                    //}

                //                                                    if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                //                                                    {
                //                                                        BuiltinGroups[rAssignment.Member.Title]++;
                //                                                    }
                //                                                    else
                //                                                    {
                //                                                        BuiltinGroups.Add(rAssignment.Member.Title, 1);
                //                                                    }
                //                                                }
                //                                                else
                //                                                {
                //                                                    //if (!ADGroups.Contains(rAssignment.Member.Title))
                //                                                    //{
                //                                                    //    ADGroups.Add(rAssignment.Member.Title);
                //                                                    //}

                //                                                    if (ADGroups.ContainsKey(rAssignment.Member.Title))
                //                                                    {
                //                                                        ADGroups[rAssignment.Member.Title]++;
                //                                                    }
                //                                                    else
                //                                                    {
                //                                                        ADGroups.Add(rAssignment.Member.Title, 1);
                //                                                    }
                //                                                }

                //                                                excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                //                                                excelWriterScoringNew.Flush();
                //                                            }
                //                                        }
                //                                        //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "--" + "\"" + "\"" + "," + "\"" + UniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");


                //                                        //break;
                //                                        //}
                //                                    }
                //                                    //}
                //                                    //else
                //                                    //{
                //                                    //    break;
                //                                    //}
                //                                }
                //                                catch (Exception ex)
                //                                {
                //                                    continue;
                //                                }
                //                            }

                //                            //if (foundatListLevel)
                //                            //{
                //                            //    ListFoldCount = 0;
                //                            //    ListFileCount = 0;

                //                            //    GetCountsatListLevel(oList.RootFolder, clientcontext);

                //                            //    excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSileCollListSite + "\"" + "," + "\"" + AdGroupsinSileCollListGroup + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"");
                //                            //    excelWriterScoringNew.Flush();
                //                            //}

                //                            #endregion
                //                        }

                //                        clientcontext.Load(oList.RootFolder.Folders);
                //                        clientcontext.ExecuteQuery();

                //                        foreach (Folder sFolder in oList.RootFolder.Folders)
                //                        {
                //                            clientcontext.Load(sFolder);
                //                            clientcontext.ExecuteQuery();

                //                            if (sFolder.Name != "Forms")
                //                            {
                //                                GetCounts(sFolder, clientcontext, excelWriterScoringNew);
                //                            }
                //                        }
                //                    }
                //                }
                //            }

                //            #endregion

                //            #region SubSites

                //            WebCollection oWebs = clientcontext.Web.Webs;
                //            clientcontext.Load(oWebs);
                //            clientcontext.ExecuteQuery();

                //            foreach (Web oWeb in oWebs)
                //            {
                //                try
                //                {
                //                    clientcontext.Load(oWeb);
                //                    clientcontext.ExecuteQuery();
                //                    this.Text = oWeb.Url + "  Processing...";
                //                    getWeb(oWeb.Url, excelWriterScoringNew);
                //                }
                //                catch (Exception ex)
                //                {
                //                    continue;
                //                }
                //            }

                //            #endregion


                //            excelWriterScoringNew.Flush();
                //            excelWriterScoringNew.Close();

                //            string bGroups = string.Empty;
                //            string AdsGroups = string.Empty;

                //            foreach (KeyValuePair<string, int> kp in BuiltinGroups)
                //            {
                //                if (kp.Key != "FUN-SPO-SITECOLL-ADMINS" && (!kp.Key.ToLower().Contains("spo admin")))
                //                {
                //                    bGroups += kp.Key.ToString().Trim() + "(" + kp.Value.ToString() + ")" + "; ";
                //                }
                //            }

                //            foreach (KeyValuePair<string, int> kp in ADGroups)
                //            {
                //                if (kp.Key != "FUN-SPO-SITECOLL-ADMINS" && (!kp.Key.ToLower().Contains("spo admin")))
                //                {
                //                    AdsGroups += kp.Key.ToString().Trim() + "(" + kp.Value.ToString() + ")" + "; ";
                //                }
                //            }

                //            //foreach (string gp in BuiltinGroups)
                //            //{
                //            //    bGroups += gp + "; ";
                //            //}

                //            //foreach (string ap in ADGroups)
                //            //{
                //            //    AdsGroups += ap + "; ";
                //            //}

                //            excelWriterScoringMatrixNew.WriteLine("\"" + siteCollNameFileName + ".xlsx" + "\"" + "," + "\"" + clientcontext.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + bGroups + "\"" + "," + "\"" + AdsGroups + "\"" + "," + "\"" + startingTime + "\"" + "," + "\"" + DateTime.Now.ToString() + "\"" + "," + "\"" + "" + "\"");
                //            excelWriterScoringMatrixNew.Flush();

                //            //excelWriterScoringMatrixNew.WriteLine(siteCollNameFileName +".xlsx" + "," + clientcontext.Web.Url.ToString() + "," + admins + "," + bGroups + "," + AdsGroups + "," + DateTime.Now.ToString());
                //            //excelWriterScoringMatrixNew.Flush();
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        excelWriterScoringMatrixNew.WriteLine("\"" + "--" + "\"" + "," + "\"" + lstSiteColl[j] + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + "" + "\"" + "," + "\"" + startingTime + "\"" + "," + "\"" + DateTime.Now.ToString() + "\"" + "," + "\"" + ex.Message + "\"");
                //        excelWriterScoringMatrixNew.Flush();

                //        continue;
                //    }
                //}

                //excelWriterScoringMatrixNew.Flush();
                //excelWriterScoringMatrixNew.Close();

                //this.Text = "Process completed successfully.";
                //MessageBox.Show("Process Completed"); 
                #endregion
            }
            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        public void GetCounts(Folder Fld, ClientContext clientcontext, StreamWriter excelWriterScoringNew)
        {

            try
            {
                clientcontext.Load(Fld);
                clientcontext.Load(Fld, li => li.ListItemAllFields.HasUniqueRoleAssignments);
                clientcontext.ExecuteQuery();

                this.Text = "Folder : " + Fld.Name + " is Processing...";

                if (Fld.Name.Contains("drophere"))
                {
                    int j = 0;
                }

                if (Fld.ListItemAllFields.HasUniqueRoleAssignments)
                {
                    FoldCount = 0;
                    FileCount = 0;

                    clientcontext.Load(Fld.Files);
                    clientcontext.ExecuteQuery();

                    FileCount += Fld.Files.Count;

                    clientcontext.Load(Fld.Folders);
                    clientcontext.ExecuteQuery();

                    foreach (Folder folder in Fld.Folders)
                    {
                        clientcontext.Load(folder);
                        clientcontext.ExecuteQuery();

                        if (folder.Name != "Forms")
                        {
                            FoldCount++;
                        }
                    }

                    #region IF Folder has Unique permissions

                    bool folderhasSecurityGroup = false;

                    string AdGroupsinFolderGroup = string.Empty;
                    string AdGroupsatFolder = string.Empty;

                    RoleAssignmentCollection roles = Fld.ListItemAllFields.RoleAssignments;
                    clientcontext.Load(roles);
                    clientcontext.ExecuteQuery();

                    foreach (RoleAssignment rAssignment in roles)
                    {
                        clientcontext.Load(rAssignment.Member);
                        clientcontext.ExecuteQuery();

                        if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                        {
                            continue;
                        }

                        #region Role Definations

                        RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                        clientcontext.Load(rdefColl);
                        clientcontext.ExecuteQuery();

                        string Design = string.Empty;
                        string Contribute = string.Empty;
                        string Read = string.Empty;
                        string FullControl = string.Empty;
                        string Edit = string.Empty;
                        string ViewOnly = string.Empty;
                        string Approve = string.Empty;
                        string ContributeLimited = string.Empty;
                        string OtherPermissions = string.Empty;

                        foreach (RoleDefinition rdef in rdefColl)
                        {
                            clientcontext.Load(rdef);
                            clientcontext.ExecuteQuery();

                            switch (rdef.Name)
                            {
                                case "Design":
                                    Design = "Yes";
                                    break;

                                case "Contribute":
                                    Contribute = "Yes";
                                    break;

                                case "Read":
                                    Read = "Yes";
                                    break;

                                case "Full Control":
                                    FullControl = "Yes";
                                    break;

                                case "Edit":
                                    Edit = "Yes";
                                    break;

                                case "View Only":
                                    ViewOnly = "Yes";
                                    break;

                                case "Contribute Limited":
                                    ContributeLimited = "Yes";
                                    break;

                                case "Approve":
                                    Approve = "Yes";
                                    break;

                                default:
                                    OtherPermissions = rdef.Name;
                                    break;
                            }
                        }

                        #endregion

                        try
                        {

                            if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                            {
                                Group ouserGroup = (Group)rAssignment.Member.TypedObject;
                                clientcontext.Load(ouserGroup);
                                clientcontext.ExecuteQuery();

                                UserCollection userColl = ouserGroup.Users;
                                clientcontext.Load(userColl);
                                clientcontext.ExecuteQuery();

                                foreach (User xUser in userColl)
                                {
                                    if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                                    {
                                        //AdGroupsinFolderGroup += ouserGroup.Title + ";";
                                        //folderhasSecurityGroup = true;
                                        //break;
                                        if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                                        {
                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                            {
                                                if (xUser.Title.ToString().ToLower().Contains("everyone"))
                                                {
                                                    //if (!BuiltinGroups.Contains(xUser.Title))
                                                    //{
                                                    //    BuiltinGroups.Add(xUser.Title);
                                                    //}

                                                    if (BuiltinGroups.ContainsKey(xUser.Title))
                                                    {
                                                        BuiltinGroups[xUser.Title]++;
                                                    }
                                                    else
                                                    {
                                                        BuiltinGroups.Add(xUser.Title, 1);
                                                    }
                                                }
                                                else
                                                {
                                                    //if (!ADGroups.Contains(xUser.Title))
                                                    //{
                                                    //    ADGroups.Add(xUser.Title);
                                                    //}

                                                    if (ADGroups.ContainsKey(xUser.Title))
                                                    {
                                                        ADGroups[xUser.Title]++;
                                                    }
                                                    else
                                                    {
                                                        ADGroups.Add(xUser.Title, 1);
                                                    }
                                                }

                                                excelWriterScoringNew.WriteLine("\"" + "Folder" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + Fld.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                excelWriterScoringNew.Flush();
                                            }
                                        }
                                    }
                                }
                            }
                            if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                            {
                                //AdGroupsatFolder += rAssignment.Member.Title + ";";
                                //folderhasSecurityGroup = true;
                                if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                                {
                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                    {
                                        if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                                        {
                                            //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                                            //{
                                            //    BuiltinGroups.Add(rAssignment.Member.Title);
                                            //}

                                            if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                                            {
                                                BuiltinGroups[rAssignment.Member.Title]++;
                                            }
                                            else
                                            {
                                                BuiltinGroups.Add(rAssignment.Member.Title, 1);
                                            }
                                        }
                                        else
                                        {
                                            //if (!ADGroups.Contains(rAssignment.Member.Title))
                                            //{
                                            //    ADGroups.Add(rAssignment.Member.Title);
                                            //}

                                            if (ADGroups.ContainsKey(rAssignment.Member.Title))
                                            {
                                                ADGroups[rAssignment.Member.Title]++;
                                            }
                                            else
                                            {
                                                ADGroups.Add(rAssignment.Member.Title, 1);
                                            }
                                        }

                                        excelWriterScoringNew.WriteLine("\"" + "Folder" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + Fld.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                        excelWriterScoringNew.Flush();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }

                    //if (folderhasSecurityGroup)
                    //{
                    //    FoldCount = 0;
                    //    FileCount = 0;

                    //    clientcontext.Load(Fld.Files);
                    //    clientcontext.ExecuteQuery();

                    //    FileCount += Fld.Files.Count;

                    //    clientcontext.Load(Fld.Folders);
                    //    clientcontext.ExecuteQuery();

                    //    foreach (Folder folder in Fld.Folders)
                    //    {
                    //        clientcontext.Load(folder);
                    //        clientcontext.ExecuteQuery();

                    //        if (folder.Name != "Forms")
                    //        {
                    //            FoldCount++;
                    //        }
                    //    }

                    //    excelWriterScoringNew.WriteLine("\"" + "Folder" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + Fld.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatFolder + "\"" + "," + "\"" + AdGroupsinFolderGroup + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");
                    //    excelWriterScoringNew.Flush();
                    //}


                    #endregion
                }

                clientcontext.Load(Fld.Files);
                clientcontext.ExecuteQuery();

                foreach (Microsoft.SharePoint.Client.File sFile in Fld.Files)
                {
                    clientcontext.Load(sFile);
                    clientcontext.Load(sFile, li => li.ListItemAllFields.HasUniqueRoleAssignments);
                    clientcontext.ExecuteQuery();

                    if (sFile.ListItemAllFields.HasUniqueRoleAssignments)
                    {
                        bool filehasSecurityGroup = false;

                        string AdGroupsinFileGroup = string.Empty;
                        string AdGroupsatFile = string.Empty;

                        #region File Level Permission check

                        RoleAssignmentCollection roles = sFile.ListItemAllFields.RoleAssignments;
                        clientcontext.Load(roles);
                        clientcontext.ExecuteQuery();

                        foreach (RoleAssignment rAssignment in roles)
                        {

                            #region Role Definations

                            RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                            clientcontext.Load(rdefColl);
                            clientcontext.ExecuteQuery();

                            string Design = string.Empty;
                            string Contribute = string.Empty;
                            string Read = string.Empty;
                            string FullControl = string.Empty;
                            string Edit = string.Empty;
                            string ViewOnly = string.Empty;
                            string Approve = string.Empty;
                            string ContributeLimited = string.Empty;
                            string OtherPermissions = string.Empty;

                            foreach (RoleDefinition rdef in rdefColl)
                            {
                                clientcontext.Load(rdef);
                                clientcontext.ExecuteQuery();

                                switch (rdef.Name)
                                {
                                    case "Design":
                                        Design = "Yes";
                                        break;

                                    case "Contribute":
                                        Contribute = "Yes";
                                        break;

                                    case "Read":
                                        Read = "Yes";
                                        break;

                                    case "Full Control":
                                        FullControl = "Yes";
                                        break;

                                    case "Edit":
                                        Edit = "Yes";
                                        break;

                                    case "View Only":
                                        ViewOnly = "Yes";
                                        break;

                                    case "Contribute Limited":
                                        ContributeLimited = "Yes";
                                        break;

                                    case "Approve":
                                        Approve = "Yes";
                                        break;

                                    default:
                                        OtherPermissions = rdef.Name;
                                        break;
                                }
                            }

                            #endregion

                            try
                            {
                                clientcontext.Load(rAssignment.Member);
                                clientcontext.ExecuteQuery();

                                if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                                {
                                    continue;
                                }

                                if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                                {
                                    Group ouserGroup = (Group)rAssignment.Member.TypedObject;
                                    clientcontext.Load(ouserGroup);
                                    clientcontext.ExecuteQuery();

                                    UserCollection userColl = ouserGroup.Users;
                                    clientcontext.Load(userColl);
                                    clientcontext.ExecuteQuery();

                                    foreach (User xUser in userColl)
                                    {
                                        if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                                        {
                                            //AdGroupsinFileGroup += ouserGroup.Title + ";";
                                            //filehasSecurityGroup = true;
                                            if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                                            {
                                                if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                                {
                                                    if (xUser.Title.ToString().ToLower().Contains("everyone"))
                                                    {
                                                        //if (!BuiltinGroups.Contains(xUser.Title))
                                                        //{
                                                        //    BuiltinGroups.Add(xUser.Title);
                                                        //}

                                                        if (BuiltinGroups.ContainsKey(xUser.Title))
                                                        {
                                                            BuiltinGroups[xUser.Title]++;
                                                        }
                                                        else
                                                        {
                                                            BuiltinGroups.Add(xUser.Title, 1);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //if (!ADGroups.Contains(xUser.Title))
                                                        //{
                                                        //    ADGroups.Add(xUser.Title);
                                                        //}

                                                        if (ADGroups.ContainsKey(xUser.Title))
                                                        {
                                                            ADGroups[xUser.Title]++;
                                                        }
                                                        else
                                                        {
                                                            ADGroups.Add(xUser.Title, 1);
                                                        }
                                                    }

                                                    excelWriterScoringNew.WriteLine("\"" + "File" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + sFile.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                    excelWriterScoringNew.Flush();
                                                }
                                            }
                                        }
                                    }
                                }
                                if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                                {
                                    //AdGroupsatFile += rAssignment.Member.Title + ";";
                                    //filehasSecurityGroup = true;
                                    if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                                    {
                                        if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                        {
                                            if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                                            {
                                                //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                                                //{
                                                //    BuiltinGroups.Add(rAssignment.Member.Title);
                                                //}

                                                if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                                                {
                                                    BuiltinGroups[rAssignment.Member.Title]++;
                                                }
                                                else
                                                {
                                                    BuiltinGroups.Add(rAssignment.Member.Title, 1);
                                                }
                                            }
                                            else
                                            {
                                                //if (!ADGroups.Contains(rAssignment.Member.Title))
                                                //{
                                                //    ADGroups.Add(rAssignment.Member.Title);
                                                //}

                                                if (ADGroups.ContainsKey(rAssignment.Member.Title))
                                                {
                                                    ADGroups[rAssignment.Member.Title]++;
                                                }
                                                else
                                                {
                                                    ADGroups.Add(rAssignment.Member.Title, 1);
                                                }
                                            }

                                            excelWriterScoringNew.WriteLine("\"" + "File" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + sFile.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                            excelWriterScoringNew.Flush();
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }

                        //if (filehasSecurityGroup)
                        //{
                        //    excelWriterScoringNew.WriteLine("\"" + "File" + "\"" + "," + "\"" + clientcontext.Web.Url.Replace(clientcontext.Web.ServerRelativeUrl, "") + sFile.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatFile + "\"" + "," + "\"" + AdGroupsinFileGroup + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                        //    excelWriterScoringNew.Flush();
                        //}

                        #endregion

                    }
                }

                clientcontext.Load(Fld.Folders);
                clientcontext.ExecuteQuery();

                foreach (Folder folder in Fld.Folders)
                {
                    clientcontext.Load(folder);
                    clientcontext.ExecuteQuery();

                    if (folder.Name != "Forms")
                    {
                        GetCounts(folder, clientcontext, excelWriterScoringNew);
                    }
                }
            }
            catch (Exception ex)
            {

            }

            #region Extra Recurssive

            //foreach (Folder folder in Fld.Folders)
            //{
            //    clientcontext.Load(folder.Files);
            //    clientcontext.ExecuteQuery();

            //    FileCount += folder.Files.Count;

            //    clientcontext.Load(folder.Folders);
            //    clientcontext.ExecuteQuery();

            //    foreach (Folder subsubfolder in folder.Folders)
            //    {
            //        clientcontext.Load(subsubfolder);
            //        clientcontext.ExecuteQuery();

            //        if (subsubfolder.Name != "Forms")
            //        {
            //            FoldCount++;
            //        }
            //    }

            //    if (folder.Folders.Count > 0)
            //    {
            //        foreach (Folder subFolder in folder.Folders)
            //        {
            //            GetCounts(subFolder, clientcontext);
            //        }
            //    }
            //} 

            #endregion
        }
        public void GetCountsatListLevel(Folder Fld, ClientContext clientcontext)
        {
            clientcontext.Load(Fld);
            clientcontext.ExecuteQuery();

            clientcontext.Load(Fld.Files);
            clientcontext.ExecuteQuery();

            ListFileCount += Fld.Files.Count;

            clientcontext.Load(Fld.Folders);
            clientcontext.ExecuteQuery();

            foreach (Folder folder in Fld.Folders)
            {
                clientcontext.Load(folder);
                clientcontext.ExecuteQuery();

                if (folder.Name != "Forms")
                {
                    ListFoldCount++;
                }
            }

            clientcontext.Load(Fld.Folders);
            clientcontext.ExecuteQuery();

            foreach (Folder folder in Fld.Folders)
            {
                clientcontext.Load(folder);
                clientcontext.ExecuteQuery();

                if (folder.Name != "Forms")
                {
                    GetCountsatListLevel(folder, clientcontext);
                }
            }
        }
        public void getWeb(string siteURL, StreamWriter excelWriterScoringNew)
        {
            try
            {
                AuthenticationManager authManager = new AuthenticationManager();

                //using (var clientcontextSub = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteURL, "adam.a@VerinonTechnology.onmicrosoft.com", "Lot62215##"))
                using (var clientcontextSub = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteURL, "spo.admin.verinon@agilent.onmicrosoft.com", "Lot62215"))
                {
                    #region SubSite                 

                    try
                    {
                        RoleAssignmentCollection webRoleAssignments = null;
                        string webUniqueRoles = string.Empty;

                        Web oWeb = clientcontextSub.Web;
                        clientcontextSub.Load(oWeb);
                        clientcontextSub.Load(oWeb, li => li.HasUniqueRoleAssignments);
                        clientcontextSub.ExecuteQuery();

                        //if (oWeb.HasUniqueRoleAssignments)
                        //{
                        //    webUniqueRoles = "Unique Permissions";
                        //}
                        //else
                        //{
                        //    webUniqueRoles = "Inherit from Parent";
                        //}

                        if (oWeb.HasUniqueRoleAssignments)
                        {
                            bool foundatSiteLevel = false;

                            string AdGroupsinGroupWeb = string.Empty;
                            string AdGroupsatSiteWeb = string.Empty;

                            #region Subsite Permission check

                            webRoleAssignments = clientcontextSub.Web.RoleAssignments;
                            clientcontextSub.Load(webRoleAssignments);
                            clientcontextSub.ExecuteQuery();

                            foreach (RoleAssignment member1 in webRoleAssignments)
                            {

                                #region Role Definations

                                RoleDefinitionBindingCollection rdefColl = member1.RoleDefinitionBindings;
                                clientcontextSub.Load(rdefColl);
                                clientcontextSub.ExecuteQuery();

                                string Design = string.Empty;
                                string Contribute = string.Empty;
                                string Read = string.Empty;
                                string FullControl = string.Empty;
                                string Edit = string.Empty;
                                string ViewOnly = string.Empty;
                                string Approve = string.Empty;
                                string ContributeLimited = string.Empty;
                                string OtherPermissions = string.Empty;

                                foreach (RoleDefinition rdef in rdefColl)
                                {
                                    clientcontextSub.Load(rdef);
                                    clientcontextSub.ExecuteQuery();

                                    switch (rdef.Name)
                                    {
                                        case "Design":
                                            Design = "Yes";
                                            break;

                                        case "Contribute":
                                            Contribute = "Yes";
                                            break;

                                        case "Read":
                                            Read = "Yes";
                                            break;

                                        case "Full Control":
                                            FullControl = "Yes";
                                            break;

                                        case "Edit":
                                            Edit = "Yes";
                                            break;

                                        case "View Only":
                                            ViewOnly = "Yes";
                                            break;

                                        case "Contribute Limited":
                                            ContributeLimited = "Yes";
                                            break;

                                        case "Approve":
                                            Approve = "Yes";
                                            break;

                                        default:
                                            OtherPermissions = rdef.Name;
                                            break;
                                    }
                                }

                                #endregion

                                try
                                {
                                    //if (!foundatSiteLevel)
                                    //{
                                    clientcontextSub.Load(member1.Member);
                                    clientcontextSub.ExecuteQuery();

                                    if (member1.Member.Title.Contains("c:0u.c|tenant|"))
                                    {
                                        continue;
                                    }

                                    if (member1.Member.PrincipalType == PrincipalType.SharePointGroup)
                                    {
                                        Group ouserGroup = (Group)member1.Member.TypedObject;
                                        clientcontextSub.Load(ouserGroup);
                                        clientcontextSub.ExecuteQuery();

                                        UserCollection userColl = ouserGroup.Users;
                                        clientcontextSub.Load(userColl);
                                        clientcontextSub.ExecuteQuery();

                                        foreach (User xUser in userColl)
                                        {
                                            if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                                            {
                                                //if (xUser.Title == "Everyone except external users")
                                                //{
                                                //AdGroupsinGroupWeb += ouserGroup.Title + ";";
                                                //foundatSiteLevel = true;
                                                //break;
                                                if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                                                {
                                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                                    {
                                                        if (xUser.Title.ToString().ToLower().Contains("everyone"))
                                                        {
                                                            //if (!BuiltinGroups.Contains(xUser.Title))
                                                            //{
                                                            //    BuiltinGroups.Add(xUser.Title);
                                                            //}

                                                            if (BuiltinGroups.ContainsKey(xUser.Title))
                                                            {
                                                                BuiltinGroups[xUser.Title]++;
                                                            }
                                                            else
                                                            {
                                                                BuiltinGroups.Add(xUser.Title, 1);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //if (!ADGroups.Contains(xUser.Title))
                                                            //{
                                                            //    ADGroups.Add(xUser.Title);
                                                            //}

                                                            if (ADGroups.ContainsKey(xUser.Title))
                                                            {
                                                                ADGroups[xUser.Title]++;
                                                            }
                                                            else
                                                            {
                                                                ADGroups.Add(xUser.Title, 1);
                                                            }

                                                        }

                                                        excelWriterScoringNew.WriteLine("\"" + "SubSite" + "\"" + "," + "\"" + clientcontextSub.Web.Url.ToString() + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                        excelWriterScoringNew.Flush();
                                                    }
                                                    //excelWriterScoringNew.WriteLine("\"" + "SubSite" + "\"" + "," + "\"" + clientcontextSub.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");

                                                    //foundatSiteLevel = true;
                                                    //break;
                                                    //}
                                                }
                                            }
                                        }
                                    }
                                    if (member1.Member.PrincipalType == PrincipalType.SecurityGroup)
                                    {
                                        //if (member1.Member.Title == "Everyone except external users")
                                        //{
                                        //excelWriterScoringNew.WriteLine("\"" + "SubSite" + "\"" + "," + "\"" + clientcontextSub.Web.Url.ToString() + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + webUniqueRoles + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                                        //AdGroupsatSiteWeb += member1.Member.Title + ";";
                                        //foundatSiteLevel = true;
                                        if (lstADGroupsColl.Contains(member1.Member.Title.ToString().Trim().ToLower()))
                                        {
                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                            {
                                                if (member1.Member.Title.ToString().ToLower().Contains("everyone"))
                                                {
                                                    //if (!BuiltinGroups.Contains(member1.Member.Title))
                                                    //{
                                                    //    BuiltinGroups.Add(member1.Member.Title);
                                                    //}

                                                    if (BuiltinGroups.ContainsKey(member1.Member.Title))
                                                    {
                                                        BuiltinGroups[member1.Member.Title]++;
                                                    }
                                                    else
                                                    {
                                                        BuiltinGroups.Add(member1.Member.Title, 1);
                                                    }
                                                }
                                                else
                                                {
                                                    //if (!ADGroups.Contains(member1.Member.Title))
                                                    //{
                                                    //    ADGroups.Add(member1.Member.Title);
                                                    //}

                                                    if (ADGroups.ContainsKey(member1.Member.Title))
                                                    {
                                                        ADGroups[member1.Member.Title]++;
                                                    }
                                                    else
                                                    {
                                                        ADGroups.Add(member1.Member.Title, 1);
                                                    }
                                                }

                                                excelWriterScoringNew.WriteLine("\"" + "SubSite" + "\"" + "," + "\"" + clientcontextSub.Web.Url.ToString() + "\"" + "," + "\"" + member1.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                excelWriterScoringNew.Flush();
                                            }
                                        }
                                        //foundatSiteLevel = true;
                                        //break;
                                        //}
                                    }
                                    //}
                                    //else
                                    //{
                                    //    break;
                                    //}
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }
                            }

                            //if (foundatSiteLevel)
                            //{
                            //    excelWriterScoringNew.WriteLine("\"" + "Subsite" + "\"" + "," + "\"" + clientcontextSub.Url + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSiteWeb + "\"" + "," + "\"" + AdGroupsinGroupWeb + "\"" + "," + "\"" + "NA" + "\"" + "," + "\"" + "NA" + "\"");
                            //    excelWriterScoringNew.Flush();
                            //}

                            #endregion
                        }
                    }
                    catch (Exception ex)
                    {

                    }

                    #endregion

                    #region Lists on Subsites

                    ListCollection olistColl = clientcontextSub.Web.Lists;
                    clientcontextSub.Load(olistColl);
                    clientcontextSub.ExecuteQuery();

                    foreach (List oList in olistColl)
                    {
                        string listUniqueRoles = string.Empty;

                        clientcontextSub.Load(oList);
                        clientcontextSub.Load(oList, li => li.HasUniqueRoleAssignments);
                        clientcontextSub.ExecuteQuery();

                        string AdGroupsinSubsiteListGroup = string.Empty;
                        string AdGroupsatSubsiteListSite = string.Empty;

                        if (oList.BaseType == BaseType.DocumentLibrary)
                        {
                            if (oList.Title != "Form Templates" && oList.Title != "Site Assets" && oList.Title != "SitePages" && oList.Title != "Style Library" && oList.Hidden == false && oList.IsCatalog == false && oList.BaseTemplate == 101)
                            {
                                if (oList.HasUniqueRoleAssignments)
                                {
                                    ListFoldCount = 0;
                                    ListFileCount = 0;

                                    GetCountsatListLevel(oList.RootFolder, clientcontextSub);

                                    bool foundatListLevel = false;

                                    #region Subsite Lists Permission check

                                    RoleAssignmentCollection roles = oList.RoleAssignments;
                                    clientcontextSub.Load(roles);
                                    clientcontextSub.ExecuteQuery();

                                    Web oWebx = clientcontextSub.Web;
                                    clientcontextSub.Load(oWebx);
                                    clientcontextSub.ExecuteQuery();

                                    //Get all the RoleAssignments for this document
                                    foreach (RoleAssignment rAssignment in roles)
                                    {

                                        #region Role Definations

                                        RoleDefinitionBindingCollection rdefColl = rAssignment.RoleDefinitionBindings;
                                        clientcontextSub.Load(rdefColl);
                                        clientcontextSub.ExecuteQuery();

                                        string Design = string.Empty;
                                        string Contribute = string.Empty;
                                        string Read = string.Empty;
                                        string FullControl = string.Empty;
                                        string Edit = string.Empty;
                                        string ViewOnly = string.Empty;
                                        string Approve = string.Empty;
                                        string ContributeLimited = string.Empty;
                                        string OtherPermissions = string.Empty;

                                        foreach (RoleDefinition rdef in rdefColl)
                                        {
                                            clientcontextSub.Load(rdef);
                                            clientcontextSub.ExecuteQuery();

                                            switch (rdef.Name)
                                            {
                                                case "Design":
                                                    Design = "Yes";
                                                    break;

                                                case "Contribute":
                                                    Contribute = "Yes";
                                                    break;

                                                case "Read":
                                                    Read = "Yes";
                                                    break;

                                                case "Full Control":
                                                    FullControl = "Yes";
                                                    break;

                                                case "Edit":
                                                    Edit = "Yes";
                                                    break;

                                                case "View Only":
                                                    ViewOnly = "Yes";
                                                    break;

                                                case "Contribute Limited":
                                                    ContributeLimited = "Yes";
                                                    break;

                                                case "Approve":
                                                    Approve = "Yes";
                                                    break;

                                                default:
                                                    OtherPermissions = rdef.Name;
                                                    break;
                                            }
                                        }

                                        #endregion

                                        try
                                        {
                                            //if (!foundatListLevel)
                                            //{
                                            clientcontextSub.Load(rAssignment.Member);
                                            clientcontextSub.ExecuteQuery();

                                            if (rAssignment.Member.Title.Contains("c:0u.c|tenant|"))
                                            {
                                                continue;
                                            }

                                            if (rAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
                                            {
                                                Group ouserGroup = (Group)rAssignment.Member.TypedObject;
                                                clientcontextSub.Load(ouserGroup);
                                                clientcontextSub.ExecuteQuery();

                                                UserCollection userColl = ouserGroup.Users;
                                                clientcontextSub.Load(userColl);
                                                clientcontextSub.ExecuteQuery();

                                                foreach (User xUser in userColl)
                                                {
                                                    if (xUser.PrincipalType == PrincipalType.SecurityGroup)
                                                    {
                                                        //clientcontextSub.Load(oList.RootFolder);
                                                        //clientcontextSub.ExecuteQuery();

                                                        //AdGroupsinSubsiteListGroup += ouserGroup.Title + ";";
                                                        //foundatListLevel = true;
                                                        //break;

                                                        if (lstADGroupsColl.Contains(xUser.Title.ToString().Trim().ToLower()))
                                                        {
                                                            if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                                            {
                                                                if (xUser.Title.ToString().ToLower().Contains("everyone"))
                                                                {
                                                                    //if (!BuiltinGroups.Contains(xUser.Title))
                                                                    //{
                                                                    //    BuiltinGroups.Add(xUser.Title);
                                                                    //}

                                                                    if (BuiltinGroups.ContainsKey(xUser.Title))
                                                                    {
                                                                        BuiltinGroups[xUser.Title]++;
                                                                    }
                                                                    else
                                                                    {
                                                                        BuiltinGroups.Add(xUser.Title, 1);
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    //if (!ADGroups.Contains(xUser.Title))
                                                                    //{
                                                                    //    ADGroups.Add(xUser.Title);
                                                                    //}

                                                                    if (ADGroups.ContainsKey(xUser.Title))
                                                                    {
                                                                        ADGroups[xUser.Title]++;
                                                                    }
                                                                    else
                                                                    {
                                                                        ADGroups.Add(xUser.Title, 1);
                                                                    }

                                                                }

                                                                excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + xUser.Title + "\"" + "," + "\"" + ouserGroup.Title + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                                excelWriterScoringNew.Flush();
                                                            }
                                                        }
                                                        //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + listUniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");

                                                        //foundatListLevel = true;
                                                        //break;
                                                    }
                                                }
                                            }
                                            if (rAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
                                            {
                                                //if (rAssignment.Member.Title == "Everyone except external users")
                                                //{
                                                //clientcontextSub.Load(oList.RootFolder);
                                                //clientcontextSub.ExecuteQuery();

                                                //AdGroupsatSubsiteListSite += rAssignment.Member.Title + ";";
                                                //foundatListLevel = true;
                                                if (lstADGroupsColl.Contains(rAssignment.Member.Title.ToString().Trim().ToLower()))
                                                {
                                                    if (((!string.IsNullOrEmpty(Design)) || (!string.IsNullOrEmpty(Contribute)) || (!string.IsNullOrEmpty(Read)) || (!string.IsNullOrEmpty(FullControl)) || (!string.IsNullOrEmpty(Edit)) || (!string.IsNullOrEmpty(ViewOnly)) || (!string.IsNullOrEmpty(Approve)) || (!string.IsNullOrEmpty(ContributeLimited))) || ((!string.IsNullOrEmpty(OtherPermissions)) && (OtherPermissions != "Limited Access")))
                                                    {
                                                        if (rAssignment.Member.Title.ToString().ToLower().Contains("everyone"))
                                                        {
                                                            //if (!BuiltinGroups.Contains(rAssignment.Member.Title))
                                                            //{
                                                            //    BuiltinGroups.Add(rAssignment.Member.Title);
                                                            //}

                                                            if (BuiltinGroups.ContainsKey(rAssignment.Member.Title))
                                                            {
                                                                BuiltinGroups[rAssignment.Member.Title]++;
                                                            }
                                                            else
                                                            {
                                                                BuiltinGroups.Add(rAssignment.Member.Title, 1);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //if (!ADGroups.Contains(rAssignment.Member.Title))
                                                            //{
                                                            //    ADGroups.Add(rAssignment.Member.Title);
                                                            //}

                                                            if (ADGroups.ContainsKey(rAssignment.Member.Title))
                                                            {
                                                                ADGroups[rAssignment.Member.Title]++;
                                                            }
                                                            else
                                                            {
                                                                ADGroups.Add(rAssignment.Member.Title, 1);
                                                            }
                                                        }

                                                        excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + rAssignment.Member.Title + "\"" + "," + "\"" + "Directly Assigned" + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"" + "," + "\"" + Design + "\"" + "," + "\"" + Contribute + "\"" + "," + "\"" + Read + "\"" + "," + "\"" + FullControl + "\"" + "," + "\"" + Edit + "\"" + "," + "\"" + ViewOnly + "\"" + "," + "\"" + Approve + "\"" + "," + "\"" + ContributeLimited + "\"" + "," + "\"" + OtherPermissions + "\"");
                                                        excelWriterScoringNew.Flush();
                                                    }
                                                }

                                                //excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + "Yes" + "\"" + "," + "\"" + listUniqueRoles + "\"" + "," + "\"" + FoldCount.ToString() + "\"" + "," + "\"" + FileCount.ToString() + "\"");

                                                //foundatListLevel = true;
                                                //break;
                                                //}
                                            }
                                            //}
                                            //else
                                            //{
                                            //    break;
                                            //}
                                        }
                                        catch (Exception ex)
                                        {
                                            continue;
                                        }
                                    }

                                    //if (foundatListLevel)
                                    //{
                                    //    ListFoldCount = 0;
                                    //    ListFileCount = 0;

                                    //    GetCountsatListLevel(oList.RootFolder, clientcontextSub);

                                    //    excelWriterScoringNew.WriteLine("\"" + "Doc Library" + "\"" + "," + "\"" + oWebx.Url.Replace(oWebx.ServerRelativeUrl, "") + oList.RootFolder.ServerRelativeUrl + "\"" + "," + "\"" + admins + "\"" + "," + "\"" + AdGroupsatSubsiteListSite + "\"" + "," + "\"" + AdGroupsinSubsiteListGroup + "\"" + "," + "\"" + ListFoldCount.ToString() + "\"" + "," + "\"" + ListFileCount.ToString() + "\"");
                                    //    excelWriterScoringNew.Flush();
                                    //}

                                    #endregion
                                }

                                clientcontextSub.Load(oList.RootFolder.Folders);
                                clientcontextSub.ExecuteQuery();

                                foreach (Folder sFolder in oList.RootFolder.Folders)
                                {
                                    clientcontextSub.Load(sFolder);
                                    clientcontextSub.ExecuteQuery();

                                    if (sFolder.Name != "Forms")
                                    {
                                        GetCounts(sFolder, clientcontextSub, excelWriterScoringNew);
                                    }
                                }
                            }
                        }
                    }

                    #endregion

                    #region Recursive

                    WebCollection oWebs = clientcontextSub.Web.Webs;
                    clientcontextSub.Load(oWebs);
                    clientcontextSub.ExecuteQuery();

                    foreach (Web oWeb in oWebs)
                    {
                        clientcontextSub.Load(oWeb);
                        clientcontextSub.ExecuteQuery();

                        getWeb(oWeb.Url, excelWriterScoringNew);
                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {

            }
        }
        public static ClientContext createContext(string sitCollURL)
        {
            siteTitle = string.Empty;

            ClientContext contxt = new ClientContext(sitCollURL);
            SecureString secureStrPwd = new SecureString();

            //foreach (char x in "verinon@123".ToString())//need to change according to admin user credentials
            //{
            //    secureStrPwd.AppendChar(x);
            //}

            //SharePointOnlineCredentials credentials = new SharePointOnlineCredentials("spo.admin.verinon@agilent.onmicrosoft.com", secureStrPwd);

            foreach (char x in "Lot62215##".ToString())//need to change according to admin user credentials
            {
                secureStrPwd.AppendChar(x);
            }

            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials("adam.a@VerinonTechnology.onmicrosoft.com", secureStrPwd);

            contxt.Credentials = credentials;
            contxt.ExecuteQuery();
            contxt.RequestTimeout = -1;

            web = contxt.Web;
            contxt.Load(web);
            contxt.ExecuteQuery();
            try
            {
                siteTitle = web.Title;
            }
            catch
            {

            }
            return contxt;
        }
        public void WriteToErrorLog(string msg, string stkTrace, string title)
        {
            //log it
            FileStream fs1 = new FileStream("errorlog.txt", FileMode.Append, FileAccess.Write);
            StreamWriter s1 = new StreamWriter(fs1);
            //s1.WriteLine("Title: " + title);
            s1.WriteLine("Message: " + msg);
            s1.WriteLine("StackTrace: " + stkTrace);
            s1.WriteLine("Title: " + title);
            s1.WriteLine("Date/Time: " + System.DateTime.Now.ToString());
            s1.WriteLine("===========================================================================================");
            s1.Close();
            fs1.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = folderBrowserDialog1.SelectedPath;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = openFileDialog2.FileName;
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            #endregion

            //StreamWriter excelWriterScoringMatrixNew = null;

            //excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            //excelWriterScoringMatrixNew.WriteLine("Filename" + "," + "URL" + "," + "Owners" + "," + "Built-in-Groups" + "," + "AD Groups" + "," + "Start Time" + "," + "End Date" + "," + "Remarks");
            //excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            // Import_BlogAsHomeaspx(clientcontext);
                        }
                        catch (Exception es)
                        {

                        }
                    }
                }

                catch (Exception ex)
                {

                    //excelWriterScoringMatrixNew.WriteLine("Site Coll Owners" + "," + admins + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "" + "," + "");
                    //excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }


            //excelWriterScoringMatrixNew.Flush();
            //excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button6_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("ObjectType" + "," + "ObjectName" + "," + "ObjectURL" + "," + "SiteURL");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        //if (clientcontext.Web.Title.Contains("&") || clientcontext.Web.Title.Contains("ã€€") || clientcontext.Web.Title.Contains("ã€") || clientcontext.Web.Title.Contains("�") || clientcontext.Web.Title.Contains("Ã"))
                        //{
                        //    excelWriterScoringMatrixNew.WriteLine("SiteTitle" + "," + clientcontext.Web.Title.ToString() + "," + clientcontext.Web.Url + "," + clientcontext.Web.Url);
                        //    excelWriterScoringMatrixNew.Flush();
                        //}

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        //foreach (List oList in _Lists)
                        {
                            try
                            {
                                List oList = _Lists.GetByTitle("Events");
                                clientcontext.Load(oList);
                                clientcontext.ExecuteQuery();

                                Folder _RootFolder = oList.RootFolder;
                                clientcontext.Load(_RootFolder);
                                clientcontext.ExecuteQuery();

                                //if (oList.Title.Contains("&") || oList.Title.Contains("ã€€") || oList.Title.Contains("ã€") || oList.Title.Contains("�") || oList.Title.Contains("Ã"))
                                //{
                                //    excelWriterScoringMatrixNew.WriteLine("ListTitle" + "," + oList.Title + "," + _RootFolder.ServerRelativeUrl + "," + clientcontext.Web.Url);
                                //    excelWriterScoringMatrixNew.Flush();
                                //}

                                #region COMMENTED CODE

                                //if (oList.BaseTemplate == 101)
                                //{
                                //    FileCollection oFileColl = _RootFolder.Files;
                                //    clientcontext.Load(oFileColl);
                                //    clientcontext.ExecuteQuery();

                                //    foreach (Microsoft.SharePoint.Client.File oFile in oFileColl)
                                //    {
                                //        try
                                //        {
                                //            clientcontext.Load(oFile);
                                //            clientcontext.ExecuteQuery();

                                //            clientcontext.Load(oFile.ListItemAllFields);
                                //            clientcontext.ExecuteQuery();

                                //            if (oFile.Title.Contains("&")|| oFile.Title.Contains("ã€€")|| oFile.Title.Contains("ã€"))
                                //            {
                                //                excelWriterScoringMatrixNew.WriteLine("ItemTitle" + "," + oFile.Title.ToString() + "," + oFile.ServerRelativeUrl + "," + clientcontext.Web.Url);
                                //                excelWriterScoringMatrixNew.Flush();
                                //            }
                                //        }
                                //        catch (Exception ex)
                                //        {
                                //            continue;
                                //        }
                                //    }
                                //}
                                //else 

                                #endregion

                                #region ITEMS

                                {
                                    CamlQuery camlQuery = new CamlQuery();
                                    camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                                    ListItemCollection listItems = oList.GetItems(camlQuery);
                                    clientcontext.Load(listItems);
                                    clientcontext.ExecuteQuery();

                                    foreach (ListItem _Item in listItems)
                                    {
                                        try
                                        {
                                            clientcontext.Load(_Item);
                                            clientcontext.ExecuteQuery();

                                            DateTime Modified = Convert.ToDateTime(_Item["Created"]);
                                            FieldUserValue ModifiedBy = (FieldUserValue)_Item["Author"];


                                            //_Item["Title"] = oTitle;
                                            _Item["Modified"] = Modified;
                                            _Item["Editor"] = ModifiedBy;
                                            _Item.Update();
                                            clientcontext.ExecuteQuery();

                                            //if (oItem["Title"].ToString().Contains("&") || oItem["Title"].ToString().Contains("ã€€") || oItem["Title"].ToString().Contains("ã€"))
                                            //{
                                            //    excelWriterScoringMatrixNew.WriteLine("ItemTitle" + "," + oItem["Title"].ToString() + "," + clientcontext.Web.Url + "/Lists/" + oList.RootFolder.Name + "/DispForm.aspx?ID=" + oItem.Id + "" + "," + clientcontext.Web.Url);
                                            //    excelWriterScoringMatrixNew.Flush();
                                            //}
                                        }
                                        catch (Exception ex)
                                        {
                                            continue;
                                        }
                                    }
                                }

                                #endregion

                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button7_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("ObjectURL" + "," + "SiteURL" + "," + "OldTitle" + "," + "NewTitle");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    string[] URL = new string[] { "/1_Uploaded Files" };
                    string[] listURL = new string[] { "/Lists/" };
                    string[] ID = new string[] { ".aspx?ID=" };
                    string siteURL = string.Empty;
                    string ListName = string.Empty;

                    if (lstSiteColl[j].ToString().ToLower().Contains("/1_uploaded files/"))
                    {
                        ListName = "1_Uploaded Files";
                    }
                    else if (lstSiteColl[j].ToString().ToLower().Contains("/tasks/"))
                    {
                        ListName = "Tasks";
                    }
                    else if (lstSiteColl[j].ToString().ToLower().Contains("/announcements/"))
                    {
                        ListName = "Announcements";
                    }
                    else if (lstSiteColl[j].ToString().ToLower().Contains("/events/"))
                    {
                        ListName = "Events";
                    }
                    else if (lstSiteColl[j].ToString().ToLower().Contains("/posts/"))
                    {
                        ListName = "Posts";
                    }
                    else if (lstSiteColl[j].ToString().ToLower().Contains("/sitehistory/"))
                    {
                        ListName = "SiteHistory";
                    }
                    else if (lstSiteColl[j].ToString().ToLower().Contains("/ideas/"))
                    {
                        ListName = "Ideas";
                    }
                    else if (lstSiteColl[j].ToString().ToLower().Contains("/discussions/"))
                    {
                        ListName = "Discussions";
                    }

                    string itemID = lstSiteColl[j].ToString().Split(ID, StringSplitOptions.RemoveEmptyEntries)[1];

                    if (lstSiteColl[j].ToString().ToLower().Contains("/1_uploaded files"))
                    {
                        siteURL = lstSiteColl[j].ToString().Split(URL, StringSplitOptions.RemoveEmptyEntries)[0];
                    }
                    else if (lstSiteColl[j].ToString().ToLower().Contains("/lists/"))
                    {
                        siteURL = lstSiteColl[j].ToString().Split(listURL, StringSplitOptions.RemoveEmptyEntries)[0];
                    }

                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        List oList = oLists.GetByTitle(ListName);
                        clientcontext.Load(oList);
                        clientcontext.ExecuteQuery();

                        ListItem _Item = oList.GetItemById(itemID);
                        clientcontext.Load(_Item);
                        clientcontext.ExecuteQuery();

                        //clientcontext.Load(_Item,i=>i["Title"], i => i["Modified"], i => i["Editor"]);
                        //clientcontext.ExecuteQuery();                        

                        string OldTitle = _Item["Title"].ToString();

                        string oTitle = System.Web.HttpUtility.HtmlDecode(_Item["Title"].ToString());
                        DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                        FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                        try
                        {
                            _Item["Title"] = oTitle;
                            _Item["Modified"] = Modified;
                            _Item["Editor"] = ModifiedBy;
                            _Item.Update();
                            clientcontext.ExecuteQuery();

                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Url + "," + OldTitle + "," + oTitle);
                            excelWriterScoringMatrixNew.Flush();
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Url + "," + OldTitle + "," + oTitle);
                            excelWriterScoringMatrixNew.Flush();
                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "ERROR: " + ex.Message + "," + "" + "," + "");
                    excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button8_Click(object sender, EventArgs e)
        {

            //#region Site Collection URLS CSV Reading

            //List<string> lstSiteColl = new List<string>();
            //StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            //while (!sr.EndOfStream)
            //{
            //    try
            //    {
            //        lstSiteColl.Add(sr.ReadLine().Trim());
            //    }
            //    catch
            //    {
            //        continue;
            //    }
            //}

            //#endregion

            //StreamWriter excelWriterScoringMatrixNew = null;

            //excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            //excelWriterScoringMatrixNew.WriteLine("SourceSite" + "," + "TargetSite");
            //excelWriterScoringMatrixNew.Flush();

            //for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            //{
            //    this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

            //    try
            //    {       
            //        AuthenticationManager authManager = new AuthenticationManager();

            //        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
            //        {
            //            clientcontext.Load(clientcontext.Web);
            //            clientcontext.ExecuteQuery();

            //            Web _web = clientcontext.Web;
            //            clientcontext.Load(_web);
            //            clientcontext.ExecuteQuery();

            //            WebCollection _webs = _web.Webs;
            //            clientcontext.Load(_webs);
            //            clientcontext.ExecuteQuery();

            //            bool _webExist = _webs.Cast<Web>().Any(web => string.Equals(web.Url.ToLower(), (_web.Url + "/" + "MyContent").ToLower()));

            //            if (!_webExist)
            //            {
            //                WebCreationInformation wci = new WebCreationInformation();

            //                wci.Url = "MyContent"; // Relative URL
            //                wci.Title = "My Content";
            //                wci.Description = "My Personal Content";
            //                wci.UseSamePermissionsAsParentSite = true;
            //                wci.WebTemplate = "STS#0";

            //                // wci.WebTemplate = "STS#3";
            //                wci.Language = 1033; //LCID

            //                Web _TargetWeb = _webs.Add(wci);
            //                clientcontext.Load(_TargetWeb);
            //                clientcontext.ExecuteQuery();

            //                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + _TargetWeb.Url);
            //                excelWriterScoringMatrixNew.Flush();

            //            }

            //            try
            //            {
            //                var guid = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
            //                var featDefinition = _web.Features.Add(guid, true, FeatureDefinitionScope.Farm);
            //                clientcontext.Load(_web);
            //                clientcontext.ExecuteQuery();
            //            }
            //            catch (Exception es)
            //            {

            //            }

            //            //Categories, Community Feature; Discussions, Categories Lists
            //            try
            //            {
            //                //var guid = new Guid("961d6a9c-4388-4cf2-9733-38ee8c89afd4");
            //                //var featDefinition = _web.Features.Add(guid, true, FeatureDefinitionScope.Farm);
            //                //clientcontext.Load(_web);
            //                //clientcontext.ExecuteQuery();
            //            }
            //            catch (Exception es)
            //            {

            //            }
            //            //Blog Feature: Posts,Comments Lists
            //            try
            //            {
            //                var guid = new Guid("0d1c50f7-0309-431c-adfb-b777d5473a65");
            //                var featDefinition = _web.Features.Add(guid, true, FeatureDefinitionScope.Farm);
            //                clientcontext.Load(_web);
            //                clientcontext.ExecuteQuery();

            //            }
            //            catch (Exception es)
            //            {

            //            }
            //            //Mobile Browser View: MBrowserRedirect 	"d95c97f3-e528-4da2-ae9f-32b3535fbb59"
            //            try
            //            {
            //                var guid = new Guid("d95c97f3-e528-4da2-ae9f-32b3535fbb59");
            //                _web.Features.Remove(guid, true);
            //                clientcontext.ExecuteQuery();
            //            }
            //            catch (Exception es)
            //            {

            //            }


            //            backgroundWorker2.ReportProgress(0, "Lists are creatings for " + _placetype + " : " +
            //              _placeName + " " + SP_Public_Delcarations.currentSiteURL);

            //            Rename_Categories_List(clientcontext,
            //                _web, "Categories");

            //            FieldsCreation(false);

            //            Delete_Documents_List(clientcontext,
            //                _web, "Documents");

            //            Rename_SiteAssets_List(clientcontext,
            //              _web, "Site Assets");

            //            Rename_Pages_List(clientcontext,
            //               _web, "Pages");

            //            //hereg
            //            //if (_placetype != "Group")
            //            {
            //                create_SiteHistory(clientcontext, _web, _historyplaceID);
            //            }

            //            Create_Annoucement(clientcontext,
            //                _web, "Announcements");

            //            Create_Bookmarks_List(clientcontext,
            //                _web, "Bookmarks");

            //            Create_DiscussionBoard(clientcontext,
            //                _web, "Discussions");

            //            //Create_Pages(clientcontext,
            //            //   _web, "Pages");

            //            //Create_Videos(clientcontext,
            //            //   _web, "Videos");

            //            create_Events(clientcontext,
            //                _web, "Events");

            //            Create_Files_Documents_List(clientcontext,
            //                _web, "1_Uploaded Files", 101);


            //            //pages/cat/docs


            //            //Create_Files_Documents_List(clientcontext,
            //            //    _web, "Files", 108);

            //            //Create_Files_Documents_List(clientcontext,
            //            //    _web, "SBSDocuments", 108);

            //            Create_Idea_List(clientcontext,
            //                _web, "Ideas");

            //            Create_Messages_List(clientcontext,
            //                _web, "Messages");

            //            Create_Status_List(clientcontext,
            //                _web, "Status");


            //            Create_Tasks(clientcontext,
            //               _web, "Tasks");

            //            Create_Poll1(clientcontext,
            //                _web, "Polls");

            //            try
            //            {
            //                EnableRating(clientcontext, "Posts");
            //                List _postslist = _web.Lists.GetByTitle("Posts");
            //                clientcontext.Load(_web);
            //                clientcontext.ExecuteQuery();
            //                AddFieldstoDiscussionsList(clientcontext, _web, _postslist);
            //            }
            //            catch
            //            {

            //            }

            //            try
            //            {
            //                Import_BlogAsHomeaspx(clientcontext);
            //            }
            //            catch (Exception es)
            //            {

            //            }


            //            backgroundWorker2.ReportProgress(0, "Lists are created for " + _placetype + " : " +
            //              _placeName + " " + SP_Public_Delcarations.currentSiteURL);
            //            // ------------- end creating default lists -----------------//          


            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "ERROR: " + ex.Message);
            //        excelWriterScoringMatrixNew.Flush();

            //        continue;
            //    }
            //}

            //excelWriterScoringMatrixNew.Flush();
            //excelWriterScoringMatrixNew.Close();

            //this.Text = "Process completed successfully.";
            //MessageBox.Show("Process Completed");
        }
        private void button9_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SourceSite" + "," + "TargetSite");
            excelWriterScoringMatrixNew.Flush();

            string[] URL = new string[] { "#" };

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                string _SiteTitle = lstSiteColl[j].ToString().Split(URL, StringSplitOptions.RemoveEmptyEntries)[1];
                string _SiteName = lstSiteColl[j].ToString().Split(URL, StringSplitOptions.RemoveEmptyEntries)[0];
                string _SiteCollURL = "https://rsharepoint.sharepoint.com/sites/rspacedpc";

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_SiteCollURL, "svc -jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        Web _web = clientcontext.Web;
                        clientcontext.Load(_web);
                        clientcontext.ExecuteQuery();

                        WebCollection _webs = _web.Webs;
                        clientcontext.Load(_webs);
                        clientcontext.ExecuteQuery();

                        bool _webExist = _webs.Cast<Web>().Any(web => string.Equals(web.Url.ToLower(), (_web.Url + "/" + _SiteName).ToLower()));

                        if (!_webExist)
                        {
                            WebCreationInformation wci = new WebCreationInformation();

                            wci.Url = _SiteName; // Relative URL
                            wci.Title = _SiteTitle;
                            wci.Description = "My Personal Content";
                            wci.UseSamePermissionsAsParentSite = true;
                            wci.WebTemplate = "STS#0";

                            // wci.WebTemplate = "STS#3";
                            wci.Language = 1033; //LCID

                            Web _TargetWeb = _webs.Add(wci);
                            clientcontext.Load(_TargetWeb);
                            clientcontext.ExecuteQuery();

                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + _TargetWeb.Url);
                            excelWriterScoringMatrixNew.Flush();

                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "ERROR: " + ex.Message);
                    excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button10_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("Type" + "," + "ObjectURL" + "," + "SiteURL" + "," + "TitleColumn" + "," + "OtherColumn");
            excelWriterScoringMatrixNew.Flush();

            List<string> lstNameColl = new List<string>();

            lstNameColl.Add("1_Uploaded Files");
            // lstNameColl.Add("Events");
            //lstNameColl.Add("Announcements");
            //lstNameColl.Add("Tasks");
            //lstNameColl.Add("Posts");
            //lstNameColl.Add("Discussions");
            //lstNameColl.Add("SiteHistory");
            //lstNameColl.Add("Ideas");
            //lstNameColl.Add("Manage Categories");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                string SSSS = lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant("https://rsharepoint.sharepoint.com/sites/rworldgroups2/office-hours-engineers", "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        if (clientcontext.Web.Title.Contains("&") || clientcontext.Web.Title.Contains("ã€€") || clientcontext.Web.Title.Contains("ã€") || clientcontext.Web.Title.Contains("�") || clientcontext.Web.Title.Contains("Ã"))
                        {
                            excelWriterScoringMatrixNew.WriteLine("SiteTitle" + "," + "--" + "," + lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Title + "," + "--");
                            excelWriterScoringMatrixNew.Flush();
                        }

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        foreach (string lstName in lstNameColl)
                        {
                            try
                            {
                                List oList = oLists.GetByTitle(lstName);
                                clientcontext.Load(oList);
                                clientcontext.ExecuteQuery();

                                Folder _RootFolder = oList.RootFolder;
                                clientcontext.Load(_RootFolder);
                                clientcontext.ExecuteQuery();

                                if (oList.Title.Contains("&") || oList.Title.Contains("ã€€") || oList.Title.Contains("ã€") || oList.Title.Contains("�") || oList.Title.Contains("Ã"))
                                {
                                    excelWriterScoringMatrixNew.WriteLine("ListTitle" + "," + _RootFolder.ServerRelativeUrl + "," + lstSiteColl[j].ToString().Trim() + "," + oList.Title + "," + "--");
                                    excelWriterScoringMatrixNew.Flush();
                                }

                                switch (lstName)
                                {
                                    case "Events":

                                        #region EVENTS

                                        CamlQuery camlQuery = new CamlQuery();
                                        camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                                        ListItemCollection listItems = oList.GetItems(camlQuery);
                                        clientcontext.Load(listItems);
                                        clientcontext.ExecuteQuery();

                                        foreach (ListItem _Item in listItems)
                                        {
                                            try
                                            {
                                                clientcontext.Load(_Item);
                                                clientcontext.ExecuteQuery();

                                                //if (_Item.Id.ToString() == "9")
                                                {

                                                    string OldTitle = _Item["Title"].ToString();
                                                    string OldLocation = _Item["Location"].ToString();

                                                    //if (OldLocation.Contains("Ã") || OldTitle.Contains("Ã"))
                                                    // if (OldTitle.Contains("&") || OldTitle.Contains("�") || OldTitle.Contains("ã€€") || OldTitle.Contains("ã€") || OldLocation.Contains("&") || OldLocation.Contains("ã€€") || OldLocation.Contains("ã€") || OldLocation.Contains("�") || OldTitle.Contains("Ã") || OldLocation.Contains("Ã"))
                                                    if (OldLocation.Contains("&") || OldLocation.Contains("ã€€") || OldLocation.Contains("ã€") || OldLocation.Contains("�") || OldLocation.Contains("Ã"))
                                                    {
                                                        #region COMMENTED
                                                        //    //string oTitle = string.Empty;

                                                        //    //oTitle = System.Web.HttpUtility.HtmlDecode(_Item["Title"].ToString());

                                                        //    ////Ã¡,á:Ã³,ó

                                                        //    //if (oTitle.Contains("Ã¡"))
                                                        //    //{
                                                        //    //    oTitle = oTitle.Replace("Ã¡", "á");
                                                        //    //}
                                                        //    //if (oTitle.Contains("Ã³"))
                                                        //    //{
                                                        //    //    oTitle = oTitle.Replace("Ã³", "ó");
                                                        //    //}

                                                        //    //string oLocation = string.Empty;

                                                        //    //oLocation = System.Web.HttpUtility.HtmlDecode(_Item["Location"].ToString());

                                                        //    ////Ã¡,á:Ã³,ó

                                                        //    //if (oLocation.Contains("Ã¡"))
                                                        //    //{
                                                        //    //    oLocation = oLocation.Replace("Ã¡", "á");
                                                        //    //}
                                                        //    //if (oLocation.Contains("Ã³"))
                                                        //    //{
                                                        //    //    oLocation = oLocation.Replace("Ã³", "ó");
                                                        //    //} 
                                                        #endregion

                                                        string oLocation1 = System.Web.HttpUtility.HtmlDecode(_Item["Location"].ToString());

                                                        byte[] bytes = Encoding.Default.GetBytes(OldLocation);
                                                        string oLocation = Encoding.UTF8.GetString(bytes);

                                                        //byte[] bytes1 = Encoding.Default.GetBytes(OldTitle);
                                                        //string oTitle = Encoding.UTF8.GetString(bytes1);

                                                        DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                                        FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                                        //_Item["Title"] = oTitle;
                                                        _Item["Location"] = oLocation;
                                                        _Item["Modified"] = Modified;
                                                        _Item["Editor"] = ModifiedBy;
                                                        _Item.Update();
                                                        clientcontext.ExecuteQuery();

                                                        excelWriterScoringMatrixNew.WriteLine(lstName + "," + _Item.Id.ToString() + "," + lstSiteColl[j].ToString().Trim() + "," + OldTitle + "," + OldLocation);
                                                        excelWriterScoringMatrixNew.Flush();
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //excelWriterScoringMatrixNew.WriteLine("Events" + ", " + lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Url + "," + "" + "," + "");
                                                //excelWriterScoringMatrixNew.Flush();
                                                continue;
                                            }
                                        }
                                        #endregion

                                        break;

                                    case "1_Uploaded Files":

                                        #region 1_Uploaded Files

                                        CamlQuery camlQuery1 = new CamlQuery();
                                        camlQuery1.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                                        ListItemCollection listItems1 = oList.GetItems(camlQuery1);
                                        clientcontext.Load(listItems1);
                                        clientcontext.ExecuteQuery();

                                        //foreach (ListItem _Item in listItems1)
                                        {
                                            try
                                            {
                                                oList.EnableVersioning = false;
                                                oList.Update();
                                                clientcontext.ExecuteQuery();

                                                ListItem _Item = listItems1.GetById("49");
                                                clientcontext.Load(_Item);
                                                clientcontext.ExecuteQuery();

                                                //string OldName = _Item["FileLeafRef"].ToString();
                                                //string OldTitle = _Item["Title"].ToString();
                                                //�, í ,é, ó, é, ñ, ú
                                                //Ã¯Â¿Â½
                                                //if (OldTitle.Contains("Ã")|| OldName.Contains("Ã"))
                                                //if (OldTitle.Contains("&") || OldTitle.Contains("�") || OldTitle.Contains("ã€€") || OldTitle.Contains("ã€") || OldName.Contains("&") || OldName.Contains("ã€€") || OldName.Contains("ã€") || OldName.Contains("�") || OldTitle.Contains("Ã") || OldName.Contains("Ã"))
                                                //if (OldTitle.Contains("�"))


                                                //if (OldName.Contains("Ã"))
                                                {

                                                    #region OLD

                                                    string oTitle = string.Empty;
                                                    string oName = string.Empty;

                                                    //oTitle = OldTitle.Replace("��", "çã");
                                                    //oTitle = oTitle.Replace("�", "ç");//é,ç

                                                    //oName = OldName.Replace("��", "çã");
                                                    //oName = oName.Replace("�", "ç");//é,ç

                                                    ////string oTitle = string.Empty;

                                                    ////oTitle = System.Web.HttpUtility.HtmlDecode(_Item["FileLeafRef"].ToString());

                                                    //////Ã¡,á:Ã³,ó

                                                    ////if (oTitle.Contains("Ã¡"))
                                                    ////{
                                                    ////    oTitle = oTitle.Replace("Ã¡", "á");
                                                    ////}
                                                    ////if (oTitle.Contains("Ã³"))
                                                    ////{
                                                    ////    oTitle = oTitle.Replace("Ã³", "ó");
                                                    ////}
                                                    ////if (oTitle.Contains("Ã-"))
                                                    ////{
                                                    ////    oTitle = oTitle.Replace("Ã", "ó");
                                                    ////} 
                                                    #endregion

                                                    //byte[] bytes = Encoding.Default.GetBytes(OldName);
                                                    //string oName = Encoding.UTF8.GetString(bytes);


                                                    //byte[] bytes1 = Encoding.Default.GetBytes(OldTitle);
                                                    //string oTitle = Encoding.UTF8.GetString(bytes1);

                                                    //string oName = System.Web.HttpUtility.HtmlDecode(OldName);
                                                    //string oTitle = System.Web.HttpUtility.HtmlDecode(Title);
                                                    //string oName = "DECLARACIÓN-DE-POLÍTICA-DE-SEGURIDAD-DE-LA-INFORMACIÓN-ES.pdf";
                                                    //string oTitle = "Presentación de Recursos Humanos en Reunión General de Junio 2016";
                                                    //string oTitle = "PR-FI-127 Facturación Venta Directa de Equipos, Insumos y Repuestos (Spanish)";

                                                    User CreatedUser = default(User);
                                                    try
                                                    {
                                                        //Ensure user in Site
                                                        CreatedUser = clientcontext.Web.EnsureUser("wendy.jungbauer@ricoh-usa.com");

                                                        clientcontext.Load(CreatedUser);
                                                        clientcontext.ExecuteQuery();

                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        CreatedUser = clientcontext.Web.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                                        clientcontext.Load(CreatedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }
                                                    //Updating Ownership Information
                                                    FieldUserValue CreatedUserValue = new FieldUserValue();
                                                    CreatedUserValue.LookupId = CreatedUser.Id;

                                                    //string oTitle = _Item["Title"].ToString();

                                                    //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                                    FieldUserValue ModifiedBy = (FieldUserValue)_Item["Author"];
                                                    //DateTime Created = Convert.ToDateTime(_Item["Modified"]);

                                                    string xx = "February 26 2018 10:59 AM";
                                                    //string yy = "01-03-2018 16:20:31";
                                                    //DateTime Created = Convert.ToDateTime(xx);
                                                    DateTime Modified = getdateformat(xx);

                                                    string yy = "February 26 2018 10:37 AM";
                                                    DateTime Created = getdateformat(yy);

                                                    //_Item["FileLeafRef"] = oName;
                                                    //_Item.Update();
                                                    //clientcontext.ExecuteQuery();

                                                    //clientcontext.Load(_Item);
                                                    //clientcontext.ExecuteQuery();

                                                    //_Item["Title"] = oTitle;

                                                    _Item["Created"] = Created;
                                                    _Item["Author"] = CreatedUserValue;
                                                    _Item["Modified"] = Modified;
                                                    _Item["Editor"] = CreatedUserValue;
                                                    _Item.Update();
                                                    clientcontext.ExecuteQuery();

                                                    oList.EnableVersioning = true;
                                                    oList.Update();
                                                    clientcontext.ExecuteQuery();

                                                    // excelWriterScoringMatrixNew.WriteLine(lstName + "," + _Item.Id.ToString() + "," + lstSiteColl[j].ToString().Trim() + "," + OldTitle + "," + OldName);
                                                    // excelWriterScoringMatrixNew.Flush();
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //excelWriterScoringMatrixNew.WriteLine(lstName + "," + lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Url + "," + "" + "," + "");
                                                //excelWriterScoringMatrixNew.Flush();
                                            }
                                        }

                                        #endregion

                                        break;

                                    case "SiteHistory":

                                        #region SiteHistory

                                        CamlQuery camlQuery2 = new CamlQuery();
                                        camlQuery2.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                                        ListItemCollection listItems2 = oList.GetItems(camlQuery2);
                                        clientcontext.Load(listItems2);
                                        clientcontext.ExecuteQuery();

                                        foreach (ListItem _Item in listItems2)
                                        {
                                            try
                                            {
                                                clientcontext.Load(_Item);
                                                clientcontext.ExecuteQuery();

                                                string OldTitle = _Item["Title"].ToString();
                                                string OldFileDescription = _Item["PlaceDescription"].ToString();

                                                //if (OldFileDescription.Contains("Ã"))
                                                //if (OldTitle.Contains("Ã"))
                                                //if (OldTitle.Contains("&") || OldTitle.Contains("�") || OldTitle.Contains("ã€€") || OldTitle.Contains("ã€") || OldFileDescription.Contains("&") || OldFileDescription.Contains("ã€€") || OldFileDescription.Contains("ã€") || OldFileDescription.Contains("�") || OldTitle.Contains("Ã") || OldFileDescription.Contains("Ã"))
                                                {
                                                    byte[] bytes = Encoding.Default.GetBytes(OldFileDescription);
                                                    string oDescription = Encoding.UTF8.GetString(bytes);

                                                    //byte[] bytes1 = Encoding.Default.GetBytes(OldFileDescription);
                                                    //string oDescription = Encoding.UTF8.GetString(bytes1);

                                                    //string oDescription = System.Web.HttpUtility.HtmlDecode(_Item["PlaceDescription"].ToString());

                                                    #region COMMENTED

                                                    ////string oTitle = string.Empty;

                                                    ////oTitle = System.Web.HttpUtility.HtmlDecode(_Item["Title"].ToString());

                                                    //////Ã¡,á:Ã³,ó

                                                    ////if (oTitle.Contains("Ã¡"))
                                                    ////{
                                                    ////    oTitle = oTitle.Replace("Ã¡", "á");
                                                    ////}
                                                    ////if (oTitle.Contains("Ã³"))
                                                    ////{
                                                    ////    oTitle = oTitle.Replace("Ã³", "ó");
                                                    ////}

                                                    #endregion

                                                    DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                                    FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                                    //_Item["Title"] = oTitle;
                                                    _Item["PlaceDescription"] = oDescription;
                                                    _Item["Modified"] = Modified;
                                                    _Item["Editor"] = ModifiedBy;
                                                    _Item.Update();
                                                    clientcontext.ExecuteQuery();

                                                    excelWriterScoringMatrixNew.WriteLine(lstName + "," + _Item.Id.ToString() + "," + lstSiteColl[j].ToString().Trim() + "," + OldTitle + "," + OldFileDescription);
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //excelWriterScoringMatrixNew.WriteLine(lstName + "," + lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Url + "," + "" + "," + "");
                                                //excelWriterScoringMatrixNew.Flush();
                                            }
                                        }

                                        #endregion

                                        break;

                                    case "Manage Categories":
                                    case "Announcements":
                                    case "Tasks":
                                    case "Posts":
                                    case "Discussions":
                                    case "Ideas":

                                        #region OTHER LISTS

                                        CamlQuery camlQuery3 = new CamlQuery();
                                        camlQuery3.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                                        ListItemCollection listItems3 = oList.GetItems(camlQuery3);
                                        clientcontext.Load(listItems3);
                                        clientcontext.ExecuteQuery();

                                        //foreach (ListItem _Item in listItems3)
                                        {
                                            try
                                            {
                                                ListItem _Item = oList.GetItemById("3");
                                                clientcontext.Load(_Item);
                                                clientcontext.ExecuteQuery();

                                                clientcontext.Load(_Item);
                                                clientcontext.ExecuteQuery();

                                                string OldTitle = _Item["Title"].ToString();
                                                //string OldFileDescription = _Item["PlaceDescription"].ToString();

                                                //if (OldFileDescription.Contains("Ã"))
                                                //if (OldTitle.Contains("Ã"))

                                                //if (OldTitle.Contains("&") || OldTitle.Contains("�") || OldTitle.Contains("ã€€") || OldTitle.Contains("ã€") || OldTitle.Contains("Ã"))
                                                if (OldTitle.Contains("�"))
                                                {
                                                    //byte[] bytes = Encoding.Default.GetBytes(OldTitle);
                                                    //string oTitle = Encoding.UTF8.GetString(bytes);

                                                    //byte[] bytes1 = Encoding.Default.GetBytes(OldFileDescription);
                                                    //string oDescription = Encoding.UTF8.GetString(bytes1);

                                                    string oTitle = string.Empty;
                                                    oTitle = System.Web.HttpUtility.HtmlDecode(_Item["Title"].ToString());

                                                    #region COMMENTED                                                    

                                                    //////Ã¡,á:Ã³,ó

                                                    ////if (oTitle.Contains("Ã¡"))
                                                    ////{
                                                    ////    oTitle = oTitle.Replace("Ã¡", "á");
                                                    ////}
                                                    ////if (oTitle.Contains("Ã³"))
                                                    ////{
                                                    ////    oTitle = oTitle.Replace("Ã³", "ó");
                                                    ////}

                                                    #endregion

                                                    DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                                    FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                                    _Item["Title"] = oTitle;
                                                    ////_Item["PlaceDescription"] = oDescription;
                                                    _Item["Modified"] = Modified;
                                                    _Item["Editor"] = ModifiedBy;
                                                    _Item.Update();
                                                    clientcontext.ExecuteQuery();

                                                    excelWriterScoringMatrixNew.WriteLine(lstName + "," + _Item.Id.ToString() + "," + lstSiteColl[j].ToString().Trim() + "," + OldTitle + "," + "--");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //excelWriterScoringMatrixNew.WriteLine(lstName + "," + lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Url + "," + "" + "," + "");
                                                //excelWriterScoringMatrixNew.Flush();
                                            }
                                        }

                                        #endregion

                                        break;
                                }
                            }
                            catch (Exception ex)
                            {
                                excelWriterScoringMatrixNew.WriteLine("ERROR1 : " + ex.Message + "," + lstSiteColl[j].ToString().Trim() + "," + "" + "," + "" + "," + "");
                                excelWriterScoringMatrixNew.Flush();
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine("ERROR2 : " + ex.Message + "," + lstSiteColl[j].ToString().Trim() + "," + "" + "," + "" + "," + "");
                    excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }
            //
            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button11_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {

                        Web _Web = clientcontext.Web;
                        clientcontext.Load(_Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = _Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        List _list = null;

                        try
                        {
                            _list = _Lists.GetByTitle("Checkpoints");
                            clientcontext.Load(_list);
                            clientcontext.ExecuteQuery();
                        }
                        catch
                        {
                            ListCreationInformation info = new ListCreationInformation();
                            info.Description = "Checkpoints";
                            info.Title = "Checkpoints";
                            info.TemplateType = 171;
                            _list = _Lists.Add(info);
                            clientcontext.Load(_list);
                            clientcontext.ExecuteQuery();
                        }

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
        }
        private void button12_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[2] { new DataColumn("DID", typeof(string)), new DataColumn("URL", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);
            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            DataTable dtJiveObjects = new DataTable();
            dtJiveObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("DID", typeof(string)), new DataColumn("Modified", typeof(string)), new DataColumn("ModifiedBy", typeof(string)), new DataColumn("TagsSet", typeof(string)) });

            string csvData1 = System.IO.File.ReadAllText(textBox3.Text);
            foreach (string row in csvData1.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtJiveObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocumentsModfiedTagsReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("DID" + "," + "URL" + "," + "ModifyStatus" + "TagsStatus");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;
            string[] SiteSplit = new string[] { "/Pages/" };
            string[] FileSplit = new string[] { "/Documents/" };
            string[] TagsSplit = new string[] { "|" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string Modified = string.Empty;
                    var ModifiedBy = string.Empty;
                    string TagsSet = string.Empty;
                    string[] TagsColl = null;
                    string tagStatus = "NBD";

                    string _SiteTitle = drImported["URL"].ToString().Split(SiteSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    string _FilePath = drImported["URL"].ToString().Split(FileSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();

                    this.Text = (count).ToString() + " : " + _SiteTitle;
                    count++;

                    bool itemFound = false;

                    foreach (DataRow drJive in dtJiveObjects.Rows)
                    {
                        if (drImported["DID"].ToString().Trim() == drJive["DID"].ToString().Trim())
                        {
                            Modified = drJive["Modified"].ToString().Trim();
                            ModifiedBy = drJive["ModifiedBy"].ToString().Trim();
                            TagsSet = drJive["TagsSet"].ToString().Trim();
                            TagsColl = TagsSet.Split(TagsSplit, StringSplitOptions.RemoveEmptyEntries);
                            itemFound = true;
                            break;
                        }
                    }

                    if (itemFound)
                    {
                        AuthenticationManager authManager = new AuthenticationManager();
                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_SiteTitle, "svc-jivemigration1@rsharepoint.onmicrosoft.com", "Vak52950"))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();

                            List _List = null;
                            try
                            {
                                _List = oWeb.Lists.GetByTitle("2_Documents and Pages"); ;
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                _List.EnableVersioning = false;
                                _List.Update();
                                clientcontext.ExecuteQuery();

                                _List.ForceCheckout = false;
                                _List.Update();
                                clientcontext.ExecuteQuery();

                                try
                                {
                                    clientcontext.Load(_List.RootFolder);
                                    clientcontext.ExecuteQuery();

                                    Folder docFolder = null;

                                    try
                                    {
                                        docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                        clientcontext.Load(docFolder);
                                        clientcontext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    { }

                                    if (docFolder != null)
                                    {
                                        ListItem _Item = docFolder.Files.GetByUrl(_FilePath).ListItemAllFields;
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();

                                        try
                                        {
                                            TaxonomyFieldValueCollection taxFieldValues = _Item["Tags"] as TaxonomyFieldValueCollection;

                                            if (taxFieldValues.Count < 1)
                                            {
                                                tagStatus = "NeedToImport";

                                                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(oWeb.Context);
                                                clientcontext.Load(taxonomySession.TermStores);
                                                clientcontext.ExecuteQuery();

                                                TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");
                                                clientcontext.Load(termStore);
                                                clientcontext.ExecuteQuery();

                                                clientcontext.Load(termStore.Groups);
                                                clientcontext.ExecuteQuery();

                                                TermGroup group = termStore.Groups.GetByName("RicohTags");
                                                clientcontext.Load(group);
                                                clientcontext.ExecuteQuery();

                                                clientcontext.Load(group.TermSets);
                                                clientcontext.ExecuteQuery();

                                                TermSet termSet = group.TermSets.GetByName("TagsTermSet");
                                                clientcontext.Load(termSet);
                                                clientcontext.ExecuteQuery();

                                                Field _taxnomyField = _List.Fields.GetByTitle("Tags");
                                                clientcontext.Load(_taxnomyField);
                                                clientcontext.ExecuteQuery();

                                                TaxonomyField txField = clientcontext.CastTo<TaxonomyField>(_taxnomyField);
                                                clientcontext.Load(txField);
                                                clientcontext.ExecuteQuery();

                                                TaxonomyFieldValueCollection termValues = null;

                                                string termValueString = string.Empty;
                                                string termId = string.Empty;

                                                try
                                                {
                                                    foreach (string tv in TagsColl)
                                                    {
                                                        string mtermId = string.Empty;

                                                        try
                                                        {
                                                            if (string.IsNullOrEmpty(mtermId))
                                                            {
                                                                mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);
                                                            }

                                                            if (!string.IsNullOrEmpty(mtermId))
                                                                termValueString += "1033" + ";#" + tv + "|" + mtermId + ";#";

                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            continue;
                                                        }
                                                    }

                                                    if (!string.IsNullOrEmpty(termValueString))
                                                    {
                                                        termValueString = termValueString.Remove(termValueString.Length - 2);
                                                        termValues = new TaxonomyFieldValueCollection(clientcontext, termValueString, txField);
                                                        txField.SetFieldValueByValueCollection(_Item, termValues);

                                                        _Item.Update();
                                                        clientcontext.Load(_Item);
                                                        clientcontext.ExecuteQuery();

                                                        //_Item["Modified"] = Modified;
                                                        //_Item["Editor"] = ModifiedBy;
                                                        //_Item.Update();
                                                        //clientcontext.Load(_Item);
                                                        //clientcontext.ExecuteQuery();

                                                        tagStatus = "Success";
                                                    }
                                                    else
                                                    {
                                                        tagStatus = "NoTermValueString";
                                                    }
                                                }
                                                catch (Exception EX)
                                                {
                                                    tagStatus = "Failure : " + EX.Message;
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {

                                        }

                                        User CreatedUser = default(User);
                                        try
                                        {
                                            CreatedUser = oWeb.EnsureUser(ModifiedBy);
                                            clientcontext.Load(CreatedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            CreatedUser = oWeb.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                            clientcontext.Load(CreatedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        FieldUserValue CreatedUserValue = new FieldUserValue();
                                        CreatedUserValue.LookupId = CreatedUser.Id;

                                        DateTime dtModified = getdateformat(Modified);

                                        _Item["Modified"] = dtModified;
                                        _Item["Editor"] = CreatedUserValue;
                                        _Item.Update();
                                        clientcontext.ExecuteQuery();

                                        excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Success" + "," + tagStatus);
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Failure : " + ex.Message + "," + tagStatus);
                                    excelWriterScoringMatrixNew.Flush();
                                }

                                _List.EnableVersioning = true;
                                _List.Update();
                                clientcontext.ExecuteQuery();

                                _List.ForceCheckout = true;
                                _List.Update();
                                clientcontext.ExecuteQuery();
                            }
                        }
                    }
                    else
                    {
                        excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "ItemIDNotFound" + "," + "NA");
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Failure due to : " + ex.Message + "," + "NA");
                    excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }
            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        protected DateTime getdateformat(string date)
        {
            // DateTime dt1 = DateTime.Now;
            DateTime? dt1 = null;

            try
            {
                string[] cultureNames = { "en-US", "ru-RU", "ja-JP" };

                foreach (string cultureName in cultureNames)
                {
                    CultureInfo culture = new CultureInfo(cultureName);
                    try
                    {
                        dt1 = Convert.ToDateTime(date, culture);
                        break;
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            catch
            {

            }
            return dt1.Value;
        }

        protected DateTime Checkdateformat(DateTime date)
        {
            DateTime OriginalDate = date;
            try
            {
                int imortedMonth = date.Month;
                int imortedDay = date.Day;
                int imortedYear = date.Year;

                if (imortedDay <= 12)
                {
                    string actualFormat = date.Day.ToString() + "/" + date.Month.ToString() + "/" + date.Year.ToString() + " " + date.TimeOfDay.ToString();
                    OriginalDate = getdateformat(actualFormat);
                }
            }
            catch
            {

            }
            return OriginalDate;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "FileCategories" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ItemID" + "," + "Categories");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        List _List = null;

                        try
                        {
                            _List = clientcontext.Web.Lists.GetByTitle("1_Uploaded Files");
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                            ListItemCollection listItems = _List.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem oItem in listItems)
                            {
                                try
                                {
                                    string Categories = string.Empty;

                                    var lookupValues = new ArrayList();
                                    FieldLookupValue[] values = oItem["Categorization"] as FieldLookupValue[];

                                    foreach (FieldLookupValue value in values)
                                    {
                                        string value1 = value.LookupValue.ToString().Replace(",", "$");
                                        Categories += value1 + "|";
                                    }

                                    if (!string.IsNullOrEmpty(Categories))
                                    {
                                        excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + oItem.Id.ToString() + "," + Categories);
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }
                            }
                        }
                        catch (Exception ex)
                        { }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            #endregion

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button14_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DCTRemove" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ItemID" + "," + "Categories");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var context = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        context.Load(context.Web);
                        context.ExecuteQuery();
                        web = context.Web;


                        try
                        {

                            context.Load(web.Lists);
                            context.ExecuteQuery();


                            List list = web.Lists.GetByTitle("1_Uploaded Files");

                            context.Load(list);
                            context.Load(list.RootFolder);
                            context.Load(list.RootFolder.Files);

                            context.ExecuteQuery();


                            list.EnableVersioning = false;
                            list.Update();
                            context.ExecuteQuery();

                            //FileCollection filecoll = list.RootFolder.Files;
                            //context.Load(filecoll);
                            //context.ExecuteQuery();

                            var items = list.GetItems(CreateAllFilesQuery());
                            context.Load(items, icol => icol.Include(i => i.File));
                            context.ExecuteQuery();
                            var filecoll = items.Select(i => i.File).ToList();



                            ContentTypeCollection contentTypeColl = list.ContentTypes;
                            // ContentTypeCollection contentTypeColl = mweb.ContentTypes;

                            context.Load(contentTypeColl);
                            context.ExecuteQuery();
                            ContentType defaultcontentType = null;
                            ContentType ricohcontentType = null;
                            foreach (ContentType eachcontenttype in contentTypeColl)
                            {
                                context.Load(eachcontenttype);
                                context.ExecuteQuery();

                                if (eachcontenttype.Name == "RicohContentType")
                                {
                                    ricohcontentType = eachcontenttype;
                                    context.Load(eachcontenttype);
                                    context.ExecuteQuery();
                                    context.Load(ricohcontentType);
                                    ricohcontentType.ReadOnly = false;
                                    ricohcontentType.Update(false);
                                    context.Load(ricohcontentType);
                                    context.ExecuteQuery();

                                }
                                else if (eachcontenttype.Name == "Document")
                                {
                                    defaultcontentType = eachcontenttype;
                                    context.Load(eachcontenttype);
                                    context.ExecuteQuery();
                                    context.Load(defaultcontentType);
                                    context.ExecuteQuery();

                                }


                            }

                            foreach (Microsoft.SharePoint.Client.File f in filecoll)
                            {
                                try
                                {
                                    context.Load(f);
                                    context.Load(f.ListItemAllFields);
                                    context.ExecuteQuery();
                                    ListItem item = f.ListItemAllFields;
                                    context.Load(item);
                                    context.ExecuteQuery();

                                    if (item["ContentTypeId"].ToString() == defaultcontentType.Id.ToString())
                                    {
                                        DateTime Modified = Convert.ToDateTime(item["Modified"]);
                                        FieldUserValue ModifiedBy = (FieldUserValue)item["Editor"];


                                        item["ContentTypeId"] = ricohcontentType.Id;
                                        item.Update();
                                        context.ExecuteQuery();


                                        //if (item.File.CheckOutType.ToString() == "None")
                                        //{
                                        //    item.File.CheckOut();
                                        //    item["ContentTypeId"] = newCTID;
                                        //    item.Update();
                                        //    item.File.CheckIn("Checked in");
                                        //}


                                        item["Modified"] = Modified;
                                        item["Editor"] = ModifiedBy;

                                        item.Update();
                                        context.Load(item);
                                        context.ExecuteQuery();

                                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + item.Id.ToString());
                                        excelWriterScoringMatrixNew.Flush();


                                    }
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }

                            }


                            defaultcontentType.DeleteObject();
                            context.ExecuteQuery();


                            list.EnableVersioning = true;
                            list.Update();
                            context.ExecuteQuery();

                            #region Old Code

                            //_List = clientcontext.Web.Lists.GetByTitle("1_Uploaded Files");
                            //clientcontext.Load(_List);
                            //clientcontext.ExecuteQuery();

                            //CamlQuery camlQuery = new CamlQuery();
                            //camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                            //ListItemCollection listItems = _List.GetItems(camlQuery);
                            //clientcontext.Load(listItems);
                            //clientcontext.ExecuteQuery();

                            //foreach (ListItem oItem in listItems)
                            //{
                            //    try
                            //    {
                            //        string Categories = string.Empty;

                            //        var lookupValues = new ArrayList();
                            //        FieldLookupValue[] values = oItem["Categorization"] as FieldLookupValue[];

                            //        foreach (FieldLookupValue value in values)
                            //        {
                            //            Categories += value.LookupValue + "|";
                            //        }

                            //        if (string.IsNullOrEmpty(Categories))
                            //        {
                            //            excelWriterScoringMatrixNew.WriteLine(clientcontext.Web.Url + "," + oItem.Id.ToString() + "," + Categories);
                            //            excelWriterScoringMatrixNew.Flush();
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        continue;
                            //    }
                            //} 

                            #endregion
                        }
                        catch (Exception ex)
                        { }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            #endregion

            // excelWriterScoringMatrixNew.Flush();
            // excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        public static CamlQuery CreateAllFilesQuery()
        {
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";
            return qry;
        }
        private void button15_Click(object sender, EventArgs e)
        {
            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            //dtImportedObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("SiteURL", typeof(string)), new DataColumn("ItemID", typeof(string)), new DataColumn("Categories", typeof(string)) });
            dtImportedObjects.Columns.AddRange(new DataColumn[1] { new DataColumn("SiteURL", typeof(string)) });
            //Read the contents of CSV file.  
            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            //Execute a loop over the rows.  
            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;
                    //Execute a loop over the columns.  
                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "CategoryColumnReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            //excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ItemURL" + "," + "Status");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "Status");

            excelWriterScoringMatrixNew.Flush();

            int count = 0;

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string _SiteTitle = drImported["SiteURL"].ToString().Trim();
                    //string _ItemID = drImported["ItemID"].ToString().Trim();

                    this.Text = (count).ToString() + " : " + _SiteTitle;

                    count++;

                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_SiteTitle,
                        "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web _web = _cContext.Web;
                        _cContext.Load(_web);
                        _cContext.ExecuteQuery();

                        List Pagelist = _cContext.Web.Lists.GetByTitle("1_Uploaded Files");
                        _cContext.Load(Pagelist);
                        _cContext.Load(Pagelist.RootFolder);
                        _cContext.ExecuteQuery();

                        #region Report on "Categorization" column

                        //bool CategorizationIsCorrect = false;

                        //ContentTypeCollection ctColl = Pagelist.ContentTypes;
                        //_cContext.Load(ctColl);
                        //_cContext.ExecuteQuery();

                        //foreach (ContentType ct in ctColl)
                        //{
                        //    _cContext.Load(ct);
                        //    _cContext.ExecuteQuery();

                        //    if (ct.Name == "RicohContentType")
                        //    {
                        //        FieldCollection fdLnkColl = ct.Fields;
                        //        _cContext.Load(fdLnkColl);
                        //        _cContext.ExecuteQuery();

                        //        foreach (Field fd in fdLnkColl)
                        //        {
                        //            if (fd.Title == "Categorization")
                        //            {
                        //                CategorizationIsCorrect = true;
                        //                break;
                        //            }
                        //        }
                        //    }

                        //    if (CategorizationIsCorrect)
                        //        break;
                        //}

                        //if (!CategorizationIsCorrect)
                        //{
                        //    excelWriterScoringMatrixNew.WriteLine(_web.Url.ToString() + "," + "Not");
                        //    excelWriterScoringMatrixNew.Flush();
                        //}

                        #endregion

                        #region Delete and Add "Categorization" column

                        FieldCollection fields = Pagelist.Fields;
                        _cContext.Load(fields);
                        _cContext.ExecuteQuery();

                        try
                        {
                            Field field = Pagelist.Fields.GetByTitle("Categorization");
                            _cContext.Load(field);
                            _cContext.ExecuteQuery();

                            field.DeleteObject();
                            _cContext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + "Categorization Delete: " + ex.Message);
                            excelWriterScoringMatrixNew.Flush();
                        }

                        _cContext.Load(Pagelist);
                        _cContext.ExecuteQuery();

                        try
                        {
                            var x = _cContext.LoadQuery(Pagelist.ContentTypes.Where(ctD => ctD.Name == "Document"));
                            _cContext.ExecuteQuery();

                            ContentType ctx = (ContentType)x.FirstOrDefault();
                            ctx.DeleteObject();
                            _cContext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + "Document: " + ex.Message);
                            excelWriterScoringMatrixNew.Flush();
                        }

                        ContentTypeCollection contentTypeColl = Pagelist.ContentTypes;
                        _cContext.Load(contentTypeColl);
                        _cContext.ExecuteQuery();

                        foreach (ContentType eachcontenttype in contentTypeColl)
                        {
                            _cContext.Load(eachcontenttype);
                            _cContext.ExecuteQuery();

                            if (eachcontenttype.Name == "RicohContentType")
                            {
                                ContentType ricohcontentType1 = eachcontenttype;
                                _cContext.Load(ricohcontentType1);
                                _cContext.ExecuteQuery();

                                ricohcontentType1.ReadOnly = false;
                                ricohcontentType1.Update(false);
                                _cContext.Load(ricohcontentType1);
                                _cContext.ExecuteQuery();

                                break;
                            }
                        }
                        _cContext.Load(Pagelist);
                        _cContext.ExecuteQuery();

                        try
                        {
                            Field lookupField = null;
                            List list = _web.Lists.GetByTitle("Manage Categories"); //Categories  //CategoriesList
                            _cContext.Load(list);
                            _cContext.ExecuteQuery();
                            string schemaLookupField = "<Field Type='LookupMulti' Name='Categorization' StaticName='Categorization' DisplayName='Categorization' List = '" + list.Id + "' ShowField = 'Title' Mult = 'TRUE'/>";
                            lookupField = Pagelist.Fields.AddFieldAsXml(schemaLookupField, true, AddFieldOptions.AddFieldInternalNameHint);
                            Pagelist.Update();
                            _cContext.ExecuteQuery();

                            excelWriterScoringMatrixNew.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + "Success");
                            excelWriterScoringMatrixNew.Flush();
                        }
                        catch (Exception es)
                        {
                            excelWriterScoringMatrixNew.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + "Manage Categories: " + es.Message);
                            excelWriterScoringMatrixNew.Flush();
                        }

                        #endregion

                        #region Item Categorization Update

                        //ListItem _Item = Pagelist.GetItemById(_ItemID);
                        //_cContext.Load(_Item);
                        //_cContext.ExecuteQuery();

                        //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                        //FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                        //if (!string.IsNullOrEmpty(drImported["Categories"].ToString().Trim()))
                        //{
                        //    string _category = drImported["Categories"].ToString().Trim();
                        //    string[] _categories = _category.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        //    try
                        //    {

                        //        FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                        //        for (int i = 0; i <= _categories.Length - 1; i++)
                        //        {
                        //            string newValue = _categories[i].ToString();

                        //            if (_categories[i].ToString().Contains("$"))
                        //            {
                        //                newValue = _categories[i].ToString().Replace("$", ",");
                        //            }

                        //            int _cId = GetLookupIDs(newValue, _cContext, _web);

                        //            if (_cId != 0)
                        //            {
                        //                FieldLookupValue flv = new FieldLookupValue();
                        //                flv.LookupId = _cId;

                        //                lookupFieldValCollection.SetValue(flv, i);
                        //            }
                        //        }

                        //        if (lookupFieldValCollection.Length >= 1)
                        //        {
                        //            if (lookupFieldValCollection[0] != null)
                        //                _Item["Categorization"] = lookupFieldValCollection;
                        //        }

                        //        _Item.Update();
                        //        _cContext.Load(_Item);
                        //        _cContext.ExecuteQuery();
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        excelWriterScoringMatrixNew.WriteLine("ERROR:" + _SiteTitle + "," + _web.Url + "/" + Pagelist.RootFolder.Name + "/Dispform.aspx?id=" + _Item.Id + "," + drImported["Categories"].ToString().Trim());
                        //        excelWriterScoringMatrixNew.Flush();
                        //    }

                        //    try
                        //    {
                        //        _Item["Modified"] = Modified;
                        //        _Item["Editor"] = ModifiedBy;
                        //        _Item.Update();
                        //        _cContext.ExecuteQuery();

                        //        excelWriterScoringMatrixNew.WriteLine(_SiteTitle + "," + _web.Url + "/" + Pagelist.RootFolder.Name + "/Dispform.aspx?id=" + _Item.Id + "," + drImported["Categories"].ToString().Trim());
                        //        excelWriterScoringMatrixNew.Flush();
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        //excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Url + "," + OldTitle + "," + oTitle);
                        //        //excelWriterScoringMatrixNew.Flush();
                        //    }
                        //}

                        #endregion

                        #region FIXES

                        #region Site Assets

                        //try
                        //{
                        //    List olist = _web.Lists.GetByTitle("Site Assets");
                        //    _cContext.Load(olist);
                        //    _cContext.ExecuteQuery();

                        //    ContentTypeCollection contentTypeColls = olist.ContentTypes;
                        //    // ContentTypeCollection contentTypeColl = mweb.ContentTypes;

                        //    _cContext.Load(contentTypeColls);
                        //    _cContext.ExecuteQuery();
                        //    ContentType defaultcontentType = null;
                        //    ContentType ricohcontentType = null;

                        //    foreach (ContentType eachcontenttype in contentTypeColls)
                        //    {
                        //        _cContext.Load(eachcontenttype);
                        //        _cContext.ExecuteQuery();

                        //        if (eachcontenttype.Name == "RicohContentType")
                        //        {
                        //            ricohcontentType = eachcontenttype;
                        //            _cContext.Load(eachcontenttype);
                        //            _cContext.ExecuteQuery();
                        //            _cContext.Load(ricohcontentType);
                        //            ricohcontentType.ReadOnly = false;
                        //            ricohcontentType.Update(false);
                        //            _cContext.Load(ricohcontentType);
                        //            _cContext.ExecuteQuery();

                        //        }
                        //        else if (eachcontenttype.Name == "Document")
                        //        {
                        //            defaultcontentType = eachcontenttype;
                        //            _cContext.Load(eachcontenttype);
                        //            _cContext.ExecuteQuery();
                        //            _cContext.Load(defaultcontentType);
                        //            _cContext.ExecuteQuery();

                        //        }


                        //    }

                        //    bool isAttCTExist = contentTypeColls.Cast<ContentType>().Any(contentType => string.Equals(contentType.Name, "Document"));

                        //    if (isAttCTExist)
                        //    {

                        //        IList<ContentTypeId> reverseOrder = (from ct in contentTypeColls where ct.Name.Equals("Document", StringComparison.OrdinalIgnoreCase) select ct.Id).ToList();
                        //        olist.RootFolder.UniqueContentTypeOrder = reverseOrder;
                        //        olist.RootFolder.Update();
                        //        olist.Update();
                        //        _cContext.ExecuteQuery();
                        //    }

                        //    ricohcontentType.DeleteObject();
                        //    _cContext.ExecuteQuery();
                        //}
                        //catch (Exception ex)
                        //{

                        //}
                        #endregion

                        #region Enable ratings and Views

                        //ListCollection _Lists = _cContext.Web.Lists;
                        //_cContext.Load(_Lists);
                        //_cContext.ExecuteQuery();

                        //foreach (List list in _Lists)
                        //{
                        //    _cContext.Load(list);
                        //    _cContext.ExecuteQuery();

                        //    try
                        //    {
                        //        if (list.Title == "1_Uploaded Files" || list.Title == "Events" ||
                        //            list.Title == "Announcements" || list.Title == "Tasks" ||
                        //            list.Title == "Posts" || list.Title == "Discussions" ||
                        //            list.Title == "SiteHistory" || list.Title == "Ideas" ||
                        //            list.Title == "2_Documents and Pages" || list.Title == "Site Assets" ||
                        //            list.Title == "Status")
                        //        {

                        //            EnableRating(_cContext, list.Title);
                        //        }

                        // if (list.Title == "2_Documents and Pages")
                        //                v.ViewFields.Add("LinkFilename");
                        //                v.ViewFields.Add("Created");
                        //                v.ViewFields.Add("Created By");
                        //                v.ViewFields.Add("Modified");
                        //                v.ViewFields.Add("Modified By");
                        //                v.ViewFields.Add("Tags");
                        //                v.ViewFields.Add("Categorization");
                        //                v.ViewFields.Add("CheckoutUser");
                        //                v.Update();
                        //                _cContext.ExecuteQuery();

                        //                Folder docFolder = null;
                        //                try
                        //                {
                        //                    docFolder = Pageslist.RootFolder.Folders.GetByUrl("Documents");
                        //                    _cContext.Load(docFolder);
                        //                    _cContext.ExecuteQuery();
                        //                }
                        //                catch (Exception ex)
                        //                {
                        //                    docFolder = Pageslist.RootFolder.Folders.Add("Documents");
                        //                    _cContext.Load(docFolder);
                        //                    _cContext.Load(docFolder, p => p.ServerRelativeUrl);
                        //                    _cContext.ExecuteQuery();
                        //                }

                        //                try
                        //                {
                        //                    Pageslist.DraftVersionVisibility = DraftVisibilityType.Reader;
                        //                    Pageslist.Update();
                        //                    _cContext.ExecuteQuery();
                        //                }
                        //                catch (Exception ex)
                        //                {

                        //                }
                        //            }
                        //            catch (Exception ex)
                        //            {

                        //            }
                        //        }

                        //if (list.Title == "Ideas")
                        //{
                        //    try
                        //    {
                        //        List idealist = _cContext.Web.Lists.GetByTitle("Ideas");
                        //        _cContext.Load(idealist);
                        //        _cContext.ExecuteQuery();

                        //        ViewCollection ViewColl = idealist.Views;
                        //        _cContext.Load(ViewColl);
                        //        _cContext.ExecuteQuery();

                        //        Microsoft.SharePoint.Client.View v = ViewColl[6];
                        //        _cContext.Load(v);
                        //        _cContext.ExecuteQuery();

                        //        v.DeleteObject();
                        //        _cContext.ExecuteQuery();
                        //    }
                        //    catch (Exception ex)
                        //    {

                        //    }
                        //}
                        //        if (list.Title == "Status")
                        //        {
                        //            try
                        //            {
                        //                List Statuslist = _cContext.Web.Lists.GetByTitle("Status");
                        //                _cContext.Load(Statuslist);
                        //                _cContext.ExecuteQuery();

                        //                ViewCollection ViewColl = Statuslist.Views;
                        //                _cContext.Load(ViewColl);
                        //                _cContext.ExecuteQuery();

                        //                Microsoft.SharePoint.Client.View v = ViewColl[0];
                        //                _cContext.Load(v);
                        //                _cContext.ExecuteQuery();

                        //                v.ViewFields.RemoveAll();
                        //                v.Update();
                        //                _cContext.ExecuteQuery();

                        //                v.ViewFields.Add("StatusDescription");
                        //                v.Update();
                        //                _cContext.ExecuteQuery();
                        //            }
                        //            catch (Exception ex1)
                        //            {

                        //            }
                        //        }
                        //    }
                        //    catch (Exception ex1)
                        //    {
                        //        continue;
                        //    }
                        //}
                        #endregion 

                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    //excelWriterScoringMatrixNew.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + drImported["SiteURL"].ToString().Trim() + "/" + "1_Uploaded Files/ Dispform.aspx?id=" + drImported["ItemID"].ToString().Trim() + "," + drImported["Categories"].ToString().Trim());
                    //excelWriterScoringMatrixNew.Flush();

                    excelWriterScoringMatrixNew.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + "ERROR: " + ex.Message.ToString());
                    excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        public int GetLookupIDs(string _lookvalue, ClientContext _gContext, Web _gWeb)
        {
            int _lookupValue = 0;
            try
            {
                if (_lookvalue != string.Empty && _lookvalue != null)
                {

                    using (ClientContext _cContext = _gContext)
                    {

                        Web _Web = _gWeb;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        ListCollection _Lists = _Web.Lists;
                        _cContext.Load(_Lists);
                        _cContext.ExecuteQuery();

                        // List _list = _Lists.GetByTitle("Categorías");
                        List _list = _Lists.GetByTitle("Manage Categories");
                        _cContext.Load(_list);

                        CamlQuery _q = new CamlQuery();
                        _q.ViewXml = string.Empty;

                        ListItemCollection _items = _list.GetItems(_q);
                        _cContext.Load(_items);
                        _cContext.ExecuteQuery();
                        bool itmexist = false;
                        foreach (ListItem _item in _items)
                        {
                            try
                            {
                                _cContext.Load(_item);
                                _cContext.ExecuteQuery();
                                if (_item["Title"].ToString() == _lookvalue)
                                {
                                    _lookupValue = _item.Id;
                                    itmexist = true;
                                    break;
                                }
                            }
                            catch (Exception ex) { }
                        }

                        if (itmexist == false)
                        {

                            ListItemCreationInformation _ItemInfo = new ListItemCreationInformation();
                            ListItem _Item = _list.AddItem(_ItemInfo);

                            _Item["Title"] = System.Web.HttpUtility.HtmlDecode(_lookvalue);

                            _Item.Update();
                            _cContext.Load(_Item);
                            _cContext.ExecuteQuery();
                            _lookupValue = _Item.Id;
                        }
                    }
                }
            }

            catch (Exception ex)
            {
            }
            return _lookupValue;
        }

        public int GetLookupIDsManageTag(string _lookvalue, ClientContext _gContext, Web _gWeb)
        {
            int _lookupValue = 0;
            try
            {
                if (_lookvalue != string.Empty && _lookvalue != null)
                {

                    using (ClientContext _cContext = _gContext)
                    {
                        Web _Web = _gWeb;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        ListCollection _Lists = _Web.Lists;
                        _cContext.Load(_Lists);
                        _cContext.ExecuteQuery();

                        List _list = _Lists.GetByTitle("Manage Tag");
                        _cContext.Load(_list);



                        CamlQuery _q = new CamlQuery();
                        //_q.ViewXml = string.Empty;

                        _q.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>" + _lookvalue + "</Value></Eq></Where></Query></View>";

                        ListItemCollection _items = _list.GetItems(_q);
                        _cContext.Load(_items);
                        _cContext.ExecuteQuery();
                        bool itmexist = false;

                        foreach (ListItem _item in _items)
                        {
                            try
                            {
                                _cContext.Load(_item);
                                _cContext.ExecuteQuery();
                                if (_item["Title"].ToString().ToLower() == _lookvalue.ToLower())
                                {
                                    _lookupValue = _item.Id;
                                    itmexist = true;
                                    break;
                                }
                            }
                            catch (Exception ex) { }
                        }

                        if (itmexist == false)
                        {

                            ListItemCreationInformation _ItemInfo = new ListItemCreationInformation();
                            ListItem _Item = _list.AddItem(_ItemInfo);

                            _Item["Title"] = System.Web.HttpUtility.HtmlDecode(_lookvalue);

                            _Item.Update();
                            _cContext.Load(_Item);
                            _cContext.ExecuteQuery();
                            _lookupValue = _Item.Id;
                        }
                    }
                }
            }

            catch (Exception ex)
            {
            }
            return _lookupValue;
        }
        private void button16_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[2] { new DataColumn("DID", typeof(string)), new DataColumn("URL", typeof(string)) });

            //Read the contents of CSV file.  
            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            //Execute a loop over the rows.  
            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;
                    //Execute a loop over the columns.  
                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            DataTable dtJiveObjects = new DataTable();
            dtJiveObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("DID", typeof(string)), new DataColumn("Modified", typeof(string)), new DataColumn("ModifiedBy", typeof(string)) });

            //Read the contents of CSV file.  
            string csvData1 = System.IO.File.ReadAllText(textBox3.Text);

            //Execute a loop over the rows.  
            foreach (string row in csvData1.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtJiveObjects.Rows.Add();
                    int i = 0;
                    //Execute a loop over the columns.  
                    foreach (string cell in row.Split(','))
                    {
                        dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocumentsModfiedReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("DID" + "," + "URL" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;
            string[] SiteSplit = new string[] { "/Pages/" };
            string[] FileSplit = new string[] { "/Documents/" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string Modified = string.Empty;
                    var ModifiedBy = string.Empty;

                    string _SiteTitle = drImported["URL"].ToString().Split(SiteSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    string _FilePath = drImported["URL"].ToString().Split(FileSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();

                    this.Text = (count).ToString() + " : " + _SiteTitle;

                    count++;

                    foreach (DataRow drJive in dtJiveObjects.Rows)
                    {
                        if (drImported["DID"].ToString().Trim() == drJive["DID"].ToString().Trim())
                        {
                            Modified = drJive["Modified"].ToString().Trim();
                            ModifiedBy = drJive["ModifiedBy"].ToString().Trim();
                            break;
                        }
                    }

                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_SiteTitle, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        List _List = null;

                        try
                        {
                            _List = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages"); ;
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        { }

                        if (_List != null)
                        {
                            try
                            {
                                clientcontext.Load(_List.RootFolder);
                                clientcontext.ExecuteQuery();

                                Folder docFolder = null;

                                try
                                {
                                    docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                    clientcontext.Load(docFolder);
                                    clientcontext.ExecuteQuery();
                                }
                                catch (Exception ex)
                                { }

                                if (docFolder != null)
                                {
                                    ListItem _Item = docFolder.Files.GetByUrl(_FilePath).ListItemAllFields;
                                    clientcontext.Load(_Item);
                                    clientcontext.ExecuteQuery();

                                    //DateTime dtModified = DateTime.Parse(Modified).ToUniversalTime();//.ToString("MM/dd/yyyy hh:mm tt");// (Modified);
                                    //DateTime dtModified = DateTime.ParseExact(Modified, @"dd-MM-yyyy hh:mm", System.Globalization.CultureInfo.InvariantCulture);//Convert.ToDateTime(Modified);//
                                    DateTime dtModified = getdateformat(Modified);

                                    User CreatedUser = default(User);
                                    try
                                    {
                                        CreatedUser = oWeb.EnsureUser(ModifiedBy);
                                        clientcontext.Load(CreatedUser);
                                        clientcontext.ExecuteQuery();

                                    }
                                    catch (Exception ex)
                                    {
                                        CreatedUser = oWeb.EnsureUser("Rspaceadmin@rsharepoint.onmicrosoft.com");
                                        clientcontext.Load(CreatedUser);
                                        clientcontext.ExecuteQuery();
                                    }

                                    //Updating Ownership Information
                                    FieldUserValue CreatedUserValue = new FieldUserValue();
                                    CreatedUserValue.LookupId = CreatedUser.Id;

                                    _Item["Modified"] = dtModified;
                                    _Item["Editor"] = CreatedUserValue;
                                    _Item.Update();
                                    clientcontext.ExecuteQuery();

                                    excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Success");
                                    excelWriterScoringMatrixNew.Flush();
                                }
                            }
                            catch (Exception ex)
                            {
                                excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Failure due to : " + ex.Message);
                                excelWriterScoringMatrixNew.Flush();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Failure due to : " + ex.Message);
                    excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }
            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button17_Click(object sender, EventArgs e)
        {
            #region AD Groups CSV Reading

            //lstADGroupsColl.Clear();

            //if (!string.IsNullOrEmpty(textBox3.Text))
            //{
            //    StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox3.Text));

            //    while (!sr.EndOfStream)
            //    {
            //        try
            //        {
            //            lstADGroupsColl.Add(sr.ReadLine().Trim().ToLower());
            //        }
            //        catch
            //        {
            //            continue;
            //        }
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please browse the path for ADGroups.csv");
            //}

            #endregion

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteAssetsViewReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("URL" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            //List<string> ListNames = new List<string>();
            //ListNames.Add("Site Assets");
            //ListNames.Add("2_Documents and Pages");
            //ListNames.Add("1_Uploaded Files");
            //ListNames.Add("Discussions");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(),
                                "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        //try
                        //{
                        //    List Pagelist = _Lists.GetByTitle("Posts");
                        //    clientcontext.Load(Pagelist);
                        //    clientcontext.ExecuteQuery();

                        //    ViewCollection ViewColl = Pagelist.Views;
                        //    clientcontext.Load(ViewColl);
                        //    clientcontext.ExecuteQuery();

                        //    Microsoft.SharePoint.Client.View v = ViewColl[0];
                        //    clientcontext.Load(v);
                        //    clientcontext.ExecuteQuery();

                        //    v.ViewFields.RemoveAll();
                        //    v.Update();
                        //    clientcontext.ExecuteQuery();                            

                        //    v.ViewFields.Add("LinkTitle");
                        //    v.ViewFields.Add("Created");
                        //    v.ViewFields.Add("Published");
                        //    v.ViewFields.Add("Category");
                        //    v.ViewFields.Add("NumComments");
                        //    v.ViewFields.Add("Edit");
                        //    v.ViewFields.Add("Categorization");
                        //    v.ViewFields.Add("LikesCount");
                        //    v.Update();
                        //    clientcontext.ExecuteQuery();
                        //}
                        //catch (Exception ex)
                        //{
                        //    //continue;
                        //}

                        //foreach (List list in _Lists)
                        {

                            // EnableRating(clientcontext, listName);
                            try
                            {

                                #region VIEW for 2_Documents and Pages

                                List Pagelist = _Lists.GetByTitle("Site Assets");
                                clientcontext.Load(Pagelist);
                                clientcontext.ExecuteQuery();

                                //try
                                //{
                                //    Pagelist.EnableModeration = false;
                                //    Pagelist.Update();
                                //    clientcontext.ExecuteQuery();

                                //    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Success");
                                //    excelWriterScoringMatrixNew.Flush();
                                //}
                                //catch (Exception ex)
                                //{

                                //}

                                //ViewCollection ViewColl = Pagelist.Views;
                                //clientcontext.Load(ViewColl);
                                //clientcontext.ExecuteQuery();

                                //Microsoft.SharePoint.Client.View v = ViewColl[0];
                                //clientcontext.Load(v);
                                //clientcontext.ExecuteQuery();

                                //v.ViewFields.RemoveAll();
                                //v.Update();
                                //clientcontext.ExecuteQuery();

                                //v.ViewFields.Add("DocIcon");
                                //v.ViewFields.Add("Title");
                                //v.ViewFields.Add("LinkFilename");
                                //v.ViewFields.Add("Created");
                                //v.ViewFields.Add("Created By");
                                //v.ViewFields.Add("Modified");
                                //v.ViewFields.Add("Modified By");
                                //v.ViewFields.Add("Tags");
                                //v.ViewFields.Add("Categorization");
                                //v.ViewFields.Add("CheckoutUser");
                                //v.Update();
                                //clientcontext.ExecuteQuery();

                                ViewCollection ViewColl = Pagelist.Views;
                                clientcontext.Load(ViewColl);
                                clientcontext.ExecuteQuery();

                                Microsoft.SharePoint.Client.View v = ViewColl[0];
                                clientcontext.Load(v);
                                clientcontext.ExecuteQuery();

                                v.ViewFields.RemoveAll();
                                v.Update();
                                clientcontext.ExecuteQuery();

                                v.ViewFields.Add("DocIcon");
                                v.ViewFields.Add("Title");
                                v.ViewFields.Add("LinkFilename");
                                v.ViewFields.Add("Created");
                                v.ViewFields.Add("Created By");
                                v.ViewFields.Add("Modified");
                                v.ViewFields.Add("Modified By");
                                //v.ViewFields.Add("Tags");
                                //v.ViewFields.Add("Categorization");
                                v.ViewFields.Add("CheckoutUser");
                                v.Update();
                                clientcontext.ExecuteQuery();

                                #endregion


                                //bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, listName));

                                #region MyRegion
                                //if (_dListExist)
                                //{
                                //if (listName == "Ideas")
                                //{
                                //    // try
                                //    {
                                //        List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                //        clientcontext.Load(Pagelist);
                                //        clientcontext.ExecuteQuery();

                                //        ViewCollection ViewColl = Pagelist.Views;
                                //        clientcontext.Load(ViewColl);
                                //        clientcontext.ExecuteQuery();

                                //        //Microsoft.SharePoint.Client.View v = ViewColl[0];
                                //        Microsoft.SharePoint.Client.View v = ViewColl[6];
                                //        clientcontext.Load(v);
                                //        clientcontext.ExecuteQuery();

                                //        v.DeleteObject();
                                //        clientcontext.ExecuteQuery();

                                //        //v.ViewFields.RemoveAll();
                                //        //v.Update();
                                //        //clientcontext.ExecuteQuery();

                                //        //v.ViewFields.Add("StatusDescription");
                                //        //v.Update();
                                //        //clientcontext.ExecuteQuery();
                                //    }
                                //}

                                //    if ((list.BaseTemplate.ToString() == "109") && (listName != "Photos" || listName == "Images"))
                                //    {
                                //        #region Commented

                                //        FieldCollection FldColl = list.Fields;
                                //        clientcontext.Load(FldColl);
                                //        clientcontext.ExecuteQuery();
                                //        bool TagCateExist = false;

                                //        foreach (Field tagField in FldColl)
                                //        {
                                //            clientcontext.Load(tagField);
                                //            clientcontext.ExecuteQuery();

                                //            if (tagField.Title.ToLower() == "tags" || tagField.Title.ToLower() == "categorization")
                                //            {
                                //                TagCateExist = true;
                                //                break;
                                //            }
                                //        }

                                #endregion

                                //if (_dListExist)
                                //{
                                //    List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                //    clientcontext.Load(Pagelist);
                                //    clientcontext.ExecuteQuery();

                                //    ViewCollection ViewColl = Pagelist.Views;
                                //    clientcontext.Load(ViewColl);
                                //    clientcontext.ExecuteQuery();

                                //    Microsoft.SharePoint.Client.View v = ViewColl[0];
                                //    clientcontext.Load(v);
                                //    clientcontext.ExecuteQuery();

                                //    v.ViewFields.RemoveAll();
                                //    v.Update();
                                //    clientcontext.ExecuteQuery();

                                //2_Documents and Pages

                                //    v.ViewFields.Add("DocIcon");
                                //    v.ViewFields.Add("Title");
                                //    v.ViewFields.Add("LinkFilename");
                                //    v.ViewFields.Add("Created");
                                //    v.ViewFields.Add("Created By");
                                //    v.ViewFields.Add("Modified");
                                //    v.ViewFields.Add("Modified By");
                                //    v.ViewFields.Add("Tags");
                                //    v.ViewFields.Add("Categorization");
                                //    v.ViewFields.Add("Number of Likes");
                                //    v.Update();
                                //    clientcontext.ExecuteQuery();
                                //}
                                //}                                    

                                #region Commented

                                // // else
                                //  {
                                //List Pagelist = clientcontext.Web.Lists.GetByTitle(listName);
                                //clientcontext.Load(Pagelist);
                                //clientcontext.ExecuteQuery();

                                //ViewCollection ViewColl = Pagelist.Views;
                                //clientcontext.Load(ViewColl);
                                //clientcontext.ExecuteQuery();

                                //Microsoft.SharePoint.Client.View v = ViewColl[0];
                                //clientcontext.Load(v);
                                //clientcontext.ExecuteQuery();

                                //v.ViewFields.RemoveAll();
                                //v.Update();
                                //clientcontext.ExecuteQuery();

                                //v.ViewFields.Add("DocIcon");
                                //v.ViewFields.Add("Title");
                                //v.ViewFields.Add("LinkFilename");
                                //v.ViewFields.Add("Created");
                                //v.ViewFields.Add("Created By");
                                //v.ViewFields.Add("Modified");
                                //v.ViewFields.Add("Modified By");
                                //v.ViewFields.Add("Tags");
                                //v.ViewFields.Add("Categorization");
                                //v.ViewFields.Add("CheckoutUser");
                                //v.Update();
                                //clientcontext.ExecuteQuery();
                                //  }

                                //  //Pagelist.ContentTypesEnabled = true;
                                //  //Pagelist.Update();
                                ////  clientcontext.ExecuteQuery();

                                #endregion
                                //}
                            }
                            catch (Exception ex)
                            {
                                // continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }

            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        public static void EnableRating(ClientContext ctx, string listTitle)
        {
            //OfficeDevPnP.Core.AuthenticationManager authMgr = new OfficeDevPnP.Core.AuthenticationManager();

            //string siteUrl = "https://**********.sharepoint.com/sites/DeveloperSite";
            //string userName = "Sathish@********.onmicrosoft.com";
            //string password = "********";
            //string listTitle = "MyDocLib";

            try
            {
                Guid RatingsFieldGuid_AverageRating = new Guid("5a14d1ab-1513-48c7-97b3-657a5ba6c742");
                Guid RatingsFieldGuid_RatingCount = new Guid("b1996002-9167-45e5-a4df-b2c41c6723c7");
                Guid RatingsFieldGuid_RatedBy = new Guid("4D64B067-08C3-43DC-A87B-8B8E01673313");
                Guid RatingsFieldGuid_Ratings = new Guid("434F51FB-FFD2-4A0E-A03B-CA3131AC67BA");
                Guid LikeFieldGuid_LikedBy = new Guid("2cdcd5eb-846d-4f4d-9aaf-73e8e73c7312");
                Guid LikeFieldGuid_LikeCount = new Guid("6e4d832b-f610-41a8-b3e0-239608efda41");


                ctx.Load(ctx.Web);
                ctx.ExecuteQuery();
                List filesLibrary = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(filesLibrary);
                ctx.Load(filesLibrary.RootFolder, p => p.Properties);
                ctx.ExecuteQuery();
                filesLibrary.RootFolder.Properties["Ratings_VotingExperience"] = "Likes";
                filesLibrary.RootFolder.Update();
                ctx.ExecuteQuery();

                EnsureField(filesLibrary, RatingsFieldGuid_RatingCount, ctx);
                EnsureField(filesLibrary, RatingsFieldGuid_RatedBy, ctx);
                EnsureField(filesLibrary, RatingsFieldGuid_Ratings, ctx);
                EnsureField(filesLibrary, RatingsFieldGuid_AverageRating, ctx);
                EnsureField(filesLibrary, LikeFieldGuid_LikedBy, ctx);
                EnsureField(filesLibrary, LikeFieldGuid_LikeCount, ctx);

                filesLibrary.Update();
                ctx.ExecuteQuery();
                ctx.Load(filesLibrary, view => view.DefaultView);
                ctx.ExecuteQuery();
                var defaultView = filesLibrary.DefaultView;
                defaultView.ViewFields.Add("LikesCount");
                defaultView.Update();
                ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {

            }

        }
        private static Field EnsureField(List list, Guid fieldId, ClientContext _context)
        {
            Field field = null;
            try
            {
                FieldCollection fields = list.Fields;

                FieldCollection availableFields = list.ParentWeb.AvailableFields;
                field = availableFields.GetById(fieldId);

                _context.Load(fields);
                _context.Load(field, p => p.SchemaXmlWithResourceTokens, p => p.Id, p => p.InternalName, p => p.StaticName);
                _context.ExecuteQuery();

                if (!fields.Any(p => p.Id == fieldId))
                {

                    var newField = fields.AddFieldAsXml(field.SchemaXmlWithResourceTokens, false, AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddToAllContentTypes);
                    return newField;
                }
            }
            catch (Exception ex)
            { }
            return field;
        }
        private void button18_Click(object sender, EventArgs e)
        {
            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "FileCategories" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "Versions");
            excelWriterScoringMatrixNew.Flush();

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}


            //StreamWriter excelWriterScoringMatrixNew = null;

            //excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            //excelWriterScoringMatrixNew.WriteLine("Filename" + "," + "URL" + "," + "Owners" + "," + "Built-in-Groups" + "," + "AD Groups" + "," + "Start Time" + "," + "End Date" + "," + "Remarks");
            //excelWriterScoringMatrixNew.Flush();

            //List<string> ListNames = new List<string>();
            //ListNames.Add("Site Assets");
            //ListNames.Add("2_Documents and Pages");
            //ListNames.Add("1_Uploaded Files");
            //ListNames.Add("Discussions");

            //lstNameColl.Add("1_Uploaded Files");
            //lstNameColl.Add("Events");
            //lstNameColl.Add("Announcements");
            //lstNameColl.Add("Tasks");
            //lstNameColl.Add("Posts");
            //lstNameColl.Add("Discussions");
            //lstNameColl.Add("SiteHistory");
            //lstNameColl.Add("Ideas");
            //lstNameColl.Add("Manage Categories");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(
                        lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        List list = _Lists.GetByTitle("2_Documents and Pages");
                        clientcontext.Load(list);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            //if (list.EnableVersioning == false)
                            //{
                            //list.EnableVersioning = true;
                            list.UpdateListVersioning(true, true, true);// = true;
                            //list.Update();
                            clientcontext.ExecuteQuery();

                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim()
                                + "," + "Yes");
                            excelWriterScoringMatrixNew.Flush();
                            //}
                        }
                        catch (Exception ex)
                        {

                        }

                        List list1 = _Lists.GetByTitle("1_Uploaded Files");
                        clientcontext.Load(list1);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            //if (list.EnableVersioning == false)
                            //{
                            //list.EnableVersioning = true;
                            // list.UpdateListVersioning(true, true, true);// = true;
                            //list.Update();
                            list1.EnableVersioning = true;
                            list1.Update();
                            clientcontext.ExecuteQuery();

                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim()
                                + "," + "Yes");
                            excelWriterScoringMatrixNew.Flush();
                            //}
                        }
                        catch (Exception ex)
                        {

                        }


                        // foreach (List list in _Lists)
                        {

                            // List list = _Lists.GetByTitle("Discussions");
                            // clientcontext.Load(list);
                            // clientcontext.ExecuteQuery();


                            try
                            {
                                //if (list.Title == "1_Uploaded Files" ||  list.Title == "Events" ||
                                //    list.Title == "Announcements" ||list.Title == "Tasks" ||
                                //    list.Title == "Posts" ||list.Title == "Discussions" ||
                                //    list.Title == "SiteHistory" || list.Title == "Ideas" ||
                                //    list.Title == "2_Documents and Pages" ||list.Title == "Site Assets" ||
                                //    list.Title == "Status")

                                //if (list.Title == "2_Documents and Pages")
                                {

                                    //  EnableRating(clientcontext, list.Title);


                                    //try
                                    //{
                                    //    list.DraftVersionVisibility = DraftVisibilityType.Reader;
                                    //    list.Update();
                                    //    clientcontext.ExecuteQuery();
                                    //}
                                    //catch (Exception ex)
                                    //{

                                    //}

                                }
                            }
                            catch (Exception ex1)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim()
                               + "," + ex.Message.ToString());
                    excelWriterScoringMatrixNew.Flush();
                    continue;
                }

            }
            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button19_Click(object sender, EventArgs e)
        {
            AuthenticationManager authManager = new AuthenticationManager();
            using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant("https://rsharepoint.sharepoint.com/sites/rspace", "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
            {
                // Get the TaxonomySession

                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientcontext);

                // Get the term store by name
                TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");

                // Get the term group by Name
                TermGroup termGroup = termStore.Groups.GetByName("RicohTags");

                // Get the term set by Name
                TermSet termSet = termGroup.TermSets.GetByName("TagsTermSet");

                // Get all the terms 
                TermCollection termColl = termSet.Terms;
                clientcontext.Load(termColl);
                clientcontext.ExecuteQuery();

                //Term tm = termColl.GetByName("acta de aceptaciÃ³n de entregables portuguÃ©s;");
                //clientcontext.Load(tm);
                //clientcontext.ExecuteQuery();

                //byte[] bytes1 = Encoding.Default.GetBytes(tm.Name.ToString());
                //string oTitle = Encoding.UTF8.GetString(bytes1);

                //tm.Name = oTitle;
                //termStore.CommitAll();
                //clientcontext.ExecuteQuery();


                //// Loop through all the terms

                foreach (Term tm in termColl)
                {
                    clientcontext.Load(tm);
                    clientcontext.ExecuteQuery();

                    if (tm.Name.StartsWith("acta de aceptac"))
                    {
                        byte[] bytes1 = Encoding.Default.GetBytes(tm.Name.ToString());
                        string oTitle = Encoding.UTF8.GetString(bytes1);

                        tm.Name = oTitle;
                        termStore.CommitAll();
                        clientcontext.ExecuteQuery();

                    }
                }
            }
        }
        private void button20_Click(object sender, EventArgs e)
        {

        }
        private void button20_Click_1(object sender, EventArgs e)
        {
            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocumentsCT" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("DID" + "," + "URL" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            StreamWriter excelWriterScoringMatrixNew1 = null;
            excelWriterScoringMatrixNew1 = System.IO.File.CreateText(textBox2.Text + "\\" + "ContentApprovalReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew1.WriteLine("URL" + "," + "Status");
            excelWriterScoringMatrixNew1.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        try
                        {

                            List list = _Lists.GetByTitle("1_Uploaded Files");
                            clientcontext.Load(list);
                            clientcontext.ExecuteQuery();

                            ContentTypeCollection contentTypeColls = list.ContentTypes;

                            clientcontext.Load(contentTypeColls);
                            clientcontext.ExecuteQuery();
                            ContentType defaultcontentType = null;

                            foreach (ContentType eachcontenttype in contentTypeColls)
                            {
                                clientcontext.Load(eachcontenttype);
                                clientcontext.ExecuteQuery();

                                if (eachcontenttype.Name == "Document")
                                {
                                    defaultcontentType = eachcontenttype;
                                    clientcontext.Load(eachcontenttype);
                                    clientcontext.ExecuteQuery();
                                    clientcontext.Load(defaultcontentType);
                                    clientcontext.ExecuteQuery();
                                }
                            }

                            bool isAttCTExist = contentTypeColls.Cast<ContentType>().Any(contentType => string.Equals(contentType.Name, "Document"));

                            if (isAttCTExist)
                            {
                                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + list.Title + "," + "Exist");
                                excelWriterScoringMatrixNew.Flush();
                            }
                        }
                        catch (Exception ex)
                        {

                        }

                        try
                        {
                            List Announcementslist = _Lists.GetByTitle("Announcements");
                            clientcontext.Load(Announcementslist);
                            clientcontext.ExecuteQuery();

                            ViewCollection ViewColl = Announcementslist.Views;
                            clientcontext.Load(ViewColl);
                            clientcontext.ExecuteQuery();

                            Microsoft.SharePoint.Client.View v = ViewColl[0];
                            clientcontext.Load(v);
                            clientcontext.ExecuteQuery();

                            v.ViewFields.RemoveAll();
                            v.Update();
                            clientcontext.ExecuteQuery();

                            v.ViewFields.Add("LinkTitle");
                            v.ViewFields.Add("Modified");
                            v.ViewFields.Add("LikesCount");
                            v.Update();
                            clientcontext.ExecuteQuery();

                        }
                        catch (Exception ex)
                        {

                        }

                        try
                        {
                            List Pagelist = _Lists.GetByTitle("2_Documents and Pages");
                            clientcontext.Load(Pagelist);
                            clientcontext.ExecuteQuery();

                            if (Pagelist.EnableModeration.ToString().ToLower() == "true")
                            {
                                excelWriterScoringMatrixNew1.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Success");
                                excelWriterScoringMatrixNew1.Flush();

                                Pagelist.EnableModeration = false;
                                Pagelist.Update();
                                clientcontext.ExecuteQuery();
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            excelWriterScoringMatrixNew1.Flush();
            excelWriterScoringMatrixNew1.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button21_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            StreamWriter excelWriterGUID = null;
            excelWriterGUID = System.IO.File.CreateText(textBox2.Text + "\\" + "GUIDReport" + ".csv");
            excelWriterGUID.WriteLine("SiteURL" + "," + "GUID");
            excelWriterGUID.Flush();

            StreamWriter excelWriterERROR = null;
            excelWriterERROR = System.IO.File.CreateText(textBox2.Text + "\\" + "Errorlog" + ".csv");
            excelWriterERROR.WriteLine("SiteURL" + "," + "Object(Name)" + "," + "Details");
            excelWriterERROR.Flush();

            List<string> ListNames = new List<string>();

            // ListNames.Add("Documents");
            ListNames.Add("1_Uploaded Files");
            ListNames.Add("2_Documents and Pages");
            ListNames.Add("Discussions");
            ListNames.Add("Events");
            ListNames.Add("Messages");
            ListNames.Add("Posts");
            ListNames.Add("Site Assets");
            ListNames.Add("SiteHistory");
            ListNames.Add("Tasks");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "adam.a@VerinonTechnology.onmicrosoft.com", "verinon@2018"))//"svc-jivemigration1@rsharepoint.onmicrosoft.com", "Vak52950"))
                    {

                        Web _Web = _cContext.Web;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        excelWriterGUID.WriteLine(_cContext.Web.Url + "," + _Web.Id);
                        excelWriterGUID.Flush();

                        StreamWriter excelWriterTagsReport = null;
                        excelWriterTagsReport = System.IO.File.CreateText(textBox2.Text + "\\" + _Web.Id.ToString() + "_TagsReport" + ".csv");
                        excelWriterTagsReport.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "Tags");
                        excelWriterTagsReport.Flush();

                        StreamWriter excelWriterUniqueTags = null;
                        excelWriterUniqueTags = System.IO.File.CreateText(textBox2.Text + "\\" + _Web.Id.ToString() + "_UniqueTagsReport" + ".csv");
                        excelWriterUniqueTags.WriteLine("SiteURL" + "," + "Tags");
                        excelWriterUniqueTags.Flush();

                        List _List = null;

                        List<string> _UniqueTags = new List<string>();
                        string _strUniqueTags = string.Empty;

                        foreach (string ls in ListNames)
                        {
                            try
                            {
                                _List = _cContext.Web.Lists.GetByTitle(ls);
                                _cContext.Load(_List);
                                _cContext.ExecuteQuery();

                                bool tagsFileldExist = _List.FieldExistsByName("Tags");

                                #region OLD

                                //_List.EnableVersioning = false;
                                //_List.Update();
                                //_cContext.ExecuteQuery(); 

                                #endregion

                                if (tagsFileldExist)
                                {
                                    CamlQuery camlQuery = new CamlQuery();
                                    camlQuery.ViewXml = "<View Scope='RecursiveAll'></View>";//<RowLimit>5000</RowLimit>

                                    ListItemCollection listItems = _List.GetItems(camlQuery);
                                    _cContext.Load(listItems);
                                    _cContext.ExecuteQuery();

                                    foreach (ListItem oItem in listItems)
                                    {
                                        try
                                        {
                                            string Tags = string.Empty;
                                            _cContext.Load(oItem);
                                            _cContext.ExecuteQuery();

                                            TaxonomyFieldValueCollection taxFieldValues = oItem["Tags"] as TaxonomyFieldValueCollection;

                                            #region OLD

                                            //TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(_Web.Context);
                                            //_cContext.Load(taxonomySession.TermStores);
                                            //_cContext.ExecuteQuery();

                                            //TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");
                                            //_cContext.Load(termStore);
                                            //_cContext.ExecuteQuery();

                                            //_cContext.Load(termStore.Groups);
                                            //_cContext.ExecuteQuery();

                                            //TermGroup group = termStore.Groups.GetByName("RicohTags");

                                            //_cContext.Load(group);
                                            //_cContext.ExecuteQuery();


                                            //_cContext.Load(group.TermSets);
                                            //_cContext.ExecuteQuery();


                                            //TermSet termSet = group.TermSets.GetByName("TagsTermSet");
                                            //_cContext.Load(termSet);
                                            //_cContext.ExecuteQuery();

                                            //Field _taxnomyField = _List.Fields.GetByTitle("Tags");
                                            //_cContext.Load(_taxnomyField);
                                            //_cContext.ExecuteQuery();

                                            //TaxonomyField txField = _cContext.CastTo<TaxonomyField>(_taxnomyField);
                                            //_cContext.Load(txField);
                                            //_cContext.ExecuteQuery();
                                            //TaxonomyFieldValue termValue = null;

                                            //TaxonomyFieldValueCollection termValues = null;

                                            //string termValueString = string.Empty;

                                            //string mtermId = string.Empty;
                                            //termValueString = string.Empty;
                                            //string termId = string.Empty; 

                                            #endregion

                                            foreach (TaxonomyFieldValue tv in taxFieldValues)
                                            {
                                                if (tv != null)
                                                {
                                                    Tags += tv.Label.ToString() + "|";

                                                    if (!_UniqueTags.Contains(tv.Label.ToString()))
                                                    {
                                                        _UniqueTags.Add(tv.Label.ToString());
                                                    }

                                                    #region OLD

                                                    //if (tv.Label.ToString().Contains("Ã") ||
                                                    //    tv.Label.ToString().Contains("Â"))
                                                    //{
                                                    //    Tags = tv.Label.ToString();


                                                    //mtermId = GetTermIdForTerm(Tags, termSet.Id, termSet, termStore, _cContext);
                                                    //if (!string.IsNullOrEmpty(mtermId))
                                                    //    termValueString += "1033" + ";#" + Tags + "|" + mtermId + ";#";


                                                    //termValueString = termValueString.Remove(termValueString.Length - 2);
                                                    //termValues = new TaxonomyFieldValueCollection(_cContext, termValueString, txField);

                                                    //txField.SetFieldValueByValueCollection(oItem, termValues);
                                                    //} 

                                                    #endregion
                                                }
                                            }

                                            #region OLD

                                            //oItem.Update();
                                            //_cContext.Load(oItem);
                                            //_cContext.ExecuteQuery(); 

                                            #endregion

                                            excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + ls + "," + oItem.Id.ToString() + "," + Tags);
                                            excelWriterTagsReport.Flush();
                                        }
                                        catch (Exception ex)
                                        {
                                            continue;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + ls + "," + "ERROR" + "," + ex.Message.Replace(",", ""));
                                excelWriterTagsReport.Flush();

                                excelWriterERROR.WriteLine(_cContext.Web.Url + "," + "List" + "," + ex.Message.Replace(",", ""));
                                excelWriterERROR.Flush();

                                continue;
                            }
                        }

                        excelWriterTagsReport.Flush();
                        excelWriterTagsReport.Close();

                        foreach (string st in _UniqueTags)
                        {
                            _strUniqueTags += st + "|";
                        }

                        excelWriterUniqueTags.WriteLine(_cContext.Web.Url + "," + _strUniqueTags);
                        excelWriterUniqueTags.Flush();
                        excelWriterUniqueTags.Close();
                    }
                }
                catch (Exception ex)
                {
                    excelWriterERROR.WriteLine(lstSiteColl[j].ToString() + "," + "Site" + "," + ex.Message.Replace(",", ""));
                    excelWriterERROR.Flush();

                    continue;
                }
            }

            #endregion                       

            excelWriterERROR.Flush();
            excelWriterERROR.Close();

            excelWriterGUID.Flush();
            excelWriterGUID.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        public string GetTermIdForTerm(string term, Guid termSetId, TermSet _termSet, TermStore _termStore, ClientContext clientContext)
        {

            //byte[] bytes = Encoding.Default.GetBytes(term);
            //term = Encoding.UTF8.GetString(bytes);


            string tid = string.Empty;
            string termId = string.Empty;
            Term _term = null;

            //TaxonomySession tSession = TaxonomySession.GetTaxonomySession(clientContext);
            //TermStore ts = tSession.GetDefaultSiteCollectionTermStore();
            //TermSet tset = ts.GetTermSet(termSetId);

            try
            {
                LabelMatchInformation lmi = new LabelMatchInformation(clientContext);

                lmi.Lcid = 1033;
                lmi.TrimUnavailable = true;
                lmi.TermLabel = term;

                TermCollection termMatches = _termSet.GetTerms(lmi);
                //clientContext.Load(tSession);
                //clientContext.Load(ts);
                //clientContext.Load(tset);
                clientContext.Load(termMatches);
                clientContext.ExecuteQuery();


                if (termMatches.Count() == 0)
                {
                    _term = _termSet.CreateTerm(term, 1033, Guid.NewGuid());
                    _termStore.CommitAll();
                    clientContext.Load(_term);
                    clientContext.ExecuteQuery();
                    tid = _term.Id.ToString();

                }

                if (termMatches != null && termMatches.Count() > 0)
                {
                    termId = termMatches.First().Id.ToString();
                    _term = termMatches.First();
                    tid = termMatches.First().Id.ToString();
                }
            }
            catch (Exception ex)
            {

            }
            return tid;

        }
        private void button22_Click(object sender, EventArgs e)
        {
            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "FileCategories" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ItemID" + "," + "Categories");
            excelWriterScoringMatrixNew.Flush();

            StreamWriter excelWriterScoringMatrixNew1 = null;
            excelWriterScoringMatrixNew1 = System.IO.File.CreateText(textBox2.Text + "\\" + "NavigationDisablePagesReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew1.WriteLine("SiteUrl" + "," + "Status");
            excelWriterScoringMatrixNew1.Flush();

            StreamWriter excelWriterScoringMatrixNew2 = null;
            excelWriterScoringMatrixNew2 = System.IO.File.CreateText(textBox2.Text + "\\" + "TopNavigationCreationReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew2.WriteLine("SiteUrl" + "," + "LibraryName" + "," + "LinkUrl");
            excelWriterScoringMatrixNew2.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {

                        #region Categorization Report Generation

                        _cContext.Load(_cContext.Web);
                        _cContext.ExecuteQuery();

                        Web _web = _cContext.Web;
                        List _List = null;

                        try
                        {
                            _List = _cContext.Web.Lists.GetByTitle("1_Uploaded Files");
                            _cContext.Load(_List);
                            _cContext.ExecuteQuery();

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                            ListItemCollection listItems = _List.GetItems(camlQuery);
                            _cContext.Load(listItems);
                            _cContext.ExecuteQuery();

                            foreach (ListItem oItem in listItems)
                            {
                                try
                                {
                                    string Categories = string.Empty;

                                    var lookupValues = new ArrayList();
                                    FieldLookupValue[] values = oItem["Categorization"] as FieldLookupValue[];

                                    foreach (FieldLookupValue value in values)
                                    {
                                        string value1 = value.LookupValue.ToString().Replace(",", "$");
                                        Categories += value1 + "|";
                                    }

                                    if (!string.IsNullOrEmpty(Categories))
                                    {
                                        excelWriterScoringMatrixNew.WriteLine(_cContext.Web.Url + "," + oItem.Id.ToString() + "," + Categories);
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }
                            }
                        }
                        catch (Exception ex)
                        { }

                        #endregion

                        #region Blog View

                        try
                        {
                            List Postslist = _cContext.Web.Lists.GetByTitle("Posts");
                            _cContext.Load(Postslist);
                            _cContext.ExecuteQuery();

                            ViewCollection ViewColl = Postslist.Views;
                            _cContext.Load(ViewColl);
                            _cContext.ExecuteQuery();

                            Microsoft.SharePoint.Client.View v = ViewColl[0];
                            _cContext.Load(v);
                            _cContext.ExecuteQuery();

                            v.ViewFields.RemoveAll();
                            v.Update();
                            _cContext.ExecuteQuery();

                            v.ViewFields.Add("Title");
                            v.ViewFields.Add("Created");
                            v.ViewFields.Add("Published");
                            v.ViewFields.Add("Category");
                            v.ViewFields.Add("NumComments");
                            v.ViewFields.Add("Edit");
                            v.ViewFields.Add("Categorization");
                            v.ViewFields.Add("LikesCount");
                            v.Update();
                            _cContext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {

                        }
                        #endregion

                        #region Site Assets

                        try
                        {
                            List olist = _web.Lists.GetByTitle("Site Assets");
                            _cContext.Load(olist);
                            _cContext.ExecuteQuery();


                            olist.EnableVersioning = false;
                            olist.Update();
                            _cContext.ExecuteQuery();

                            var items = olist.GetItems(CreateAllFilesQuery());
                            _cContext.Load(items, icol => icol.Include(i => i.File));
                            _cContext.ExecuteQuery();
                            var filecoll = items.Select(i => i.File).ToList();


                            ContentTypeCollection contentTypeColls = olist.ContentTypes;
                            // ContentTypeCollection contentTypeColl = mweb.ContentTypes;

                            _cContext.Load(contentTypeColls);
                            _cContext.ExecuteQuery();
                            ContentType defaultcontentType = null;
                            ContentType ricohcontentType = null;

                            foreach (ContentType eachcontenttype in contentTypeColls)
                            {
                                _cContext.Load(eachcontenttype);
                                _cContext.ExecuteQuery();

                                if (eachcontenttype.Name == "RicohContentType")
                                {
                                    ricohcontentType = eachcontenttype;
                                    _cContext.Load(eachcontenttype);
                                    _cContext.ExecuteQuery();
                                    _cContext.Load(ricohcontentType);
                                    ricohcontentType.ReadOnly = false;
                                    ricohcontentType.Update(false);
                                    _cContext.Load(ricohcontentType);
                                    _cContext.ExecuteQuery();

                                }
                                else if (eachcontenttype.Name == "Document")
                                {
                                    defaultcontentType = eachcontenttype;
                                    _cContext.Load(eachcontenttype);
                                    _cContext.ExecuteQuery();
                                    _cContext.Load(defaultcontentType);
                                    _cContext.ExecuteQuery();

                                }


                            }

                            bool isAttCTExist = contentTypeColls.Cast<ContentType>().Any(contentType => string.Equals(contentType.Name, "Document"));

                            if (isAttCTExist)
                            {

                                IList<ContentTypeId> reverseOrder = (from ct in contentTypeColls where ct.Name.Equals("Document", StringComparison.OrdinalIgnoreCase) select ct.Id).ToList();
                                olist.RootFolder.UniqueContentTypeOrder = reverseOrder;
                                olist.RootFolder.Update();
                                olist.Update();
                                _cContext.ExecuteQuery();
                            }


                            foreach (Microsoft.SharePoint.Client.File f in filecoll)
                            {
                                try
                                {
                                    _cContext.Load(f);
                                    _cContext.Load(f.ListItemAllFields);
                                    _cContext.ExecuteQuery();
                                    ListItem item = f.ListItemAllFields;
                                    _cContext.Load(item);
                                    _cContext.ExecuteQuery();

                                    if (item["ContentTypeId"].ToString() == ricohcontentType.Id.ToString())
                                    {
                                        DateTime Modified = Convert.ToDateTime(item["Modified"]);
                                        FieldUserValue ModifiedBy = (FieldUserValue)item["Editor"];


                                        item["ContentTypeId"] = defaultcontentType.Id;
                                        item.Update();
                                        _cContext.ExecuteQuery();


                                        item["Modified"] = Modified;
                                        item["Editor"] = ModifiedBy;

                                        item.Update();
                                        _cContext.Load(item);
                                        _cContext.ExecuteQuery();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }

                            }


                            ricohcontentType.DeleteObject();
                            _cContext.ExecuteQuery();


                            olist.EnableVersioning = true;
                            olist.Update();
                            _cContext.ExecuteQuery();



                        }
                        catch (Exception ex)
                        {

                        }
                        #endregion

                        #region List Views

                        ListCollection _Lists = _cContext.Web.Lists;
                        _cContext.Load(_Lists);
                        _cContext.ExecuteQuery();

                        foreach (List list in _Lists)
                        {
                            _cContext.Load(list);
                            _cContext.ExecuteQuery();

                            try
                            {
                                if (list.Title == "1_Uploaded Files" || list.Title == "Events" ||
                                    list.Title == "Announcements" || list.Title == "Tasks" ||
                                    list.Title == "Posts" || list.Title == "Discussions" ||
                                    list.Title == "SiteHistory" || list.Title == "Ideas" ||
                                    list.Title == "2_Documents and Pages" || list.Title == "Site Assets" ||
                                    list.Title == "Status")
                                {

                                    EnableRating(_cContext, list.Title);
                                }


                                if (list.Title == "2_Documents and Pages")
                                {

                                    ViewCollection ViewColl = list.Views;
                                    _cContext.Load(ViewColl);
                                    _cContext.ExecuteQuery();

                                    Microsoft.SharePoint.Client.View v = ViewColl[0];
                                    _cContext.Load(v);
                                    _cContext.ExecuteQuery();

                                    v.ViewFields.RemoveAll();
                                    v.Update();
                                    _cContext.ExecuteQuery();

                                    v.ViewFields.Add("DocIcon");
                                    v.ViewFields.Add("Title");
                                    v.ViewFields.Add("LinkFilename");
                                    v.ViewFields.Add("Created");
                                    v.ViewFields.Add("Created By");
                                    v.ViewFields.Add("Modified");
                                    v.ViewFields.Add("Modified By");
                                    v.ViewFields.Add("Tags");
                                    v.ViewFields.Add("Categorization");
                                    v.ViewFields.Add("CheckoutUser");
                                    v.Update();
                                    _cContext.ExecuteQuery();


                                    Folder docFolder = null;
                                    try
                                    {
                                        docFolder = list.RootFolder.Folders.GetByUrl("Documents");
                                        _cContext.Load(docFolder);
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        docFolder = list.RootFolder.Folders.Add("Documents");
                                        _cContext.Load(docFolder);
                                        _cContext.Load(docFolder, p => p.ServerRelativeUrl);
                                        _cContext.ExecuteQuery();
                                    }

                                    try
                                    {
                                        list.EnableModeration = false;
                                        list.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {

                                    }

                                    try
                                    {
                                        list.DraftVersionVisibility = DraftVisibilityType.Reader;
                                        list.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {

                                    }


                                }


                                if (list.Title == "Ideas")
                                {
                                    try
                                    {
                                        List idealist = _cContext.Web.Lists.GetByTitle("Ideas");
                                        _cContext.Load(idealist);
                                        _cContext.ExecuteQuery();

                                        ViewCollection ViewColl = idealist.Views;
                                        _cContext.Load(ViewColl);
                                        _cContext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[6];
                                        _cContext.Load(v);
                                        _cContext.ExecuteQuery();

                                        v.DeleteObject();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                                if (list.Title == "Status")
                                {
                                    try
                                    {
                                        List Statuslist = _cContext.Web.Lists.GetByTitle("Status");
                                        _cContext.Load(Statuslist);
                                        _cContext.ExecuteQuery();

                                        ViewCollection ViewColl = Statuslist.Views;
                                        _cContext.Load(ViewColl);
                                        _cContext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        _cContext.Load(v);
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.RemoveAll();
                                        v.Update();
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.Add("StatusDescription");
                                        v.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex1)
                                    {

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }

                        #endregion

                        #region Top Navigation

                        #region Disable Pages

                        Web oWeb = _cContext.Web;
                        _cContext.Load(oWeb);
                        _cContext.ExecuteQuery();

                        try
                        {

                            var pubWeb = PublishingWeb.GetPublishingWeb(_cContext, oWeb);

                            var navigation = new Helpers.ClientPortalNavigation(_cContext.Web);
                            //navigation.CurrentIncludeSubSites = false;
                            navigation.CurrentIncludePages = false;
                            navigation.GlobalIncludePages = false;
                            navigation.SaveChanges();
                            pubWeb.Web.Update();
                            oWeb.Update();
                            _cContext.ExecuteQuery();

                            excelWriterScoringMatrixNew1.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Success");
                            excelWriterScoringMatrixNew1.Flush();
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew1.WriteLine(lstSiteColl[j].ToString().Trim() + ", " + ex.Message);
                            excelWriterScoringMatrixNew1.Flush();
                            //errorLog.WriteToErrorLog(ex.Message, ex.StackTrace, ServiceSiteUrl, "Error in Disable Pages");
                        }

                        #endregion

                        #region TopNavigation Creation

                        try
                        {
                            _cContext.Load(oWeb.Navigation, we => we.UseShared);
                            _cContext.ExecuteQuery();

                            if (oWeb.Navigation.UseShared)
                                return;

                            #region DocumentLibrary Links

                            _cContext.Load(_cContext.Web.Lists);
                            _cContext.ExecuteQuery();

                            Dictionary<string, string> LinksWidUrl = new Dictionary<string, string>();

                            foreach (List list in _cContext.Web.Lists)
                            {
                                try
                                {
                                    if (list.BaseType.ToString() == "DocumentLibrary")
                                    {
                                        if (list.Title.ToString().ToLower() == "1_Uploaded Files".ToLower() || list.Title.ToString().ToLower() == "2_Documents and Pages".ToLower())
                                        {
                                            _cContext.Load(list);
                                            _cContext.ExecuteQuery();

                                            _cContext.Load(list, l => l.DefaultView);
                                            _cContext.ExecuteQuery();

                                            string Title = string.Empty;

                                            if (list.Title.ToString().ToLower() == "1_Uploaded Files".ToLower())
                                                Title = "Uploaded Files";
                                            else
                                                Title = "Documents and Pages";

                                            string url = list.DefaultView.ServerRelativeUrl.ToString();
                                            LinksWidUrl.Add(Title, url);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    //errorLog.WriteToErrorLog(ex.Message, ex.StackTrace, ServiceSiteUrl, "DocumentLibrary");
                                }
                            }

                            #endregion

                            Microsoft.SharePoint.Client.NavigationNodeCollection topnav = oWeb.Navigation.TopNavigationBar;
                            _cContext.Load(topnav);
                            _cContext.ExecuteQuery();

                            #region Creation of TopNavLinks

                            foreach (var link in LinksWidUrl)
                            {
                                NavigationNodeCreationInformation nodeInfo = new NavigationNodeCreationInformation();
                                nodeInfo.AsLastNode = true;
                                nodeInfo.Title = link.Key.ToString();
                                nodeInfo.IsExternal = true;
                                nodeInfo.Url = link.Value.ToString();

                                Microsoft.SharePoint.Client.NavigationNode node = topnav.Add(nodeInfo);
                                _cContext.Load(node);
                                _cContext.ExecuteQuery();

                                excelWriterScoringMatrixNew2.WriteLine(lstSiteColl[j].ToString().Trim() + "," + link.Key.ToString() + "," + link.Value.ToString());
                                excelWriterScoringMatrixNew2.Flush();
                            }

                            #endregion

                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew2.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "CreateTopNavLinkFailure" + ", " + ex.Message);
                            excelWriterScoringMatrixNew2.Flush();
                            //errorLog.WriteToErrorLog(ex.Message, ex.StackTrace, ServiceSiteUrl, "CreateTopNavLink");
                        }

                        #endregion

                        #endregion

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            excelWriterScoringMatrixNew1.Flush();
            excelWriterScoringMatrixNew1.Close();

            excelWriterScoringMatrixNew2.Flush();
            excelWriterScoringMatrixNew2.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button23_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "FileTags" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ItemID" + "," + "Tags");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                string SiteURL = lstSiteColl[j].ToString();//.Split('?')[0].ToString();
                                                           // string ListName = lstSiteColl[j].ToString().Split('?')[1].ToString();
                                                           // string ItmID = lstSiteColl[j].ToString().Split('?')[2].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(SiteURL.ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        _cContext.Load(_cContext.Web);
                        _cContext.ExecuteQuery();

                        Web _Web = _cContext.Web;

                        List _List = null;

                        try
                        {
                            _List = _cContext.Web.Lists.GetByTitle("1_Uploaded Files");
                            _cContext.Load(_List);
                            _cContext.ExecuteQuery();

                            _List.EnableVersioning = false;
                            _List.Update();
                            _cContext.ExecuteQuery();

                            //_List.ForceCheckout = false;
                            //_List.Update();
                            //_cContext.ExecuteQuery();

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                            ListItemCollection listItems = _List.GetItems(camlQuery);
                            _cContext.Load(listItems);
                            _cContext.ExecuteQuery();

                            ListItem oItem = listItems.GetById("2");


                            //foreach (ListItem oItem in listItems)
                            {
                                try
                                {

                                    string Tags = string.Empty;
                                    _cContext.Load(oItem);
                                    _cContext.ExecuteQuery();

                                    this.Text = oItem.Id.ToString();

                                    DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                                    FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                                    //string StartDate = "13-09-2016  14:35:00";
                                    //DateTime Modified = getdateformat(StartDate);
                                    //FieldUserValue ModifiedBy = (FieldUserValue)oItem["Author"];

                                    TaxonomyFieldValueCollection taxFieldValues = oItem["Tags"] as TaxonomyFieldValueCollection;

                                    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(_Web.Context);
                                    _cContext.Load(taxonomySession.TermStores);
                                    _cContext.ExecuteQuery();

                                    // TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_0zsJIvEgw+vz6bQNsT5BzQ==");
                                    TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");
                                    _cContext.Load(termStore);
                                    _cContext.ExecuteQuery();



                                    _cContext.Load(termStore.Groups);
                                    _cContext.ExecuteQuery();

                                    //TermGroup group = termStore.Groups.GetByName("Site Collection - verinontechnology.sharepoint.com-sites-JiveDemo18");
                                    TermGroup group = termStore.Groups.GetByName("RicohTags");

                                    _cContext.Load(group);
                                    _cContext.ExecuteQuery();


                                    _cContext.Load(group.TermSets);
                                    _cContext.ExecuteQuery();


                                    TermSet termSet = group.TermSets.GetByName("TagsTermSet");
                                    _cContext.Load(termSet);
                                    _cContext.ExecuteQuery();


                                    Field _taxnomyField = _List.Fields.GetByTitle("Tags");
                                    _cContext.Load(_taxnomyField);
                                    _cContext.ExecuteQuery();

                                    TaxonomyField txField = _cContext.CastTo<TaxonomyField>(_taxnomyField);
                                    _cContext.Load(txField);
                                    _cContext.ExecuteQuery();
                                    TaxonomyFieldValue termValue = null;

                                    TaxonomyFieldValueCollection termValues = null;

                                    string termValueString = string.Empty;

                                    string mtermId = string.Empty;
                                    termValueString = string.Empty;
                                    string termId = string.Empty;

                                    try
                                    {
                                        ////foreach (TaxonomyFieldValue tv in taxFieldValues)
                                        //{
                                        //    try
                                        //    {

                                        //        //if (tv != null)
                                        //        {

                                        //            //if (tv.Label.ToString().Contains("Ã"))
                                        //            //{

                                        //            //    byte[] bytes = Encoding.Default.GetBytes(tv.Label.ToString());
                                        //            //    Tags = Encoding.UTF8.GetString(bytes);
                                        //            //}
                                        //            //else
                                        //            //{

                                        //            //    Tags = tv.Label.ToString();
                                        //            //}


                                        //            mtermId = GetTermIdForTerm(Tags, termSet.Id, termSet, termStore, _cContext);

                                        //            if (string.IsNullOrEmpty(mtermId))
                                        //            {
                                        //                mtermId = GetTermIdForTerm(Tags, termSet.Id, termSet, termStore, _cContext);
                                        //            }

                                        //            if (!string.IsNullOrEmpty(mtermId))
                                        //                termValueString += "1033" + ";#" + Tags + "|" + mtermId + ";#";
                                        //        }
                                        //    }
                                        //    catch (Exception ex)
                                        //    {
                                        //        continue;
                                        //    }
                                        //}

                                        //if (taxFieldValues.Count > 0)
                                        {
                                            //termValueString = termValueString.Remove(termValueString.Length - 2);


                                            termValues = new TaxonomyFieldValueCollection(_cContext, termValueString, txField);

                                            txField.SetFieldValueByValueCollection(oItem, termValues);

                                            oItem.Update();
                                            _cContext.Load(oItem);
                                            _cContext.ExecuteQuery();

                                            string oName = string.Empty;
                                            string oTitle = string.Empty;

                                            try
                                            {
                                                oName = oItem["FileLeafRef"].ToString();
                                            }
                                            catch (Exception ex)
                                            { }

                                            oItem["FileLeafRef"] = oName;
                                            oItem.Update();
                                            _cContext.ExecuteQuery();

                                            _cContext.Load(oItem);
                                            _cContext.ExecuteQuery();

                                            try
                                            {
                                                oTitle = oItem["Title"].ToString();
                                            }
                                            catch (Exception ex)
                                            { }

                                            oItem["Title"] = oTitle;

                                            if (ModifiedBy.LookupValue.ToString().ToLower() == "svc jivemigration")
                                            {
                                                DateTime Created = Convert.ToDateTime(oItem["Created"]);
                                                FieldUserValue CreatedBy = (FieldUserValue)oItem["Author"];
                                                oItem["Modified"] = Created;
                                                oItem["Editor"] = CreatedBy;
                                            }
                                            else
                                            {
                                                oItem["Modified"] = Modified;
                                                oItem["Editor"] = ModifiedBy;
                                            }

                                            oItem.Update();
                                            _cContext.Load(oItem);
                                            _cContext.ExecuteQuery();
                                        }

                                        //if (!string.IsNullOrEmpty(Tags))
                                        //{
                                        //    excelWriterScoringMatrixNew.WriteLine(_cContext.Web.Url + "," + oItem.Id.ToString() + "," + Tags);
                                        //    excelWriterScoringMatrixNew.Flush();
                                        //}
                                    }
                                    catch (Exception EX)
                                    {
                                    }

                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }


                                //_List.ForceCheckout = true;
                                //_List.Update();
                                //_cContext.ExecuteQuery();


                            }

                            _List.EnableVersioning = true;
                            _List.Update();
                            _cContext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        { }

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            #endregion

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");

        }
        private void button24_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("SiteURL", typeof(string)), new DataColumn("ItemID", typeof(string)), new DataColumn("Categories", typeof(string)) });

            //Read the contents of CSV file.  
            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            //Execute a loop over the rows.  
            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;
                    //Execute a loop over the columns.  
                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "CategoryApplyReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ItemURL" + "," + "Status");

            excelWriterScoringMatrixNew.Flush();

            int count = 0;

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string _SiteTitle = drImported["SiteURL"].ToString().Trim();
                    string _ItemID = drImported["ItemID"].ToString().Trim();

                    this.Text = (count).ToString() + " : " + _SiteTitle;

                    count++;

                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_SiteTitle,
                        "svc-jivemigration7@rsharepoint.onmicrosoft.com", "Nuq92882"))
                    {
                        Web _web = _cContext.Web;
                        _cContext.Load(_web);
                        _cContext.ExecuteQuery();

                        List Pagelist = _cContext.Web.Lists.GetByTitle("1_Uploaded Files");//1_Uploaded Files//Posts
                        _cContext.Load(Pagelist);
                        _cContext.Load(Pagelist.RootFolder);
                        _cContext.ExecuteQuery();

                        Pagelist.EnableVersioning = false;
                        Pagelist.Update();
                        _cContext.ExecuteQuery();

                        #region Item Categorization Update

                        ListItem _Item = Pagelist.GetItemById(_ItemID);
                        _cContext.Load(_Item);
                        _cContext.ExecuteQuery();

                        DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                        FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                        if (!string.IsNullOrEmpty(drImported["Categories"].ToString().Trim()))
                        {
                            string _category = drImported["Categories"].ToString().Trim();
                            string[] _categories = _category.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                            try
                            {

                                FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                                for (int i = 0; i <= _categories.Length - 1; i++)
                                {
                                    string newValue = _categories[i].ToString();

                                    if (_categories[i].ToString().Contains("$"))
                                    {
                                        newValue = _categories[i].ToString().Replace("$", ",");
                                    }

                                    int _cId = GetLookupIDs(newValue, _cContext, _web);

                                    if (_cId != 0)
                                    {
                                        FieldLookupValue flv = new FieldLookupValue();
                                        flv.LookupId = _cId;

                                        lookupFieldValCollection.SetValue(flv, i);
                                    }
                                }

                                if (lookupFieldValCollection.Length >= 1)
                                {
                                    if (lookupFieldValCollection[0] != null)
                                        _Item["Categorization"] = lookupFieldValCollection;
                                }

                                _Item.Update();
                                _cContext.Load(_Item);
                                _cContext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            {
                                excelWriterScoringMatrixNew.WriteLine("ERROR:" + _SiteTitle + "," + _web.Url + "/" + Pagelist.RootFolder.Name + "/Dispform.aspx?id=" + _Item.Id + "," + drImported["Categories"].ToString().Trim());
                                excelWriterScoringMatrixNew.Flush();
                            }

                            try
                            {
                                _Item["Modified"] = Modified;
                                _Item["Editor"] = ModifiedBy;
                                _Item.Update();
                                _cContext.ExecuteQuery();

                                excelWriterScoringMatrixNew.WriteLine(_SiteTitle + "," + _web.Url + "/" + Pagelist.RootFolder.Name + "/Dispform.aspx?id=" + _Item.Id + "," + drImported["Categories"].ToString().Trim());
                                excelWriterScoringMatrixNew.Flush();
                            }
                            catch (Exception ex)
                            {
                                //excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + clientcontext.Web.Url + "," + OldTitle + "," + oTitle);
                                //excelWriterScoringMatrixNew.Flush();
                            }
                        }

                        #endregion

                        Pagelist.EnableVersioning = true;
                        Pagelist.Update();
                        _cContext.ExecuteQuery();

                    }
                }
                catch (Exception ex)
                {
                    //excelWriterScoringMatrixNew.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + drImported["SiteURL"].ToString().Trim() + "/" + "1_Uploaded Files/ Dispform.aspx?id=" + drImported["ItemID"].ToString().Trim() + "," + drImported["Categories"].ToString().Trim());
                    //excelWriterScoringMatrixNew.Flush();

                    excelWriterScoringMatrixNew.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + "ERROR: " + ex.Message.ToString());
                    excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button25_Click(object sender, EventArgs e)
        {
            List<string> lstSiteColl = new List<string>();


            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocumentsCT" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("DID" + "," + "URL" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}


            //StreamWriter excelWriterScoringMatrixNew = null;

            //excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ScoringMatrix" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            //excelWriterScoringMatrixNew.WriteLine("Filename" + "," + "URL" + "," + "Owners" + "," + "Built-in-Groups" + "," + "AD Groups" + "," + "Start Time" + "," + "End Date" + "," + "Remarks");
            //excelWriterScoringMatrixNew.Flush();

            List<string> ListNames = new List<string>();
            // ListNames.Add("Site Assets");
            ListNames.Add("2_Documents and Pages");
            // ListNames.Add("1_Uploaded Files");
            ListNames.Add("Discussions");

            // lstNameColl.Add("1_Uploaded Files");
            ListNames.Add("Events");
            // lstNameColl.Add("Announcements");
            // ListNames.Add("Tasks");
            ListNames.Add("Posts");
            // lstNameColl.Add("Discussions");
            ListNames.Add("SiteHistory");
            // lstNameColl.Add("Ideas");
            // lstNameColl.Add("Manage Categories");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();


                        List list = _Lists.GetByTitle("1_Uploaded Files");
                        clientcontext.Load(list);
                        clientcontext.ExecuteQuery();


                        list.ContentTypesEnabled = false;
                        list.Update();
                        clientcontext.ExecuteQuery();

                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + list.Title + "," + "Success");
                        excelWriterScoringMatrixNew.Flush();


                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "1_Uploaded Files" + "," + "Failure");
                    excelWriterScoringMatrixNew.Flush();
                    continue;

                }

            }
            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button26_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew2 = null;
            excelWriterScoringMatrixNew2 = System.IO.File.CreateText(textBox2.Text + "\\" + "BlogsViewReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew2.WriteLine("SiteUrl" + "," + "Status");
            excelWriterScoringMatrixNew2.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            #region VIEW for Posts

                            bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, "Posts"));

                            if (_dListExist)
                            {
                                List Pagelist = _Lists.GetByTitle("Posts");
                                clientcontext.Load(Pagelist);
                                clientcontext.ExecuteQuery();

                                ViewCollection ViewColl = Pagelist.Views;
                                clientcontext.Load(ViewColl);
                                clientcontext.ExecuteQuery();

                                Microsoft.SharePoint.Client.View v = ViewColl[0];
                                clientcontext.Load(v);
                                clientcontext.ExecuteQuery();

                                v.ViewFields.RemoveAll();
                                v.Update();
                                clientcontext.ExecuteQuery();

                                v.ViewFields.Add("LinkTitle");
                                v.ViewFields.Add("Created");
                                v.ViewFields.Add("Published");
                                v.ViewFields.Add("Category");
                                v.ViewFields.Add("NumComments");
                                v.ViewFields.Add("Edit");
                                v.ViewFields.Add("Categorization");
                                v.ViewFields.Add("LikesCount");
                                v.Update();
                                clientcontext.ExecuteQuery();

                                excelWriterScoringMatrixNew2.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Success");
                                excelWriterScoringMatrixNew2.Flush();
                            }

                            #endregion
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew2.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Failure" + ", ");
                            excelWriterScoringMatrixNew2.Flush();

                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew2.Flush();
            excelWriterScoringMatrixNew2.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button27_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew2 = null;

            excelWriterScoringMatrixNew2 = System.IO.File.CreateText(textBox2.Text + "\\" + "TopNavigationCreationReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew2.WriteLine("SiteUrl" + "," + "LibraryName" + "," + "LinkUrl");
            excelWriterScoringMatrixNew2.Flush();

            StreamWriter excelWriterScoringMatrixNew1 = null;

            excelWriterScoringMatrixNew1 = System.IO.File.CreateText(textBox2.Text + "\\" + "NavigationDisablePagesReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew1.WriteLine("SiteUrl" + "," + "Status");
            excelWriterScoringMatrixNew1.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var Ctx = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web oWeb = Ctx.Web;
                        Ctx.Load(oWeb);
                        Ctx.ExecuteQuery();

                        #region DisPages

                        //try
                        //{

                        //    var pubWeb = PublishingWeb.GetPublishingWeb(Ctx, oWeb);

                        //    var navigation = new Helpers.ClientPortalNavigation(Ctx.Web);
                        //    //navigation.CurrentIncludeSubSites = false;
                        //    navigation.CurrentIncludePages = false;
                        //    navigation.GlobalIncludePages = false;
                        //    navigation.SaveChanges();
                        //    pubWeb.Web.Update();
                        //    oWeb.Update();
                        //    Ctx.ExecuteQuery();

                        //    excelWriterScoringMatrixNew1.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Success");
                        //    excelWriterScoringMatrixNew1.Flush();
                        //}
                        //catch (Exception ex)
                        //{
                        //    excelWriterScoringMatrixNew1.WriteLine(lstSiteColl[j].ToString().Trim() + ", " + ex.Message);
                        //    excelWriterScoringMatrixNew1.Flush();
                        //    //errorLog.WriteToErrorLog(ex.Message, ex.StackTrace, ServiceSiteUrl, "Error in Disable Pages");
                        //}

                        #endregion

                        #region Delete Navigation if Already

                        Ctx.RequestTimeout = -1;

                        Ctx.Load(oWeb.Navigation, we => we.UseShared);
                        Ctx.ExecuteQuery();

                        if (oWeb.Navigation.UseShared)
                            return;

                        int deletedCount = 0;
                        TNav: Microsoft.SharePoint.Client.NavigationNodeCollection topnav1 = oWeb.Navigation.TopNavigationBar;
                        Ctx.Load(topnav1);
                        Ctx.ExecuteQuery();

                        foreach (var nColl in topnav1)
                        {
                            Ctx.Load(nColl);
                            Ctx.ExecuteQuery();

                            //if (nColl.Title.ToLower() == "Overview".ToLower())
                            if (nColl.Title.ToLower() == "Uploaded Files".ToLower() || nColl.Title.ToLower() == "Documents and Pages".ToLower())
                            {
                                nColl.DeleteObject();
                                //nColl.Update();
                                Ctx.ExecuteQuery();
                                deletedCount++;

                                //if (deletedCount == 2)
                                //    break;
                                //else
                                goto TNav;
                            }
                        }

                        #endregion

                        #region CreTopNav

                        try
                        {
                            Ctx.Load(oWeb.Navigation, we => we.UseShared);
                            Ctx.ExecuteQuery();

                            if (oWeb.Navigation.UseShared)
                                return;

                            #region DocumentLibrary

                            Ctx.Load(Ctx.Web.Lists);
                            Ctx.ExecuteQuery();

                            Dictionary<string, string> LinksWidUrl = new Dictionary<string, string>();

                            foreach (List list in Ctx.Web.Lists)
                            {
                                try
                                {
                                    if (list.BaseType.ToString() == "DocumentLibrary")
                                    {
                                        if (list.Title.ToString().ToLower() == "1_Uploaded Files".ToLower() || list.Title.ToString().ToLower() == "2_Documents and Pages".ToLower())
                                        {
                                            Ctx.Load(list);
                                            Ctx.ExecuteQuery();

                                            Ctx.Load(list, l => l.DefaultView);
                                            Ctx.ExecuteQuery();

                                            string Title = string.Empty;

                                            if (list.Title.ToString().ToLower() == "1_Uploaded Files".ToLower())
                                                Title = "Uploaded Files";
                                            else
                                                Title = "Documents and Pages";

                                            string url = list.DefaultView.ServerRelativeUrl.ToString();
                                            LinksWidUrl.Add(Title, url);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    //errorLog.WriteToErrorLog(ex.Message, ex.StackTrace, ServiceSiteUrl, "DocumentLibrary");
                                }
                            }

                            #endregion

                            Microsoft.SharePoint.Client.NavigationNodeCollection topnav = oWeb.Navigation.TopNavigationBar;
                            Ctx.Load(topnav);
                            Ctx.ExecuteQuery();

                            #region
                            foreach (var link in LinksWidUrl)
                            {
                                NavigationNodeCreationInformation nodeInfo = new NavigationNodeCreationInformation();
                                nodeInfo.AsLastNode = true;
                                nodeInfo.Title = link.Key.ToString();
                                nodeInfo.IsExternal = true;
                                nodeInfo.Url = link.Value.ToString();

                                Microsoft.SharePoint.Client.NavigationNode node = topnav.Add(nodeInfo);
                                Ctx.Load(node);
                                Ctx.ExecuteQuery();

                                excelWriterScoringMatrixNew2.WriteLine(lstSiteColl[j].ToString().Trim() + "," + link.Key.ToString() + "," + link.Value.ToString());
                                excelWriterScoringMatrixNew2.Flush();
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew2.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "CreateTopNavLinkFailure" + ", " + ex.Message);
                            excelWriterScoringMatrixNew2.Flush();
                            //errorLog.WriteToErrorLog(ex.Message, ex.StackTrace, ServiceSiteUrl, "CreateTopNavLink");
                        }

                        #endregion                            
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }

            }

            excelWriterScoringMatrixNew2.Flush();
            excelWriterScoringMatrixNew2.Close();

            excelWriterScoringMatrixNew1.Flush();
            excelWriterScoringMatrixNew1.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button28_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remaining

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            #region VIEW for "Announcements" List

                            bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, "Announcements"));

                            if (_dListExist)
                            {
                                List Pagelist = _Lists.GetByTitle("Announcements");
                                clientcontext.Load(Pagelist);
                                clientcontext.ExecuteQuery();

                                ViewCollection ViewColl = Pagelist.Views;
                                clientcontext.Load(ViewColl);
                                clientcontext.ExecuteQuery();

                                Microsoft.SharePoint.Client.View v = ViewColl[0];
                                clientcontext.Load(v);
                                clientcontext.ExecuteQuery();

                                v.ViewFields.RemoveAll();
                                v.Update();
                                clientcontext.ExecuteQuery();

                                v.ViewFields.Add("Title");
                                v.ViewFields.Add("Modified");
                                v.ViewFields.Add("LikesCount");
                                v.Update();
                                clientcontext.ExecuteQuery();
                            }

                            #endregion
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");

            #endregion
        }
        private void button29_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "Activity&Backendlist" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteUrl" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            #region Remaining

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        clientcontext.Web.SetHomePage("Pages/Activity.aspx");
                        clientcontext.ExecuteQuery();

                        //ListCollection _Lists = clientcontext.Web.Lists;
                        //clientcontext.Load(_Lists);
                        //clientcontext.ExecuteQuery();

                        //try
                        //{
                        //    #region Delete Activity Page

                        //    List _List = null;

                        //    try
                        //    {
                        //        _List = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages"); ;
                        //        clientcontext.Load(_List);
                        //        clientcontext.ExecuteQuery();
                        //    }
                        //    catch (Exception ex)
                        //    { }

                        //    if (_List != null)
                        //    {
                        //        try
                        //        {
                        //            clientcontext.Load(_List.RootFolder);
                        //            clientcontext.ExecuteQuery();

                        //            try
                        //            {
                        //                ListItem _Item = _List.RootFolder.Files.GetByUrl("Activity.aspx").ListItemAllFields;
                        //                clientcontext.Load(_Item);
                        //                clientcontext.ExecuteQuery();

                        //                _Item.DeleteObject();
                        //                _List.Update();
                        //                clientcontext.ExecuteQuery();

                        //                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Success");
                        //                excelWriterScoringMatrixNew.Flush();
                        //            }
                        //            catch (Exception ex)
                        //            {
                        //                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Failure");
                        //                excelWriterScoringMatrixNew.Flush();
                        //            }                                    
                        //        }
                        //        catch (Exception ex)
                        //        {

                        //        }
                        //    }

                        //    #endregion

                        //    #region Delete "Key Content and Places" List

                        //    //bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, "Key Content and Places"));

                        //    //if (_dListExist)
                        //    //{
                        //    //    try
                        //    //    {
                        //    //        List Pagelist = _Lists.GetByTitle("Key Content and Places");
                        //    //        clientcontext.Load(Pagelist);
                        //    //        clientcontext.ExecuteQuery();

                        //    //        Pagelist.DeleteObject();
                        //    //        clientcontext.ExecuteQuery();

                        //    //        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Key Content and Places");
                        //    //        excelWriterScoringMatrixNew.Flush();
                        //    //    }
                        //    //    catch(Exception ex)
                        //    //    {
                        //    //        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "KeyFailure");
                        //    //        excelWriterScoringMatrixNew.Flush();
                        //    //    }
                        //    //}

                        //    #endregion

                        //}
                        //catch (Exception ex)
                        //{
                        //    continue;
                        //}
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");

            #endregion
        }
        private void button30_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "CrawlHideShowRibbonReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("URL" + "," + "CrawlStatus" + "," + "HideShowRibbon" + "," + "MajorVersion");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                string Crawl = "NA";
                string HideShowRibbon = "Failure";
                string MajorVersion = "Failure";

                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration4@rsharepoint.onmicrosoft.com", "Pac30064"))
                    {
                        Web web = clientcontext.Web;
                        clientcontext.Load(web);
                        clientcontext.ExecuteQuery();
                        clientcontext.RequestTimeout = -1;

                        try
                        {
                            web.NoCrawl = false;
                            web.Update();
                            clientcontext.Load(web);
                            clientcontext.ExecuteQuery();

                            Crawl = "Success";
                        }
                        catch (Exception ex2)
                        {
                            Crawl = ex2.Message;
                        }

                        #region HideShowRibbon

                        try
                        {
                            var pubWeb = PublishingWeb.GetPublishingWeb(clientcontext, web);
                            pubWeb.Web.AllProperties["__DisplayShowHideRibbonActionId"] = false.ToString();
                            pubWeb.Web.Update();
                            web.Update();
                            clientcontext.ExecuteQuery();

                            HideShowRibbon = "Success";
                        }
                        catch (Exception ex)
                        {
                            HideShowRibbon = ex.Message;
                        }

                        try
                        {
                            List _List = clientcontext.Web.Lists.GetByTitle("2_Documents and Pages");
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();

                            _List.UpdateListVersioning(true, false, true);
                            clientcontext.ExecuteQuery();

                            MajorVersion = "Success";
                        }
                        catch (Exception ex)
                        { }

                        #endregion

                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + Crawl + "," + HideShowRibbon + "," + MajorVersion);
                        excelWriterScoringMatrixNew.Flush();

                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "ERROR" + "," + ex.Message + "," + "");
                    excelWriterScoringMatrixNew.Flush();
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button31_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DuplicateCategoriesReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteUrl" + "," + "CategoryTitle" + "," + "Count");
            excelWriterScoringMatrixNew.Flush();

            #region Remaining

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            #region Delete Activity Page

                            List _List = null;

                            try
                            {
                                _List = clientcontext.Web.Lists.GetByTitle("Manage Categories");
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                CamlQuery camlQuery = new CamlQuery();
                                camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";

                                ListItemCollection listItems = _List.GetItems(camlQuery);
                                clientcontext.Load(listItems);
                                clientcontext.ExecuteQuery();

                                Dictionary<string, int> DuplicateCategories = new Dictionary<string, int>();

                                foreach (ListItem _Item in listItems)
                                {
                                    try
                                    {
                                        string oTitle = _Item["Title"].ToString();

                                        if (!DuplicateCategories.Keys.Contains(oTitle))
                                        {
                                            DuplicateCategories.Add(oTitle, 1);
                                        }
                                        else
                                        {
                                            int c = DuplicateCategories[oTitle] + 1;
                                            DuplicateCategories[oTitle] = c;
                                        }

                                        //try
                                        //{
                                        //    DuplicateCategories.Add(oTitle, 1);
                                        //}
                                        //catch(Exception ex)
                                        //{
                                        //    int c = DuplicateCategories[oTitle] + 1;
                                        //    DuplicateCategories[oTitle]= c;
                                        //}

                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }

                                foreach (KeyValuePair<string, int> kvp in DuplicateCategories)
                                {
                                    if (kvp.Value > 1)
                                    {
                                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString().Trim() + "," + kvp.Key.ToString().Trim() + "," + kvp.Value.ToString().Trim());
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                }
                            }

                            #endregion
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");

            #endregion
        }
        private void button32_Click(object sender, EventArgs e)
        {
            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("ObjectURL", typeof(string)), new DataColumn("ObjectType", typeof(string)), new DataColumn("ObjectID", typeof(string)), new DataColumn("Tags", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            //DataTable dtJiveObjects = new DataTable();
            //dtJiveObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("PlaceID", typeof(string)), new DataColumn("ObjectId", typeof(string)), new DataColumn("ObjectType", typeof(string)), new DataColumn("TagsSet", typeof(string)) });

            //string csvData1 = System.IO.File.ReadAllText(textBox3.Text);

            //foreach (string row in csvData1.Split('\n'))
            //{
            //    if (!string.IsNullOrEmpty(row))
            //    {
            //        dtJiveObjects.Rows.Add();
            //        int i = 0;

            //        foreach (string cell in row.Split(','))
            //        {
            //            dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
            //            i++;
            //        }
            //    }
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "TagsImportReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("ObjectType" + "," + "ListName" + "," + "TagsSet" + "," + "Status" + "," + "SPURL");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;
            string[] SiteSplit = new string[] { "/Lists/" };
            string[] IDSplit = new string[] { "?ID=" };
            string[] DocumentSplit = new string[] { "/Documents/" };
            string[] TagsSplit = new string[] { "|" };
            string[] FileURLSplit = new string[] { "/1_Uploaded Files/" };
            string[] PageURLSplit = new string[] { "/Pages/" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string TagsSet = string.Empty;
                    string _SiteURL = string.Empty;
                    string[] TagsColl = null;

                    string _objList = drImported["ObjectType"].ToString().Trim();
                    string _objId = drImported["ObjectID"].ToString().Trim();
                    string _objectURL = drImported["ObjectURL"].ToString().Trim();
                    string _importedURL = drImported["ObjectURL"].ToString().Trim();
                    TagsSet = drImported["Tags"].ToString().Trim();
                    TagsColl = TagsSet.Split(TagsSplit, StringSplitOptions.RemoveEmptyEntries);

                    if (_importedURL.Contains("/Lists/"))
                    {
                        _importedURL = drImported["ObjectURL"].ToString().Split(SiteSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }
                    if (_importedURL.Contains("/1_Uploaded Files/"))
                    {
                        _importedURL = drImported["ObjectURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }
                    if (_importedURL.Contains("/Pages/"))
                    {
                        _importedURL = drImported["ObjectURL"].ToString().Split(PageURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }

                    this.Text = (count).ToString() + " : " + _objId;
                    count++;

                    if (!string.IsNullOrEmpty(TagsSet))
                    {
                        AuthenticationManager authManager = new AuthenticationManager();

                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, textBox6.Text, textBox5.Text))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();


                            List _List = null;
                            string listName = string.Empty;
                            string _FilePath = string.Empty;
                            string _itemID = string.Empty;

                            switch (_objList)
                            {
                                case "Document":
                                    listName = "2_Documents and Pages";
                                    _FilePath = drImported["ObjectURL"].ToString().Split(DocumentSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    break;

                                case "File":
                                    listName = "1_Uploaded Files";
                                    _itemID = drImported["ObjectURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    break;

                                //case "Announcement":
                                //    listName = "Announcements";
                                //    _itemID = drImported["ObjectURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                //    break;

                                case "Blog":
                                    listName = "Posts";
                                    _itemID = drImported["ObjectURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    break;

                                case "Discussion":
                                    listName = "Discussions";
                                    _itemID = drImported["ObjectURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    break;

                                case "Event":
                                    listName = "Events";
                                    _itemID = drImported["ObjectURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    break;

                                case "Task":
                                    listName = "Tasks";
                                    _itemID = drImported["ObjectURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    break;

                                case "Idea":
                                    listName = "Ideas";
                                    _itemID = drImported["ObjectURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    break;

                                    //case "Poll":

                                    //    break;
                            }

                            try
                            {
                                _List = clientcontext.Web.Lists.GetByTitle(listName);
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                bool tagsFileldExist = _List.FieldExistsByName("Tag");

                                if (tagsFileldExist)
                                {
                                    if (_List.Title == "2_Documents and Pages")
                                    {
                                        _List.EnableVersioning = false;
                                        _List.Update();
                                        clientcontext.ExecuteQuery();

                                        _List.ForceCheckout = false;
                                        _List.Update();
                                        clientcontext.ExecuteQuery();

                                        try
                                        {
                                            clientcontext.Load(_List.RootFolder);
                                            clientcontext.ExecuteQuery();

                                            Folder docFolder = null;

                                            try
                                            {
                                                docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                                clientcontext.Load(docFolder);
                                                clientcontext.ExecuteQuery();
                                            }
                                            catch (Exception ex)
                                            {
                                            }

                                            if (docFolder != null)
                                            {
                                                ListItem oItem = docFolder.Files.GetByUrl(_FilePath).ListItemAllFields;
                                                //ListItem oItem = targetList.GetItemById(_ItemID);
                                                clientcontext.Load(oItem);
                                                clientcontext.ExecuteQuery();

                                                DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                                                FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                                                try
                                                {
                                                    FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[TagsColl.Length];

                                                    for (int i = 0; i <= TagsColl.Length - 1; i++)
                                                    {
                                                        string newValue = TagsColl[i].ToString();

                                                        if (TagsColl[i].ToString().Contains("$"))
                                                        {
                                                            newValue = TagsColl[i].ToString().Replace("$", ",");
                                                        }

                                                        int _cId = GetLookupIDsManageTag(newValue, clientcontext, oWeb);

                                                        if (_cId != 0)
                                                        {
                                                            FieldLookupValue flv = new FieldLookupValue();
                                                            flv.LookupId = _cId;

                                                            lookupFieldValCollection.SetValue(flv, i);
                                                        }
                                                    }

                                                    if (lookupFieldValCollection.Length >= 1)
                                                    {
                                                        if (lookupFieldValCollection[0] != null)
                                                            oItem["Tag"] = lookupFieldValCollection;
                                                    }

                                                    oItem.Update();
                                                    clientcontext.Load(oItem);
                                                    clientcontext.ExecuteQuery();

                                                    excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Success" + "," + _objectURL);
                                                    excelWriterScoringMatrixNew.Flush();
                                                    //}
                                                }
                                                catch (Exception EX)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Error : " + EX.Message + "," + _objectURL);
                                                    excelWriterScoringMatrixNew.Flush();
                                                }

                                                try
                                                {
                                                    oItem["Modified"] = Modified;
                                                    oItem["Editor"] = ModifiedBy;
                                                    oItem.Update();
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                        }

                                        _List.EnableVersioning = true;
                                        _List.Update();
                                        clientcontext.ExecuteQuery();

                                        _List.ForceCheckout = true;
                                        _List.Update();
                                        clientcontext.ExecuteQuery();

                                        #region TAGS APPLY

                                        //try
                                        //{
                                        //    clientcontext.Load(_List.RootFolder);
                                        //    clientcontext.ExecuteQuery();

                                        //    Folder docFolder = null;

                                        //    try
                                        //    {
                                        //        docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                        //        clientcontext.Load(docFolder);
                                        //        clientcontext.ExecuteQuery();
                                        //    }
                                        //    catch (Exception ex)
                                        //    { }

                                        //    if (docFolder != null)
                                        //    {
                                        //        ListItem _Item = docFolder.Files.GetByUrl(_FilePath).ListItemAllFields;
                                        //        clientcontext.Load(_Item);
                                        //        clientcontext.ExecuteQuery();

                                        //        DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                        //        FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];
                                        //        TaxonomyFieldValueCollection taxFieldValues = _Item["Tags"] as TaxonomyFieldValueCollection;

                                        //        if (taxFieldValues.Count < 1)
                                        //        {
                                        //            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(oWeb.Context);
                                        //            clientcontext.Load(taxonomySession.TermStores);
                                        //            clientcontext.ExecuteQuery();

                                        //            TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");
                                        //            clientcontext.Load(termStore);
                                        //            clientcontext.ExecuteQuery();

                                        //            clientcontext.Load(termStore.Groups);
                                        //            clientcontext.ExecuteQuery();

                                        //            TermGroup group = termStore.Groups.GetByName("RicohTags");
                                        //            clientcontext.Load(group);
                                        //            clientcontext.ExecuteQuery();

                                        //            clientcontext.Load(group.TermSets);
                                        //            clientcontext.ExecuteQuery();

                                        //            TermSet termSet = group.TermSets.GetByName("TagsTermSet");
                                        //            clientcontext.Load(termSet);
                                        //            clientcontext.ExecuteQuery();

                                        //            Field _taxnomyField = _List.Fields.GetByTitle("Tags");
                                        //            clientcontext.Load(_taxnomyField);
                                        //            clientcontext.ExecuteQuery();

                                        //            TaxonomyField txField = clientcontext.CastTo<TaxonomyField>(_taxnomyField);
                                        //            clientcontext.Load(txField);
                                        //            clientcontext.ExecuteQuery();

                                        //            TaxonomyFieldValueCollection termValues = null;

                                        //            string termValueString = string.Empty;
                                        //            string termId = string.Empty;

                                        //            try
                                        //            {
                                        //                foreach (string tv in TagsColl)
                                        //                {
                                        //                    string mtermId = string.Empty;

                                        //                    try
                                        //                    {
                                        //                        //mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);

                                        //                        if (string.IsNullOrEmpty(mtermId))
                                        //                        {
                                        //                            mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);
                                        //                        }

                                        //                        if (!string.IsNullOrEmpty(mtermId))
                                        //                            termValueString += "1033" + ";#" + tv + "|" + mtermId + ";#";

                                        //                    }
                                        //                    catch (Exception ex)
                                        //                    {
                                        //                        continue;
                                        //                    }
                                        //                }

                                        //                //if (taxFieldValues.Count > 0)
                                        //                //{
                                        //                termValueString = termValueString.Remove(termValueString.Length - 2);
                                        //                termValues = new TaxonomyFieldValueCollection(clientcontext, termValueString,
                                        //                    txField);

                                        //                txField.SetFieldValueByValueCollection(_Item, termValues);

                                        //                _Item.Update();
                                        //                clientcontext.Load(_Item);
                                        //                clientcontext.ExecuteQuery();

                                        //                _Item["Modified"] = Modified;
                                        //                _Item["Editor"] = ModifiedBy;

                                        //                _Item.Update();
                                        //                clientcontext.Load(_Item);
                                        //                clientcontext.ExecuteQuery();

                                        //                excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Success" + "," + _objectURL);
                                        //                excelWriterScoringMatrixNew.Flush();
                                        //                //}
                                        //            }
                                        //            catch (Exception EX)
                                        //            {
                                        //                excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Error : " + EX.Message + "," + _objectURL);
                                        //                excelWriterScoringMatrixNew.Flush();
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //catch (Exception ex)
                                        //{
                                        //    //excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Failure due to : " + ex.Message);
                                        //    //excelWriterScoringMatrixNew.Flush();
                                        //} 

                                        #endregion

                                    }
                                    else
                                    {
                                        _List.EnableVersioning = false;
                                        _List.Update();
                                        clientcontext.ExecuteQuery();

                                        try
                                        {
                                            ListItem oItem = _List.GetItemById(_itemID);
                                            clientcontext.Load(oItem);
                                            clientcontext.ExecuteQuery();

                                            DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                                            FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                                            try
                                            {
                                                FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[TagsColl.Length];

                                                for (int i = 0; i <= TagsColl.Length - 1; i++)
                                                {
                                                    string newValue = TagsColl[i].ToString();

                                                    if (TagsColl[i].ToString().Contains("$"))
                                                    {
                                                        newValue = TagsColl[i].ToString().Replace("$", ",");
                                                    }

                                                    int _cId = GetLookupIDsManageTag(newValue, clientcontext, oWeb);

                                                    if (_cId != 0)
                                                    {
                                                        FieldLookupValue flv = new FieldLookupValue();
                                                        flv.LookupId = _cId;

                                                        lookupFieldValCollection.SetValue(flv, i);
                                                    }
                                                }

                                                if (lookupFieldValCollection.Length >= 1)
                                                {
                                                    if (lookupFieldValCollection[0] != null)
                                                        oItem["Tag"] = lookupFieldValCollection;
                                                }

                                                oItem.Update();
                                                clientcontext.Load(oItem);
                                                clientcontext.ExecuteQuery();

                                                excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Success" + "," + _objectURL);
                                                excelWriterScoringMatrixNew.Flush();
                                                //}
                                            }
                                            catch (Exception EX)
                                            {
                                                excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Error : " + EX.Message + "," + _objectURL);
                                                excelWriterScoringMatrixNew.Flush();
                                            }

                                            try
                                            {
                                                oItem["Modified"] = Modified;
                                                oItem["Editor"] = ModifiedBy;
                                                oItem.Update();
                                                clientcontext.ExecuteQuery();
                                            }
                                            catch (Exception ex)
                                            {
                                            }
                                        }
                                        catch (Exception ed)
                                        {

                                        }

                                        _List.EnableVersioning = false;
                                        _List.Update();
                                        clientcontext.ExecuteQuery();

                                        #region TAGS APPLY

                                        //try
                                        //{
                                        //    ListItem _Item = _List.GetItemById(_itemID);
                                        //    clientcontext.Load(_Item);
                                        //    clientcontext.ExecuteQuery();

                                        //    DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                        //    FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];
                                        //    TaxonomyFieldValueCollection taxFieldValues = _Item["Tags"] as TaxonomyFieldValueCollection;

                                        //    if (taxFieldValues.Count < 1)
                                        //    {
                                        //        TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(oWeb.Context);
                                        //        clientcontext.Load(taxonomySession.TermStores);
                                        //        clientcontext.ExecuteQuery();

                                        //        TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");
                                        //        clientcontext.Load(termStore);
                                        //        clientcontext.ExecuteQuery();

                                        //        clientcontext.Load(termStore.Groups);
                                        //        clientcontext.ExecuteQuery();

                                        //        TermGroup group = termStore.Groups.GetByName("RicohTags");
                                        //        clientcontext.Load(group);
                                        //        clientcontext.ExecuteQuery();

                                        //        clientcontext.Load(group.TermSets);
                                        //        clientcontext.ExecuteQuery();

                                        //        TermSet termSet = group.TermSets.GetByName("TagsTermSet");
                                        //        clientcontext.Load(termSet);
                                        //        clientcontext.ExecuteQuery();


                                        //        Field _taxnomyField = _List.Fields.GetByTitle("Tags");
                                        //        clientcontext.Load(_taxnomyField);
                                        //        clientcontext.ExecuteQuery();

                                        //        TaxonomyField txField = clientcontext.CastTo<TaxonomyField>(_taxnomyField);
                                        //        clientcontext.Load(txField);
                                        //        clientcontext.ExecuteQuery();

                                        //        TaxonomyFieldValueCollection termValues = null;

                                        //        string termValueString = string.Empty;
                                        //        string termId = string.Empty;

                                        //        try
                                        //        {
                                        //            foreach (string tv in TagsColl)
                                        //            {
                                        //                string mtermId = string.Empty;

                                        //                try
                                        //                {
                                        //                    //mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);

                                        //                    if (string.IsNullOrEmpty(mtermId))
                                        //                    {
                                        //                        mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);
                                        //                    }

                                        //                    if (!string.IsNullOrEmpty(mtermId))
                                        //                        termValueString += "1033" + ";#" + tv + "|" + mtermId + ";#";

                                        //                }
                                        //                catch (Exception ex)
                                        //                {
                                        //                    continue;
                                        //                }
                                        //            }

                                        //            //if (taxFieldValues.Count > 0)
                                        //            //{
                                        //            termValueString = termValueString.Remove(termValueString.Length - 2);
                                        //            termValues = new TaxonomyFieldValueCollection(clientcontext, termValueString,
                                        //                txField);

                                        //            txField.SetFieldValueByValueCollection(_Item, termValues);

                                        //            _Item.Update();
                                        //            clientcontext.Load(_Item);
                                        //            clientcontext.ExecuteQuery();

                                        //            _Item["Modified"] = Modified;
                                        //            _Item["Editor"] = ModifiedBy;

                                        //            _Item.Update();
                                        //            clientcontext.Load(_Item);
                                        //            clientcontext.ExecuteQuery();

                                        //            excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Success" + "," + _objectURL);
                                        //            excelWriterScoringMatrixNew.Flush();
                                        //            //}
                                        //        }
                                        //        catch (Exception EX)
                                        //        {
                                        //            excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Error : " + EX.Message + "," + _objectURL);
                                        //            excelWriterScoringMatrixNew.Flush();
                                        //        }
                                        //    }
                                        //}
                                        //catch (Exception EX)
                                        //{
                                        //    excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Error : " + EX.Message + "," + _objectURL);
                                        //    excelWriterScoringMatrixNew.Flush();
                                        //}

                                        //_List.EnableVersioning = true;
                                        //_List.Update();
                                        //clientcontext.ExecuteQuery(); 

                                        #endregion
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        excelWriterScoringMatrixNew.WriteLine(_objList + "," + _objectURL + "," + "NoTags" + "," + "TagSetEmpty" + "," + _objectURL);
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button33_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remaining

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            #region VIEW for "Status" List

                            bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, "Status"));

                            if (_dListExist)
                            {
                                try
                                {
                                    List oList = _Lists.GetByTitle("Status");
                                    clientcontext.Load(oList);
                                    clientcontext.ExecuteQuery();


                                    ListItemCollection listItems = oList.GetItems(CamlQuery.CreateAllItemsQuery());
                                    clientcontext.Load(listItems,
                                                        eachItem => eachItem.Include(
                                                        item => item,
                                                        item => item["ID"]));
                                    clientcontext.ExecuteQuery();

                                    var totalListItems = listItems.Count;

                                    if (totalListItems > 0)
                                    {
                                        for (var counter = totalListItems - 1; counter > -1; counter--)
                                        {
                                            listItems[counter].DeleteObject();
                                            clientcontext.ExecuteQuery();
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {

                                }

                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");

            #endregion
        }
        private void button34_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteHistoryGroupReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("ObjectURL");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    //using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant("https://rsharepoint.sharepoint.com/sites/rworldgroups2/federal-output-manager-implementation", "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            List oList = oLists.GetByTitle("SiteHistory");
                            clientcontext.Load(oList);
                            clientcontext.ExecuteQuery();

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = "<View><RowLimit>10</RowLimit></View>";

                            ListItemCollection listItems = oList.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem _Item in listItems)
                            {
                                clientcontext.Load(_Item);
                                clientcontext.ExecuteQuery();

                                if (_Item["Group_Type"].ToString() == "MEMBER_ONLY")
                                {
                                    DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                    FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];
                                    //DateTime Modified = getdateformat("7/25/2018 10:24:50 PM");

                                    _Item["Group_Type"] = "Members only";
                                    _Item["Modified"] = Modified;
                                    _Item["Editor"] = ModifiedBy;
                                    _Item.Update();
                                    clientcontext.ExecuteQuery();

                                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString());
                                    excelWriterScoringMatrixNew.Flush();
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button35_Click(object sender, EventArgs e)
        {
            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("SpaceID", typeof(string)), new DataColumn("ObjectId", typeof(string)), new DataColumn("ObjectType", typeof(string)), new DataColumn("ImportedURL", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            DataTable dtJiveObjects = new DataTable();
            dtJiveObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("ID", typeof(string)), new DataColumn("startDate", typeof(string)), new DataColumn("endDate", typeof(string)) });
            ///***********

            string csvData1 = System.IO.File.ReadAllText(textBox3.Text);

            foreach (string row in csvData1.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtJiveObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "EventSTENDDatesReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("ObjectID" + "," + "SPURL" + "," + "StartDate" + "," + "EndDate" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;
            string[] SiteSplit = new string[] { "/Lists/" };
            string[] IDSplit = new string[] { "?ID=" };

            foreach (DataRow drJive in dtJiveObjects.Rows)
            {
                try
                {
                    string _JiveObjectID = drJive["ID"].ToString().Trim();
                    string EventDate = drJive["startDate"].ToString().Trim(); ///***********
                    string EndDate = drJive["endDate"].ToString().Trim();///***********

                    string _objectID = string.Empty;
                    string _itemID = string.Empty;
                    string _objectURL = string.Empty;
                    string _importedURL = string.Empty;

                    this.Text = (count).ToString() + " : " + _JiveObjectID;

                    count++;
                    bool itemFound = false;

                    foreach (DataRow drImported in dtImportedObjects.Rows)
                    {
                        if (drImported["ObjectId"].ToString().Trim() == drJive["ID"].ToString().Trim())
                        {
                            _objectID = drImported["ImportedURL"].ToString().Trim();
                            _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                            _objectURL = drImported["ImportedURL"].ToString().Trim();
                            _importedURL = drImported["ImportedURL"].ToString().Trim();

                            if (_importedURL.Contains("/Lists/"))
                            {
                                _importedURL = drImported["ImportedURL"].ToString().Split(SiteSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                            }

                            itemFound = true;
                            break;
                        }
                    }

                    if (itemFound)
                    {
                        AuthenticationManager authManager = new AuthenticationManager();

                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();

                            List _List = null;

                            try
                            {
                                _List = clientcontext.Web.Lists.GetByTitle("Events");
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                _List.EnableVersioning = false;
                                _List.Update();
                                clientcontext.ExecuteQuery();

                                try
                                {
                                    ListItem _Item = _List.GetItemById(_itemID);
                                    clientcontext.Load(_Item);
                                    clientcontext.ExecuteQuery();

                                    DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                    FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                    DateTime SDate = getdateformat(EventDate);
                                    DateTime EDate = getdateformat(EndDate);

                                    _Item["EventDate"] = SDate;
                                    _Item["EndDate"] = EDate;
                                    _Item["Modified"] = Modified;
                                    _Item["Editor"] = ModifiedBy;
                                    _Item.Update();
                                    clientcontext.ExecuteQuery();

                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + EventDate + "," + EndDate + "," + "Success");
                                    excelWriterScoringMatrixNew.Flush();
                                }
                                catch (Exception EX)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + EventDate + "," + EndDate + "," + "Failure");
                                    excelWriterScoringMatrixNew.Flush();
                                }

                                _List.EnableVersioning = true;
                                _List.Update();
                                clientcontext.ExecuteQuery();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button36_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("SpaceID", typeof(string)), new DataColumn("ObjectId", typeof(string)), new DataColumn("ObjectType", typeof(string)), new DataColumn("ImportedURL", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            DataTable dtJiveObjects = new DataTable();
            dtJiveObjects.Columns.AddRange(new DataColumn[2] { new DataColumn("ID", typeof(string)), new DataColumn("Categories", typeof(string)) });
            ///***********

            string csvData1 = System.IO.File.ReadAllText(textBox3.Text);

            foreach (string row in csvData1.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtJiveObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "CategoriesFixReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("ObjectID" + "," + "SPURL" + "," + "Categories" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;
            string[] FileURLSplit = new string[] { "/1_Uploaded Files/" };
            string[] IDSplit = new string[] { "?ID=" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string Categories = string.Empty;

                    string _objectID = string.Empty;
                    string _itemID = string.Empty;
                    string _objectURL = string.Empty;
                    string _importedURL = string.Empty;

                    _objectID = drImported["ObjectId"].ToString().Trim();
                    _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                    _objectURL = drImported["ImportedURL"].ToString().Trim();
                    _importedURL = drImported["ImportedURL"].ToString().Trim();

                    if (_importedURL.Contains("/1_Uploaded Files/"))
                    {
                        _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }

                    this.Text = (count).ToString() + " : " + _objectURL;

                    count++;
                    bool itemFound = false;

                    foreach (DataRow drJive in dtJiveObjects.Rows)
                    {
                        if (drImported["ObjectId"].ToString().Trim() == drJive["ID"].ToString().Trim())
                        {
                            Categories = drJive["Categories"].ToString().Trim();
                            itemFound = true;
                            break;
                        }
                    }

                    if (itemFound && !string.IsNullOrEmpty(Categories))
                    {
                        AuthenticationManager authManager = new AuthenticationManager();

                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();

                            List _List = null;

                            try
                            {
                                _List = clientcontext.Web.Lists.GetByTitle("1_Uploaded Files");
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                _List.EnableVersioning = false;
                                _List.Update();
                                clientcontext.ExecuteQuery();

                                try
                                {
                                    ListItem _Item = _List.GetItemById(_itemID);
                                    clientcontext.Load(_Item);
                                    clientcontext.ExecuteQuery();

                                    DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                    FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                    if (!string.IsNullOrEmpty(Categories))
                                    {
                                        string[] _categories = Categories.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                                        try
                                        {
                                            FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                                            for (int i = 0; i <= _categories.Length - 1; i++)
                                            {
                                                string newValue = _categories[i].ToString();

                                                if (_categories[i].ToString().Contains("$"))
                                                {
                                                    newValue = _categories[i].ToString().Replace("$", ",");
                                                }

                                                int _cId = GetLookupIDs(newValue, clientcontext, oWeb);

                                                if (_cId != 0)
                                                {
                                                    FieldLookupValue flv = new FieldLookupValue();
                                                    flv.LookupId = _cId;

                                                    lookupFieldValCollection.SetValue(flv, i);
                                                }
                                            }

                                            if (lookupFieldValCollection.Length >= 1)
                                            {
                                                if (lookupFieldValCollection[0] != null)
                                                    _Item["Categorization"] = lookupFieldValCollection;
                                            }

                                            _Item.Update();
                                            clientcontext.Load(_Item);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + Categories + "," + "CategoryFailure");
                                            excelWriterScoringMatrixNew.Flush();
                                        }

                                        try
                                        {
                                            _Item["Modified"] = Modified;
                                            _Item["Editor"] = ModifiedBy;
                                            _Item.Update();
                                            clientcontext.ExecuteQuery();

                                            excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + Categories + "," + "Success");
                                            excelWriterScoringMatrixNew.Flush();
                                        }
                                        catch (Exception ex)
                                        {
                                            excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + Categories + "," + "ModifyFailure");
                                            excelWriterScoringMatrixNew.Flush();
                                        }
                                    }
                                }
                                catch (Exception EX)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + Categories + "," + "ItemIDFailure");
                                    excelWriterScoringMatrixNew.Flush();
                                }

                                _List.EnableVersioning = true;
                                _List.Update();
                                clientcontext.ExecuteQuery();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button37_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "OwnergroupAsOwnerReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SPURL" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            StreamWriter excelWriterScoringMatrixNew1 = null;
            excelWriterScoringMatrixNew1 = System.IO.File.CreateText(textBox2.Text + "\\" + "NoGroupsReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew1.WriteLine("SPURL");
            excelWriterScoringMatrixNew1.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        Web owebt = clientcontext.Web;
                        clientcontext.Load(owebt, oweb => oweb.Title, oweb => oweb.HasUniqueRoleAssignments);//, clientcontext.Web.Title, clientcontext.Web.HasUniqueRoleAssignments);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            if (owebt.HasUniqueRoleAssignments)
                            {
                                GroupCollection AllGroups = clientcontext.Web.RoleAssignments.Groups;
                                clientcontext.Load(AllGroups);
                                clientcontext.ExecuteQuery();

                                if (AllGroups.Count < 1)
                                {
                                    excelWriterScoringMatrixNew1.WriteLine(lstSiteColl[j].ToString());
                                    excelWriterScoringMatrixNew1.Flush();
                                }
                                else
                                {

                                    Group ownergrp2 = clientcontext.Web.RoleAssignments.Groups.GetByName(removingSpecialCharactersForGroups(clientcontext.Web.Title) + " Owners");
                                    clientcontext.Load(ownergrp2);
                                    clientcontext.ExecuteQuery();

                                    foreach (Microsoft.SharePoint.Client.Group grp in AllGroups)
                                    {
                                        try
                                        {
                                            clientcontext.Load(grp);
                                            clientcontext.ExecuteQuery();

                                            grp.Owner = ownergrp2;
                                            grp.Update();
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + grp.Title + " : GroupIssue");
                                            excelWriterScoringMatrixNew.Flush();
                                            continue;
                                        }

                                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Success");
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                }
                            }
                            else
                            {
                                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "InheritedPermissions");
                                excelWriterScoringMatrixNew.Flush();
                            }
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Failure");
                            excelWriterScoringMatrixNew.Flush();
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            excelWriterScoringMatrixNew1.Flush();
            excelWriterScoringMatrixNew1.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        public string removingSpecialCharactersForGroups(string splchr)
        {
            try
            {
                char strChar = '\0';
                string splchars = ":|\"'/[]:<>+=,;?*@";

                for (int i = 0; i <= splchr.Length - 1; i++)
                {
                    strChar = splchr[i];

                    if (splchars.Contains(strChar))
                        splchr = splchr.Replace(strChar, '-');

                    //if (splchars.IndexOf(strChar) == -1)
                    //{
                    //    strSplChaR += strChar.ToString();
                    //}

                }
            }
            catch (Exception ex)
            {
            }
            return splchr;
        }
        private void button38_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteHistoryGroupReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SPURL" + "," + "GroupType");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            List oList = oLists.GetByTitle("SiteHistory");
                            clientcontext.Load(oList);
                            clientcontext.ExecuteQuery();

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = "<View><RowLimit>10</RowLimit></View>";

                            ListItemCollection listItems = oList.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem _Item in listItems)
                            {
                                clientcontext.Load(_Item);
                                clientcontext.ExecuteQuery();

                                if (_Item["Group_Type"].ToString() == "MEMBER_ONLY" || _Item["Group_Type"].ToString() == "Open" || _Item["Group_Type"].ToString() == "Members only")
                                {
                                    try
                                    {
                                        RoleDefinition _cRoleDef = null;
                                        string groupType = _Item["Group_Type"].ToString();

                                        RoleDefinitionCollection _newRoleDefs = clientcontext.Web.RoleDefinitions;
                                        clientcontext.Load(_newRoleDefs);
                                        clientcontext.ExecuteQuery();

                                        if (_Item["Group_Type"].ToString() == "MEMBER_ONLY" || _Item["Group_Type"].ToString() == "Members only")
                                        {
                                            _cRoleDef = _newRoleDefs.GetByName("Read");
                                        }

                                        if (_Item["Group_Type"].ToString() == "Open")
                                        {
                                            _cRoleDef = _newRoleDefs.GetByName("Contribute");// create
                                        }

                                        User CreatedUser = default(User);

                                        try
                                        {
                                            CreatedUser = clientcontext.Web.EnsureUser("Everyone except external users");
                                            clientcontext.Load(CreatedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            CreatedUser = clientcontext.Web.EnsureUser("Rworldadmin@rsharepoint.onmicrosoft.com");
                                            clientcontext.Load(CreatedUser);
                                            clientcontext.ExecuteQuery();
                                        }

                                        Principal _User = clientcontext.CastTo<Principal>(CreatedUser);

                                        RoleDefinitionBindingCollection _rdbColl = new RoleDefinitionBindingCollection(clientcontext);
                                        _rdbColl.Add(_cRoleDef);

                                        clientcontext.Web.RoleAssignments.Add(_User, _rdbColl);
                                        clientcontext.ExecuteQuery();

                                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + groupType);
                                        excelWriterScoringMatrixNew.Flush();

                                        break;

                                    }
                                    catch (Exception ex)
                                    { }
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button39_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteHistoryGroupReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SPURL" + "," + "GroupType");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            List oList = oLists.GetByTitle("SiteHistory");
                            clientcontext.Load(oList);
                            clientcontext.ExecuteQuery();

                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = "<View><RowLimit>10</RowLimit></View>";

                            ListItemCollection listItems = oList.GetItems(camlQuery);
                            clientcontext.Load(listItems);
                            clientcontext.ExecuteQuery();

                            foreach (ListItem _Item in listItems)
                            {
                                clientcontext.Load(_Item);
                                clientcontext.ExecuteQuery();

                                if (_Item["PlaceType"].ToString() == "Project")
                                {
                                    clientcontext.Web.ResetRoleInheritance();// (false, false);
                                    clientcontext.Load(clientcontext.Web);
                                    clientcontext.ExecuteQuery();

                                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Success");
                                    excelWriterScoringMatrixNew.Flush();

                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button40_Click(object sender, EventArgs e)
        {

            //    #region ImportedObjects CSV Reading

            //    DataTable dtImportedObjects = new DataTable();
            //    dtImportedObjects.Columns.AddRange(new DataColumn[2] { new DataColumn("SpaceID", typeof(string)), new DataColumn("SPURL", typeof(string)) });

            //    string csvData = System.IO.File.ReadAllText(textBox1.Text);

            //    foreach (string row in csvData.Split('\n'))
            //    {
            //        if (!string.IsNullOrEmpty(row))
            //        {
            //            dtImportedObjects.Rows.Add();
            //            int i = 0;

            //            foreach (string cell in row.Split(','))
            //            {
            //                dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
            //                i++;
            //            }
            //        }
            //    }

            //    #endregion


            //    StreamWriter excelWriterScoringMatrixNew = null;
            //    excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "PermissionReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            //    excelWriterScoringMatrixNew.WriteLine("ObjectType" + "," + "ListName" + "," + "TagsSet" + "," + "Status" + "," + "SPURL");
            //    excelWriterScoringMatrixNew.Flush();

            //     int count = 0;
            //    //string[] SiteSplit = new string[] { "/Lists/" };
            //    //string[] IDSplit = new string[] { "?ID=" };
            //    //string[] DocumentSplit = new string[] { "/Documents/" };
            //    //string[] TagsSplit = new string[] { "|" };
            //    //string[] FileURLSplit = new string[] { "/1_Uploaded Files/" };
            //    //string[] PageURLSplit = new string[] { "/Pages/" };

            //    foreach (DataRow drImported in dtImportedObjects.Rows)
            //    {
            //        try
            //        {                 
            //            string _importedURL = drImported["ImportedURL"].ToString().Trim();
            //            string _SpaceID = drImported["SpaceID"].ToString().Trim();

            //            this.Text = (count).ToString() + " : " + _importedURL;


            //                AuthenticationManager authManager = new AuthenticationManager();

            //                using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
            //                {
            //                    Web _web = clientcontext.Web;
            //                clientcontext.Load(_web);
            //                clientcontext.ExecuteQuery();                        

            //               string Office365SiteGenericUserID = "Rspaceadmin@rsharepoint.onmicrosoft.com";


            //                _web.BreakRoleInheritance(false, false);
            //                clientcontext.Load(_web);
            //                clientcontext.ExecuteQuery();

            //                try
            //                {

            //                    Microsoft.SharePoint.Client.GroupCollection AllGroups = _web.RoleAssignments.Groups;//.SiteGroups;

            //                    clientcontext.Load(AllGroups);
            //                    clientcontext.ExecuteQuery();

            //                    Microsoft.SharePoint.Client.Group ownergrp = null;

            //                    #region Owners Group Creation
            //                    try
            //                    {
            //                        Microsoft.SharePoint.Client.Group grp = AllGroups.GetByName(_web.Title + " Owners");
            //                        clientcontext.Load(grp);
            //                        clientcontext.ExecuteQuery();
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        try
            //                        {
            //                            GroupCreationInformation gr = new GroupCreationInformation();
            //                            gr.Title = _web.Title + " Owners";

            //                            Microsoft.SharePoint.Client.Group siteG = null;
            //                            try
            //                            {
            //                                siteG = _web.SiteGroups.Add(gr);
            //                                clientcontext.ExecuteQuery();
            //                            }
            //                            catch (Exception ecc)
            //                            {
            //                                siteG = _web.SiteGroups.GetByName(_web.Title + " Owners");
            //                            }

            //                            try
            //                            {
            //                                clientcontext.Load(siteG);
            //                                clientcontext.ExecuteQuery();

            //                                RoleDefinition rd = clientcontext.Web.RoleDefinitions.GetByName("Site Admin");
            //                                RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(clientcontext);
            //                                rdb.Add(rd);
            //                                clientcontext.Web.RoleAssignments.Add(siteG, rdb);
            //                                clientcontext.ExecuteQuery();
            //                            }
            //                            catch
            //                            {
            //                                RoleDefinition rd = clientcontext.Web.RoleDefinitions.GetByName("Full Control");
            //                                RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(clientcontext);
            //                                rdb.Add(rd);
            //                                clientcontext.Web.RoleAssignments.Add(siteG, rdb);
            //                                clientcontext.ExecuteQuery();
            //                            }
            //                        }
            //                        catch (Exception ecc)
            //                        {

            //                        }
            //                        //Microsoft.SharePoint.Client.Group sGroup = _web.SiteGroups.GetByName(_web.Title + " Owners");
            //                        //clientcontext.Load(sGroup);
            //                        //clientcontext.ExecuteQuery();
            //                        //Group secGroup = _web.SiteGroups[_web.Title + "Owners"];

            //                        //    _web.RoleAssignments.Groups.a
            //                        //  clientcontext.ExecuteQuery();


            //                    }
            //                    #endregion

            //                    #region Visitors Group Creation
            //                    try
            //                    {
            //                        Microsoft.SharePoint.Client.Group grp = AllGroups.GetByName(_web.Title + " Visitors");
            //                        clientcontext.Load(grp);
            //                        clientcontext.ExecuteQuery();
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        try
            //                        {
            //                            GroupCreationInformation gr = new GroupCreationInformation();
            //                            gr.Title = _web.Title + " Visitors";

            //                            //Microsoft.SharePoint.Client.Group siteG = _web.SiteGroups.Add(gr);
            //                            //clientcontext.ExecuteQuery();

            //                            Microsoft.SharePoint.Client.Group siteG = null;
            //                            try
            //                            {
            //                                siteG = _web.SiteGroups.Add(gr);
            //                                clientcontext.ExecuteQuery();
            //                            }
            //                            catch (Exception ecc)
            //                            {
            //                                siteG = _web.SiteGroups.GetByName(_web.Title + " Visitors");
            //                            }

            //                            clientcontext.Load(siteG);
            //                            clientcontext.ExecuteQuery();

            //                            RoleDefinition rd = clientcontext.Web.RoleDefinitions.GetByName("Read");
            //                            RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(clientcontext);
            //                            rdb.Add(rd);
            //                            clientcontext.Web.RoleAssignments.Add(siteG, rdb);
            //                            clientcontext.ExecuteQuery();
            //                        }
            //                        catch (Exception ecc)
            //                        {

            //                        }
            //                    }
            //                    #endregion

            //                    #region Members Group Creation
            //                    try
            //                    {
            //                        Microsoft.SharePoint.Client.Group grp = AllGroups.GetByName(_web.Title + " Members");
            //                        clientcontext.Load(grp);
            //                        clientcontext.ExecuteQuery();
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        try
            //                        {
            //                            GroupCreationInformation gr = new GroupCreationInformation();
            //                            gr.Title = _web.Title + " Members";

            //                            //Microsoft.SharePoint.Client.Group siteG = _web.SiteGroups.Add(gr);
            //                            //clientcontext.ExecuteQuery();

            //                            Microsoft.SharePoint.Client.Group siteG = null;
            //                            try
            //                            {
            //                                siteG = _web.SiteGroups.Add(gr);
            //                                clientcontext.ExecuteQuery();
            //                            }
            //                            catch (Exception ecc)
            //                            {
            //                                siteG = _web.SiteGroups.GetByName(_web.Title + " Members");
            //                            }

            //                            clientcontext.Load(siteG);
            //                            clientcontext.ExecuteQuery();

            //                            RoleDefinition rd = clientcontext.Web.RoleDefinitions.GetByName("Contribute");
            //                            RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(clientcontext);
            //                            rdb.Add(rd);
            //                            clientcontext.Web.RoleAssignments.Add(siteG, rdb);
            //                            clientcontext.ExecuteQuery();

            //                        }
            //                        catch (Exception ecc)
            //                        {

            //                        }      

            //                    }
            //                    #endregion


            //                }
            //                catch (Exception ex)
            //                {

            //                }

            //                //  if (SP_Public_Delcarations.SkippedUsermappingFlag == false)
            //                try
            //                {
            //                    if (System.IO.File.Exists(textBox2.Text + "\\" + "Groups&UserPermissions_" + _SpaceID + ".xml"))
            //                    {
            //                        ReadingGroupXMLFile(textBox2.Text + "\\" + "Groups&UserPermissions_" + _SpaceID + ".xml", _SpaceID, "Group");
            //                    }

            //                }
            //                catch (Exception ex)
            //                {

            //                }


            //                try
            //                {
            //                    //here actual
            //                    Microsoft.SharePoint.Client.Group ownergrp2 = _web.RoleAssignments.Groups.GetByName(_web.Title + " Owners");
            //                    clientcontext.Load(ownergrp2);
            //                    clientcontext.ExecuteQuery();

            //                    foreach (Microsoft.SharePoint.Client.Group grp in _web.RoleAssignments.Groups)
            //                    {
            //                        clientcontext.Load(grp);
            //                        clientcontext.ExecuteQuery();

            //                        grp.Owner = ownergrp2;
            //                        grp.Update();
            //                        clientcontext.ExecuteQuery();

            //                    }
            //                }
            //                catch (Exception ex)
            //                {

            //                }





            //                List _List = null;
            //                    string listName = string.Empty;
            //                    string _FilePath = string.Empty;
            //                    string _itemID = string.Empty;

            //                    switch (_objList)
            //                    {
            //                        case "Document":
            //                            listName = "2_Documents and Pages";
            //                            _FilePath = drImported["ImportedURL"].ToString().Split(DocumentSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
            //                            break;

            //                        case "File":
            //                            listName = "1_Uploaded Files";
            //                            _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
            //                            break;

            //                        //case "Announcement":
            //                        //    listName = "Announcements";
            //                        //    _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
            //                        //    break;

            //                        case "Blog":
            //                            listName = "Posts";
            //                            _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
            //                            break;

            //                        case "Discussion":
            //                            listName = "Discussions";
            //                            _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
            //                            break;

            //                        case "Event":
            //                            listName = "Events";
            //                            _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
            //                            break;

            //                        case "Task":
            //                            listName = "Tasks";
            //                            _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
            //                            break;

            //                        case "Idea":
            //                            listName = "Ideas";
            //                            _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
            //                            break;

            //                            //case "Poll":

            //                            //    break;
            //                    }


            //                    try
            //                    {
            //                        _List = clientcontext.Web.Lists.GetByTitle(listName);
            //                        clientcontext.Load(_List);
            //                        clientcontext.ExecuteQuery();
            //                    }
            //                    catch (Exception ex)
            //                    { }

            //                    if (_List != null)
            //                    {
            //                        if (_List.Title == "2_Documents and Pages")
            //                        {
            //                            _List.EnableVersioning = false;
            //                            _List.Update();
            //                            clientcontext.ExecuteQuery();

            //                            _List.ForceCheckout = false;
            //                            _List.Update();
            //                            clientcontext.ExecuteQuery();

            //                            try
            //                            {
            //                                clientcontext.Load(_List.RootFolder);
            //                                clientcontext.ExecuteQuery();

            //                                Folder docFolder = null;

            //                                try
            //                                {
            //                                    docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
            //                                    clientcontext.Load(docFolder);
            //                                    clientcontext.ExecuteQuery();
            //                                }
            //                                catch (Exception ex)
            //                                { }

            //                                if (docFolder != null)
            //                                {
            //                                    ListItem _Item = docFolder.Files.GetByUrl(_FilePath).ListItemAllFields;
            //                                    clientcontext.Load(_Item);
            //                                    clientcontext.ExecuteQuery();

            //                                    DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
            //                                    FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];
            //                                    TaxonomyFieldValueCollection taxFieldValues = _Item["Tags"] as TaxonomyFieldValueCollection;

            //                                    if (taxFieldValues.Count < 1)
            //                                    {
            //                                        TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(oWeb.Context);
            //                                        clientcontext.Load(taxonomySession.TermStores);
            //                                        clientcontext.ExecuteQuery();

            //                                        TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");
            //                                        clientcontext.Load(termStore);
            //                                        clientcontext.ExecuteQuery();

            //                                        clientcontext.Load(termStore.Groups);
            //                                        clientcontext.ExecuteQuery();

            //                                        TermGroup group = termStore.Groups.GetByName("RicohTags");
            //                                        clientcontext.Load(group);
            //                                        clientcontext.ExecuteQuery();

            //                                        clientcontext.Load(group.TermSets);
            //                                        clientcontext.ExecuteQuery();

            //                                        TermSet termSet = group.TermSets.GetByName("TagsTermSet");
            //                                        clientcontext.Load(termSet);
            //                                        clientcontext.ExecuteQuery();


            //                                        Field _taxnomyField = _List.Fields.GetByTitle("Tags");
            //                                        clientcontext.Load(_taxnomyField);
            //                                        clientcontext.ExecuteQuery();

            //                                        TaxonomyField txField = clientcontext.CastTo<TaxonomyField>(_taxnomyField);
            //                                        clientcontext.Load(txField);
            //                                        clientcontext.ExecuteQuery();

            //                                        TaxonomyFieldValueCollection termValues = null;

            //                                        string termValueString = string.Empty;
            //                                        string termId = string.Empty;

            //                                        try
            //                                        {
            //                                            foreach (string tv in TagsColl)
            //                                            {
            //                                                string mtermId = string.Empty;

            //                                                try
            //                                                {
            //                                                    //mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);

            //                                                    if (string.IsNullOrEmpty(mtermId))
            //                                                    {
            //                                                        mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);
            //                                                    }

            //                                                    if (!string.IsNullOrEmpty(mtermId))
            //                                                        termValueString += "1033" + ";#" + tv + "|" + mtermId + ";#";

            //                                                }
            //                                                catch (Exception ex)
            //                                                {
            //                                                    continue;
            //                                                }
            //                                            }

            //                                            //if (taxFieldValues.Count > 0)
            //                                            //{
            //                                            termValueString = termValueString.Remove(termValueString.Length - 2);
            //                                            termValues = new TaxonomyFieldValueCollection(clientcontext, termValueString,
            //                                                txField);

            //                                            txField.SetFieldValueByValueCollection(_Item, termValues);

            //                                            _Item.Update();
            //                                            clientcontext.Load(_Item);
            //                                            clientcontext.ExecuteQuery();

            //                                            _Item["Modified"] = Modified;
            //                                            _Item["Editor"] = ModifiedBy;

            //                                            _Item.Update();
            //                                            clientcontext.Load(_Item);
            //                                            clientcontext.ExecuteQuery();

            //                                            excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Success" + "," + _objectURL);
            //                                            excelWriterScoringMatrixNew.Flush();
            //                                            //}
            //                                        }
            //                                        catch (Exception EX)
            //                                        {
            //                                            excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Error : " + EX.Message + "," + _objectURL);
            //                                            excelWriterScoringMatrixNew.Flush();
            //                                        }
            //                                    }
            //                                }
            //                            }
            //                            catch (Exception ex)
            //                            {
            //                                //excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Failure due to : " + ex.Message);
            //                                //excelWriterScoringMatrixNew.Flush();
            //                            }

            //                            _List.EnableVersioning = true;
            //                            _List.Update();
            //                            clientcontext.ExecuteQuery();

            //                            _List.ForceCheckout = true;
            //                            _List.Update();
            //                            clientcontext.ExecuteQuery();
            //                        }
            //                        else
            //                        {
            //                            _List.EnableVersioning = false;
            //                            _List.Update();
            //                            clientcontext.ExecuteQuery();

            //                            try
            //                            {
            //                                ListItem _Item = _List.GetItemById(_itemID);
            //                                clientcontext.Load(_Item);
            //                                clientcontext.ExecuteQuery();

            //                                DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
            //                                FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];
            //                                TaxonomyFieldValueCollection taxFieldValues = _Item["Tags"] as TaxonomyFieldValueCollection;

            //                                if (taxFieldValues.Count < 1)
            //                                {
            //                                    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(oWeb.Context);
            //                                    clientcontext.Load(taxonomySession.TermStores);
            //                                    clientcontext.ExecuteQuery();

            //                                    TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");
            //                                    clientcontext.Load(termStore);
            //                                    clientcontext.ExecuteQuery();

            //                                    clientcontext.Load(termStore.Groups);
            //                                    clientcontext.ExecuteQuery();

            //                                    TermGroup group = termStore.Groups.GetByName("RicohTags");
            //                                    clientcontext.Load(group);
            //                                    clientcontext.ExecuteQuery();

            //                                    clientcontext.Load(group.TermSets);
            //                                    clientcontext.ExecuteQuery();

            //                                    TermSet termSet = group.TermSets.GetByName("TagsTermSet");
            //                                    clientcontext.Load(termSet);
            //                                    clientcontext.ExecuteQuery();


            //                                    Field _taxnomyField = _List.Fields.GetByTitle("Tags");
            //                                    clientcontext.Load(_taxnomyField);
            //                                    clientcontext.ExecuteQuery();

            //                                    TaxonomyField txField = clientcontext.CastTo<TaxonomyField>(_taxnomyField);
            //                                    clientcontext.Load(txField);
            //                                    clientcontext.ExecuteQuery();

            //                                    TaxonomyFieldValueCollection termValues = null;

            //                                    string termValueString = string.Empty;
            //                                    string termId = string.Empty;

            //                                    try
            //                                    {
            //                                        foreach (string tv in TagsColl)
            //                                        {
            //                                            string mtermId = string.Empty;

            //                                            try
            //                                            {
            //                                                //mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);

            //                                                if (string.IsNullOrEmpty(mtermId))
            //                                                {
            //                                                    mtermId = GetTermIdForTerm(tv, termSet.Id, termSet, termStore, clientcontext);
            //                                                }

            //                                                if (!string.IsNullOrEmpty(mtermId))
            //                                                    termValueString += "1033" + ";#" + tv + "|" + mtermId + ";#";

            //                                            }
            //                                            catch (Exception ex)
            //                                            {
            //                                                continue;
            //                                            }
            //                                        }

            //                                        //if (taxFieldValues.Count > 0)
            //                                        //{
            //                                        termValueString = termValueString.Remove(termValueString.Length - 2);
            //                                        termValues = new TaxonomyFieldValueCollection(clientcontext, termValueString,
            //                                            txField);

            //                                        txField.SetFieldValueByValueCollection(_Item, termValues);

            //                                        _Item.Update();
            //                                        clientcontext.Load(_Item);
            //                                        clientcontext.ExecuteQuery();

            //                                        _Item["Modified"] = Modified;
            //                                        _Item["Editor"] = ModifiedBy;

            //                                        _Item.Update();
            //                                        clientcontext.Load(_Item);
            //                                        clientcontext.ExecuteQuery();

            //                                        excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Success" + "," + _objectURL);
            //                                        excelWriterScoringMatrixNew.Flush();
            //                                        //}
            //                                    }
            //                                    catch (Exception EX)
            //                                    {
            //                                        excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Error : " + EX.Message + "," + _objectURL);
            //                                        excelWriterScoringMatrixNew.Flush();
            //                                    }
            //                                }
            //                            }
            //                            catch (Exception EX)
            //                            {
            //                                excelWriterScoringMatrixNew.WriteLine(_objList + "," + _List.Title + "," + TagsSet + "," + "Error : " + EX.Message + "," + _objectURL);
            //                                excelWriterScoringMatrixNew.Flush();
            //                            }

            //                            _List.EnableVersioning = true;
            //                            _List.Update();
            //                            clientcontext.ExecuteQuery();
            //                        }
            //                    }
            //                }

            //        }
            //        catch (Exception ex)
            //        {
            //            continue;
            //        }
            //    }

            //    excelWriterScoringMatrixNew.Flush();
            //    excelWriterScoringMatrixNew.Close();

            //    this.Text = "Process completed successfully.";
            //    MessageBox.Show("Process Completed");
        }
        private void ReadingGroupXMLFile(string xmlFilePath, string _placeID, string _placetype)
        {
            //try
            //{
            //    ///'' Reading XML File ........... 
            //    XmlDocument Groupobjxml = new XmlDocument();

            //    Groupobjxml.Load(xmlFilePath);
            //    XmlNodeList parentnode = Groupobjxml.SelectNodes("/Places");

            //    XmlNodeList xnList = null;
            //    switch (_placetype)
            //    {
            //        case "space":
            //            xnList = Groupobjxml.SelectNodes("/Places/Space[@spaceId='" + _placeID + "']");

            //            break;
            //        case "group":
            //            xnList = Groupobjxml.SelectNodes("/Places/Group[@groupId='" + _placeID + "']");

            //            break;
            //            //case "project":
            //            //     xnList = Groupobjxml.SelectNodes("/Places/Space[@spaceId='" + _placeID + "']");
            //            //  break;
            //    }
            //    //groupid

            //    // XmlNodeList xnList = Groupobjxml.SelectNodes("/Spaces/Space[@spaceId='" + _placeID + "']");
            //    foreach (XmlNode xn in xnList)
            //    {
            //        Add_UsersToSharePoinSite(xn, _placeID, _placetype);
            //    }


            //}
            //catch (Exception ex)
            //{
            //    ErrorLogger er = new ErrorLogger();
            //    er.WriteToErrorLog(ex.Message, ex.StackTrace, "Error");



            //}
            ///''' End of Reading XML File ........... 
        }
        private void Add_UsersToSharePoinSite(XmlNode xn, string _placeID, string placeType)
        {

            //try
            //{               
            //    {

            //        for (int i = 0; i <= xn.ChildNodes.Count - 1; i++)
            //        {
            //            try
            //            {
            //                XmlNode objxml = xn.ChildNodes[i];

            //                switch (objxml.Name)
            //                { 
            //                    case "Members":
            //                        foreach (XmlNode child_objxml in objxml.ChildNodes)
            //                        {


            //                            string child_userName = child_objxml.Attributes["userName"].Value.ToString();

            //                            string child_userRole = string.Empty;

            //                            if (child_objxml.Attributes["memberType"].Value.ToString() == "None")
            //                            {
            //                                child_userRole = "Read";
            //                            }
            //                            else
            //                            {
            //                                //child_userRole = "Read";
            //                                // _groupRole = "Full Control";
            //                                //uncomment
            //                                try
            //                                {
            //                                    child_userRole = COMPermissions.MappingDictionary[child_objxml.Attributes["memberType"].Value.ToString().Trim()].ToString();
            //                                }
            //                                catch
            //                                {
            //                                    child_userRole = child_objxml.Attributes["memberType"].Value.ToString().Trim();
            //                                }
            //                            }

            //                            Import_User(child_userName, child_userRole, _placeID,
            //                                child_objxml.Attributes["memberType"].Value.ToString().Trim(), placeType);
            //                        }
            //                        break;
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                continue;
            //            }
            //        }
            //    }


            //}
            //catch (Exception ex)
            //{
            //    ErrorLogger er = new ErrorLogger();
            //    er.WriteToErrorLog(ex.Message, ex.StackTrace, "Error");

            //}
        }
        private void Import_User(string _UserName, string _UserRole, string _placeID, string _jiveRole, string placeType)
        {
            //string _Status = string.Empty;
            //try
            //{
            //    // ClientContext _cContext= CreateContext2("");

            //    using (ClientContext _cContext = SP_Public_Delcarations._currentclientContext)
            //    {

            //        // _cContext = CreateContext2("");
            //        //WriteUserImportLog(_placeID + "\t" + "User" + "-" + _UserName + "\t" +
            //        //    _UserName + "\t" + _jiveRole + "\t" +
            //        //    _UserRole + "\t" + "Imported" + "\t" + "--", _placeID);

            //        SP_Public_Delcarations._currentWeb = _cContext.Web;
            //        _cContext.Load(SP_Public_Delcarations._currentWeb);
            //        _cContext.ExecuteQuery();

            //        Web _Web = SP_Public_Delcarations._currentWeb; //_cContext.Web;

            //        _cContext.Load(_Web);
            //        _cContext.ExecuteQuery();





            //        //here proj
            //        if (placeType != "project")
            //        {
            //            _Web.BreakRoleInheritance(true, false);
            //            _cContext.ExecuteQuery();

            //        }

            //        User CreatedUser = default(User);
            //        try
            //        {
            //            //Ensure user in Site
            //            CreatedUser = _cContext.Web.EnsureUser(_UserName);
            //            _cContext.Load(CreatedUser);
            //            _cContext.ExecuteQuery();
            //        }
            //        catch (Exception ex)
            //        {
            //            CreatedUser = _cContext.Web.EnsureUser(SP_Public_Delcarations.Office365SiteGenericUserID);
            //            _cContext.Load(CreatedUser);
            //            _cContext.ExecuteQuery();
            //        }

            //        //RoleDefinitionCollection _newRoleDefs = _cContext.Web.RoleDefinitions;
            //        //_cContext.Load(_newRoleDefs);
            //        //_cContext.ExecuteQuery();

            //        //RoleDefinition _cRoleDef = _newRoleDefs.GetByName(_UserRole);



            //        ////get the user object
            //        //Principal _group = _cContext.CastTo<Principal>(CreatedUser);

            //        ////add the role definition to the collection
            //        //RoleDefinitionBindingCollection _rdbColl = new RoleDefinitionBindingCollection(_cContext);
            //        //_rdbColl.Add(_cRoleDef);

            //        //create a RoleAssigment with the user and role definition
            //        // _Web.RoleAssignments.Add(_group, _rdbColl);
            //        // _cContext.ExecuteQuery();


            //        Microsoft.SharePoint.Client.GroupCollection AllGroups = SP_Public_Delcarations._currentWeb.SiteGroups;


            //        switch (_UserRole)
            //        {
            //            case "Read":
            //            case "View":
            //            case "ReadOnly":
            //                {

            //                    Microsoft.SharePoint.Client.Group g = AllGroups.GetByName(SP_Public_Delcarations._currentWeb.Title + " " + "Visitors");

            //                    g.Users.AddUser(CreatedUser);

            //                    _cContext.ExecuteQuery();


            //                    //   _cContext.ExecuteQuery();

            //                }


            //                break;

            //            case "Administer":
            //            case "Admin":
            //            case "Admin + Moderate":
            //            case "Full Control":
            //            case "Site Admin":
            //                {

            //                    Microsoft.SharePoint.Client.Group g = AllGroups.GetByName(SP_Public_Delcarations._currentWeb.Title + " " + "Owners");
            //                    g.Users.AddUser(CreatedUser);
            //                    _cContext.ExecuteQuery();

            //                }

            //                break;

            //            default:
            //                {

            //                    Microsoft.SharePoint.Client.Group g = AllGroups.GetByName(SP_Public_Delcarations._currentWeb.Title + " " + "Members");
            //                    g.Users.AddUser(CreatedUser);
            //                    _cContext.ExecuteQuery();



            //                }


            //                break;



            //        }







            //        WriteUserImportLog(_placeID + "\t" + "User" + "-" + _UserName + "\t" +
            //              _UserName + "\t" + _jiveRole + "\t" +
            //              _UserRole + "\t" + "Imported" + "\t" + "--", _placeID);




            //        backgroundWorker2.ReportProgress(0, "User added: " + CreatedUser);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    backgroundWorker2.ReportProgress(0, "Failed User : " + _UserName + " due to : " +
            //                       ex.Message.ToString());

            //    WriteUserImportLog(_placeID + "\t" + "User" + "-" + _UserName + "\t" +
            //                _UserName + "\t" + _jiveRole + "\t" +
            //                _UserRole + "\t" + "Not Imported" + "\t" + ex.Message.ToString(), _placeID);

            //}
        }
        private void button41_Click(object sender, EventArgs e)
        {
            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[7] { new DataColumn("ObjectId", typeof(string)), new DataColumn("CreatedDate", typeof(string)), new DataColumn("Createdby", typeof(string)), new DataColumn("ModifiedDate", typeof(string)), new DataColumn("ModifiedBy", typeof(string)), new DataColumn("SPURL", typeof(string)), new DataColumn("ListName", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion          

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SBD_FileMetadataReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SPURL" + "," + "ListName" + "," + "ItemID" + "," + "CreatedDate" + "," + "CreatedBy" + "," + "ModifiedDate" + "," + "ModifiedBy" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string _objId = drImported["ObjectId"].ToString().Trim();
                    string _importedURL = drImported["SPURL"].ToString().Trim();
                    string ListName = drImported["ListName"].ToString().Trim();

                    string createdDate = drImported["CreatedDate"].ToString().Trim();
                    string modifiedDate = drImported["ModifiedDate"].ToString().Trim();
                    string Author = drImported["Createdby"].ToString().Trim();
                    string Editor = drImported["ModifiedBy"].ToString().Trim();

                    this.Text = (count).ToString() + " : " + _objId;
                    count++;

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "sreekanth.grandhasila@sbdinc.com", "FN5O!CQa"))
                    {
                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        #region Web.EnsureUser TEST

                        ////Greg.Keier@sbdinc.com
                        //User ModiUser = default(User);
                        //try
                        //{
                        //    ModiUser = clientcontext.Web.EnsureUser("Greg.Keier@sbdinc.com");
                        //    clientcontext.Load(ModiUser);
                        //    clientcontext.ExecuteQuery();
                        //}
                        //catch (Exception ex)
                        //{
                        //    ModiUser = clientcontext.Web.EnsureUser("sreekanth.grandhasila@sbdinc.com");
                        //    clientcontext.Load(ModiUser);
                        //    clientcontext.ExecuteQuery();
                        //} 

                        #endregion

                        List _List = null;

                        try
                        {
                            _List = clientcontext.Web.Lists.GetByTitle(ListName);
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        { }

                        if (_List != null)
                        {
                            _List.EnableVersioning = false;
                            _List.Update();
                            clientcontext.ExecuteQuery();

                            ListItem _Item = _List.GetItemById(_objId);
                            clientcontext.Load(_Item);
                            clientcontext.ExecuteQuery();

                            DateTime dtModified = getdateformat(modifiedDate);
                            DateTime dtCreated = getdateformat(createdDate);

                            User CreatedUser = default(User);
                            try
                            {
                                CreatedUser = clientcontext.Web.EnsureUser(Author);
                                clientcontext.Load(CreatedUser);
                                clientcontext.ExecuteQuery();

                            }
                            catch (Exception ex)
                            {
                                CreatedUser = clientcontext.Web.EnsureUser("sreekanth.grandhasila@sbdinc.com");
                                clientcontext.Load(CreatedUser);
                                clientcontext.ExecuteQuery();
                            }
                            FieldUserValue CreatedUserValue = new FieldUserValue();
                            CreatedUserValue.LookupId = CreatedUser.Id;

                            User ModifiedUser = default(User);
                            try
                            {
                                ModifiedUser = clientcontext.Web.EnsureUser(Editor);
                                clientcontext.Load(ModifiedUser);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            {
                                ModifiedUser = clientcontext.Web.EnsureUser("sreekanth.grandhasila@sbdinc.com");
                                clientcontext.Load(ModifiedUser);
                                clientcontext.ExecuteQuery();
                            }

                            FieldUserValue ModifiedUserValue = new FieldUserValue();
                            ModifiedUserValue.LookupId = CreatedUser.Id;

                            try
                            {
                                _Item["Created"] = dtCreated;
                                _Item["Author"] = CreatedUserValue;
                                _Item["Modified"] = dtModified;
                                _Item["Editor"] = ModifiedUserValue;

                                _Item.Update();
                                clientcontext.ExecuteQuery();

                                excelWriterScoringMatrixNew.WriteLine(_importedURL + "," + ListName + "," + _objId + "," + createdDate + "," + Author + "," + modifiedDate + "," + Editor + "," + "Success");
                                excelWriterScoringMatrixNew.Flush();
                            }
                            catch (Exception ex)
                            {
                                excelWriterScoringMatrixNew.WriteLine(_importedURL + "," + ListName + "," + _objId + "," + createdDate + "," + Author + "," + modifiedDate + "," + Editor + "," + "Failure : " + ex.Message);
                                excelWriterScoringMatrixNew.Flush();
                            }

                            _List.EnableVersioning = true;
                            _List.Update();
                            clientcontext.ExecuteQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button42_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "PagesListDocumentsReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SPURL" + "," + "PagesListExist" + "," + "DocumnetFolderExist" + "," + "DocumentsCount");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                string SSSS = lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        List _List = null;

                        try
                        {
                            _List = oLists.GetByTitle("2_Documents and Pages"); ;
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        { }

                        if (_List != null)
                        {
                            clientcontext.Load(_List.RootFolder);
                            clientcontext.ExecuteQuery();

                            Folder docFolder = null;
                            try
                            {
                                docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                clientcontext.Load(docFolder);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            {
                            }

                            if (docFolder != null)
                            {
                                try
                                {
                                    FileCollection oFiles = docFolder.Files;
                                    clientcontext.Load(oFiles);
                                    clientcontext.ExecuteQuery();

                                    int DocumentsCount = oFiles.Count;

                                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Exist" + "Exist" + "," + DocumentsCount.ToString());
                                    excelWriterScoringMatrixNew.Flush();
                                }
                                catch (Exception ex)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Exist" + "Exist" + "," + "ItemCountError");
                                    excelWriterScoringMatrixNew.Flush();
                                }
                            }
                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Exist" + "NotExist" + "," + "NA");
                            excelWriterScoringMatrixNew.Flush();
                        }
                        else
                        {
                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "NotExist" + "NotExist" + "," + "NA");
                            excelWriterScoringMatrixNew.Flush();
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button43_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteDeleteReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SPURL" + "," + "DeleteStatus");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                string SSSS = lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        oWeb.DeleteObject();
                        clientcontext.ExecuteQuery();

                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Success");
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Failure : " + ex.Message);
                    excelWriterScoringMatrixNew.Flush();
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button44_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("SiteUrl", typeof(string)), new DataColumn("Modified", typeof(string)), new DataColumn("ModifiedBy", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion           

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocModifyReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("ObjectURL" + "," + "Modified" + "," + "ModifiedBy" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;
            string[] DocumentSplit = new string[] { "/Documents/" };
            string[] PageURLSplit = new string[] { "/Pages/" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string _objectURL = drImported["SiteUrl"].ToString().Trim();
                    string docModified = drImported["Modified"].ToString().Trim();
                    string docModifiedBy = drImported["ModifiedBy"].ToString().Trim();
                    string _importedURL = drImported["SiteUrl"].ToString().Trim();

                    if (_importedURL.Contains("/Pages/"))
                    {
                        _importedURL = drImported["SiteUrl"].ToString().Split(PageURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }

                    this.Text = (count).ToString() + " : " + _objectURL;
                    count++;

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        List _List = null;
                        string listName = string.Empty;
                        string _FilePath = string.Empty;

                        //listName = "Documents and Pages";
                        _FilePath = drImported["SiteUrl"].ToString().Split(DocumentSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();

                        try
                        {
                            _List = clientcontext.Web.Lists.GetByTitle(textBox5.Text);
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        { }

                        if (_List != null)
                        {
                            //if (_List.Title == "2_Documents and Pages")
                            {
                                _List.EnableVersioning = false;
                                _List.Update();
                                clientcontext.ExecuteQuery();

                                //_List.ForceCheckout = false;
                                //_List.Update();
                                //clientcontext.ExecuteQuery();

                                try
                                {
                                    clientcontext.Load(_List.RootFolder);
                                    clientcontext.ExecuteQuery();

                                    Folder docFolder = null;

                                    try
                                    {
                                        docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                        clientcontext.Load(docFolder);
                                        clientcontext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    { }

                                    if (docFolder != null)
                                    {
                                        ListItem _Item = docFolder.Files.GetByUrl(_FilePath).ListItemAllFields;
                                        //ListItem _Item = _List.GetItemById("43");
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();

                                        User ModifiedUser = default(User);
                                        try
                                        {
                                            ModifiedUser = clientcontext.Web.EnsureUser(docModifiedBy);
                                            clientcontext.Load(ModifiedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            ModifiedUser = clientcontext.Web.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                            clientcontext.Load(ModifiedUser);
                                            clientcontext.ExecuteQuery();
                                        }

                                        FieldUserValue ModifiedUserValue = new FieldUserValue();
                                        ModifiedUserValue.LookupId = ModifiedUser.Id;

                                        DateTime Modified = getdateformat(docModified);

                                        _Item["Modified"] = Modified;
                                        _Item["Editor"] = ModifiedUserValue;
                                        _Item.Update();
                                        clientcontext.ExecuteQuery();

                                        excelWriterScoringMatrixNew.WriteLine(_objectURL + "," + docModified + "," + docModifiedBy + "," + "Success");
                                        excelWriterScoringMatrixNew.Flush();

                                    }
                                }
                                catch (Exception ex)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(_objectURL + "," + docModified + "," + docModifiedBy + "," + "Failure : " + ex.Message);
                                    excelWriterScoringMatrixNew.Flush();
                                }

                                _List.EnableVersioning = true;
                                _List.Update();
                                clientcontext.ExecuteQuery();

                                //_List.ForceCheckout = true;
                                //_List.Update();
                                //clientcontext.ExecuteQuery();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button45_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "ItemDetailedReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "ItemTitle" + "," + "Modified" + "," + "ModifiedBy" + "," + "Tags");
            excelWriterScoringMatrixNew.Flush();

            List<string> ListNames = new List<string>();

            ListNames.Add("Site Assets");
            ListNames.Add("2_Documents and Pages");
            ListNames.Add("1_Uploaded Files");
            ListNames.Add("Discussions");
            ListNames.Add("Events");
            ListNames.Add("Announcements");
            ListNames.Add("Tasks");
            ListNames.Add("Posts");
            ListNames.Add("SiteHistory");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(),
                        "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        _cContext.Load(_cContext.Web);
                        _cContext.ExecuteQuery();

                        Web _Web = _cContext.Web;
                        List _List = null;

                        foreach (string ls in ListNames)
                        {

                            try
                            {
                                _List = _cContext.Web.Lists.GetByTitle(ls);
                                _cContext.Load(_List);
                                _cContext.ExecuteQuery();

                                CamlQuery camlQuery = new CamlQuery();
                                camlQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>";

                                ListItemCollection listItems = _List.GetItems(camlQuery);
                                _cContext.Load(listItems);
                                _cContext.ExecuteQuery();

                                foreach (ListItem oItem in listItems)
                                {
                                    try
                                    {
                                        _cContext.Load(oItem);
                                        _cContext.ExecuteQuery();

                                        string oTitle = string.Empty;
                                        string Tags = string.Empty;
                                        string Modified = string.Empty;
                                        string ModifiedBy = string.Empty;
                                        try
                                        {
                                            oTitle = oItem["Title"].ToString();
                                        }
                                        catch (Exception ex)
                                        {

                                        }
                                        try
                                        {
                                            TaxonomyFieldValueCollection taxFieldValues = oItem["Tags"] as TaxonomyFieldValueCollection;

                                            foreach (TaxonomyFieldValue tv in taxFieldValues)
                                            {
                                                if (tv != null)
                                                {
                                                    Tags = tv.Label.ToString() + "|";
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {

                                        }

                                        try
                                        {
                                            DateTime dtModified = Convert.ToDateTime(oItem["Modified"]);
                                            Modified = dtModified.ToString();

                                            FieldUserValue userModifiedBy = (FieldUserValue)oItem["Editor"];
                                            ModifiedBy = userModifiedBy.Email;
                                        }
                                        catch (Exception es)
                                        {

                                        }

                                        excelWriterScoringMatrixNew.WriteLine(_cContext.Web.Url + "," + ls + "," + oItem.Id.ToString() + "," + oTitle + "," + Modified + "," + ModifiedBy + "," + Tags);
                                        excelWriterScoringMatrixNew.Flush();

                                    }
                                    catch (Exception ex)
                                    {
                                        continue;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            #endregion

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button46_Click(object sender, EventArgs e)
        {
            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            //dtImportedObjects.Columns.AddRange(new DataColumn[5] { new DataColumn("SiteURL", typeof(string)), new DataColumn("ListName", typeof(string)), new DataColumn("ItemID", typeof(string)), new DataColumn("Modified", typeof(string)), new DataColumn("ModifiedBy", typeof(string)) });
            dtImportedObjects.Columns.AddRange(new DataColumn[7] { new DataColumn("SiteURL", typeof(string)), new DataColumn("ListName", typeof(string)), new DataColumn("ItemID", typeof(string)), new DataColumn("Modified", typeof(string)), new DataColumn("ModifiedBy", typeof(string)), new DataColumn("Created", typeof(string)), new DataColumn("CreatedBy", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion           

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocModifyReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "Modified" + "," + "ModifiedBy" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;
            string[] DocumentSplit = new string[] { "/Documents/" };
            string[] PageURLSplit = new string[] { "/Pages/" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string _objectURL = drImported["SiteUrl"].ToString().Trim();
                    string lstName = drImported["ListName"].ToString().Trim();
                    string _oItemID = drImported["ItemID"].ToString().Trim();
                    string docCreated = drImported["Created"].ToString().Trim();
                    string docCreatedBy = drImported["CreatedBy"].ToString().Trim();
                    string docModified = drImported["Modified"].ToString().Trim();
                    string docModifiedBy = drImported["ModifiedBy"].ToString().Trim();
                    string _importedURL = drImported["SiteUrl"].ToString().Trim();

                    if (_importedURL.Contains("/Pages/"))
                    {
                        _importedURL = drImported["SiteUrl"].ToString().Split(PageURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }

                    this.Text = (count).ToString() + " : " + _objectURL;
                    count++;

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_objectURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        List _List = null;
                        string listName = string.Empty;
                        string _FilePath = string.Empty;


                        try
                        {//Announcements
                            _List = clientcontext.Web.Lists.GetByTitle(lstName);
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        { }

                        if (_List != null)
                        {
                            _List.EnableVersioning = false;
                            _List.Update();
                            clientcontext.ExecuteQuery();


                            try
                            {
                                ListItem _Item = _List.GetItemById(_oItemID);
                                clientcontext.Load(_Item);
                                clientcontext.ExecuteQuery();

                                User CreatedUser = default(User);
                                try
                                {
                                    CreatedUser = clientcontext.Web.EnsureUser(docCreatedBy);
                                    clientcontext.Load(CreatedUser);
                                    clientcontext.ExecuteQuery();
                                }
                                catch (Exception ex)
                                {
                                    CreatedUser = clientcontext.Web.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                    clientcontext.Load(CreatedUser);
                                    clientcontext.ExecuteQuery();
                                }
                                FieldUserValue CreatedUserValue = new FieldUserValue();
                                CreatedUserValue.LookupId = CreatedUser.Id;

                                DateTime Created = getdateformat(docCreated);

                                User ModifiedUser = default(User);
                                try
                                {
                                    ModifiedUser = clientcontext.Web.EnsureUser(docModifiedBy);
                                    clientcontext.Load(ModifiedUser);
                                    clientcontext.ExecuteQuery();
                                }
                                catch (Exception ex)
                                {
                                    ModifiedUser = clientcontext.Web.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                    clientcontext.Load(ModifiedUser);
                                    clientcontext.ExecuteQuery();
                                }
                                FieldUserValue ModifiedUserValue = new FieldUserValue();
                                ModifiedUserValue.LookupId = ModifiedUser.Id;

                                DateTime Modified = getdateformat(docModified);

                                //FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];
                                //DateTime Modified = Convert.ToDateTime(_Item["Created"]);
                                //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);

                                _Item["Created"] = Created;
                                _Item["Author"] = CreatedUserValue;

                                _Item["Modified"] = Modified;
                                _Item["Editor"] = ModifiedUserValue;
                                _Item.Update();
                                clientcontext.ExecuteQuery();

                                excelWriterScoringMatrixNew.WriteLine(_objectURL + "," + lstName + "," + _oItemID + "," + docModified + "," + docModifiedBy + "," + "Success");
                                excelWriterScoringMatrixNew.Flush();

                            }
                            catch (Exception ex)
                            {
                                excelWriterScoringMatrixNew.WriteLine(_objectURL + "," + lstName + "," + _oItemID + "," + docModified + "," + docModifiedBy + "," + "Failure");
                                excelWriterScoringMatrixNew.Flush();
                            }

                            _List.EnableVersioning = true;
                            _List.Update();
                            clientcontext.ExecuteQuery();

                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private void button47_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remaining




            for (int j = 1; j <= lstSiteColl.Count - 1; j++)
            {


                string[] ImportedURL = System.Text.RegularExpressions.Regex.Split(lstSiteColl[j].ToString().Trim(), "/Lists");
                string spURl = ImportedURL[0].ToString();
                string listURL = ImportedURL[1].ToString();
                string ID = System.Text.RegularExpressions.Regex.Split(ImportedURL[1].ToString(), "ID=")[1].ToString();


                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(spURl, "migjive3@gwl.bz", "Verinon@2018"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        if (listURL.ToString().Contains("Comments"))
                        {
                            try
                            {
                                #region VIEW for "Status" List

                                bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, "Comments"));

                                if (_dListExist)
                                {
                                    try
                                    {
                                        List oList = _Lists.GetByTitle("Comments");
                                        clientcontext.Load(oList);
                                        clientcontext.ExecuteQuery();


                                        ListItem _Item = oList.GetItemById(ID);
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();


                                        _Item.DeleteObject();
                                        clientcontext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {

                                    }

                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }

                        }
                        else if (listURL.ToString().Contains("Discussions"))
                        {
                            try
                            {
                                #region VIEW for "Status" List

                                bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, "Discussions"));

                                if (_dListExist)
                                {
                                    try
                                    {
                                        List oList = _Lists.GetByTitle("Discussions");
                                        clientcontext.Load(oList);
                                        clientcontext.ExecuteQuery();


                                        ListItem _Item = oList.GetItemById(ID);
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();


                                        _Item.DeleteObject();
                                        clientcontext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {

                                    }

                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");

            #endregion
        }
        private void button48_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion


            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "MarkAsFeatureedDeleted" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ListName");
            excelWriterScoringMatrixNew.Flush();

            List<string> ListNames = new List<string>();

            ListNames.Add("Site Assets");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                //string url = "https://rsharepoint.sharepoint.com/sites/rworldgroups/solutioncenter-team/what-we-sell-conversion-to-tiles";
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(),
                        "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        _cContext.Load(_cContext.Web);
                        _cContext.ExecuteQuery();

                        Web _Web = _cContext.Web;
                        List _List = null;

                        foreach (string ls in ListNames)
                        {

                            try
                            {
                                _List = _cContext.Web.Lists.GetByTitle(ls);
                                _cContext.Load(_List);
                                _cContext.ExecuteQuery();

                                FieldCollection FldColl = _List.Fields;
                                _cContext.Load(FldColl);
                                _cContext.ExecuteQuery();
                                bool TagCateExist = false;


                                try
                                {

                                    Field tagField1 = FldColl.GetByTitle("MarkasFeatured");
                                    _cContext.Load(tagField1);
                                    _cContext.ExecuteQuery();

                                    tagField1.DeleteObject();
                                    _cContext.ExecuteQuery();

                                    excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ListName" + "," + "Success");
                                    excelWriterScoringMatrixNew.Flush();
                                }
                                catch (Exception ex)
                                {
                                    excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "Failed" + "," + ex.Message.ToString());
                                    excelWriterScoringMatrixNew.Flush();
                                }




                            }
                            catch (Exception ex)
                            {

                            }
                        }
                    }
                }
                catch (Exception ex)
                { }
            }
        }
        private void button49_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));




            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remaining


            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocModifyReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "ItemTitle" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {


                string[] ImportedURL = System.Text.RegularExpressions.Regex.Split(lstSiteColl[j].ToString().Trim(), "/Lists");
                string spURl = ImportedURL[0].ToString();
                string listURL = ImportedURL[1].ToString();
                string IDAfter = System.Text.RegularExpressions.Regex.Split(ImportedURL[1].ToString(), "ID=")[1].ToString();


                string ID = IDAfter.ToString().Split(new char[] { '|' })[0].ToString();  //System.Text.RegularExpressions.Regex.Split(IDAfter, "|")[0].ToString();
                string Archive = IDAfter.ToString().Split(new char[] { '|' })[1].ToString();

                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(spURl, "migjive3@gwl.bz", "Verinon@2018"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        if (listURL.ToString().Contains("Shared Documents"))
                        {
                            try
                            {
                                #region VIEW for "Status" List

                                bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, "Documents"));

                                if (_dListExist)
                                {
                                    try
                                    {
                                        List oList = _Lists.GetByTitle("Documents");
                                        clientcontext.Load(oList);
                                        clientcontext.ExecuteQuery();

                                        ListItem _Item = oList.GetItemById(ID);
                                        clientcontext.Load(_Item, i => i["Title"], i => i["FileDirRef"]);
                                        clientcontext.ExecuteQuery();

                                        string title = Convert.ToString(_Item["Name"]);

                                        #region PArent Folder delete for Documents

                                        //string folderUrl = (string)_Item["FileDirRef"];
                                        //Folder parentFolder = clientcontext.Web.GetFolderByServerRelativeUrl(folderUrl);

                                        //clientcontext.Load(parentFolder);
                                        //clientcontext.ExecuteQuery();

                                        //// MessageBox.Show(parentFolder.Name);
                                        ///// MessageBox.Show(title); 

                                        //parentFolder.DeleteObject();
                                        //clientcontext.ExecuteQuery(); 

                                        #endregion                                       

                                        _Item.DeleteObject();
                                        clientcontext.ExecuteQuery();

                                        //string folderUrl = (string)_Item["FileDirRef"];



                                        // file.Context.Load(parentFolder);
                                        // file.Context.ExecuteQuery();

                                        ///   MessageBox.Show(Archive);
                                        if (Archive == "A")
                                        {
                                            excelWriterScoringMatrixNew.WriteLine(spURl + "," + "Documents" + ", " + ID + "," + title + "," + "Success");
                                            excelWriterScoringMatrixNew.Flush();
                                        }

                                    }
                                    catch (Exception ex)
                                    {

                                    }

                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }

                        }
                        else if (listURL.ToString().Contains("Discussions"))
                        {
                            try
                            {
                                #region VIEW for "Status" List

                                bool _dListExist = clientcontext.Web.Lists.Cast<List>().Any(xlist => string.Equals(xlist.Title, "Discussions"));

                                if (_dListExist)
                                {
                                    try
                                    {
                                        List oList = _Lists.GetByTitle("Discussions");
                                        clientcontext.Load(oList);
                                        clientcontext.ExecuteQuery();


                                        ListItem _Item = oList.GetItemById(ID);
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();

                                        string title = Convert.ToString(_Item["Title"]);

                                        _Item.DeleteObject();
                                        clientcontext.ExecuteQuery();

                                        if (Archive == "A")
                                        {
                                            excelWriterScoringMatrixNew.WriteLine(spURl + "," + "Documents" + ", " + ID + "," + _Item["Title"].ToString() + "," + "Success");
                                            excelWriterScoringMatrixNew.Flush();
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                    }

                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");

            #endregion
        }
        private void button50_Click(object sender, EventArgs e)
        {
            #region AD Groups CSV Reading

            //lstADGroupsColl.Clear();

            //if (!string.IsNullOrEmpty(textBox3.Text))
            //{
            //    StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox3.Text));

            //    while (!sr.EndOfStream)
            //    {
            //        try
            //        {
            //            lstADGroupsColl.Add(sr.ReadLine().Trim().ToLower());
            //        }
            //        catch
            //        {
            //            continue;
            //        }
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please browse the path for ADGroups.csv");
            //}

            #endregion

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            //if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

                while (!sr.EndOfStream)
                {
                    try
                    {
                        lstSiteColl.Add(sr.ReadLine().Trim());
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            //else
            //{
            //    MessageBox.Show("Please browse the path for SiteColl.csv / Reports folder");
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;

            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteAssetsViewReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");

            excelWriterScoringMatrixNew.WriteLine("URL" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            //List<string> ListNames = new List<string>();
            //ListNames.Add("Site Assets");
            //ListNames.Add("2_Documents and Pages");
            //ListNames.Add("1_Uploaded Files");
            //ListNames.Add("Discussions");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(),
                                "sreekanth.grandhasila@sbdinc.com", "FN5O!CQa"))
                    {
                        _cContext.Load(_cContext.Web);
                        _cContext.ExecuteQuery();

                        _cContext.Web.QuickLaunchEnabled = false;
                        _cContext.Web.TreeViewEnabled = true;
                        _cContext.Web.Update();
                        _cContext.ExecuteQuery();

                        ListCollection _Lists = _cContext.Web.Lists;
                        _cContext.Load(_Lists);
                        _cContext.ExecuteQuery();



                        #region List Views



                        foreach (List list in _Lists)
                        {
                            _cContext.Load(list);
                            _cContext.ExecuteQuery();

                            try
                            {
                                if (list.Title != "Site Pages" && list.BaseTemplate == (int)ListTemplateType.DocumentLibrary && list.Hidden == false &&
                                    list.Title != "Site Assets")
                                {

                                    ViewCollection ViewColl = list.Views;
                                    _cContext.Load(ViewColl);
                                    _cContext.ExecuteQuery();

                                    Microsoft.SharePoint.Client.View v = ViewColl[0];
                                    _cContext.Load(v);
                                    _cContext.ExecuteQuery();

                                    v.ViewFields.RemoveAll();
                                    v.Update();
                                    _cContext.ExecuteQuery();

                                    v.ViewFields.Add("DocIcon");
                                    // v.ViewFields.Add("Title");
                                    v.ViewFields.Add("LinkFilename");
                                    v.ViewFields.Add("Modified");
                                    v.ViewFields.Add("Modified By");
                                    v.ViewFields.Add("Created By");
                                    v.ViewFields.Add("FileSizeDisplay");

                                    v.Update();
                                    _cContext.ExecuteQuery();

                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }

                        #endregion



                    }
                }
                catch (Exception ex)
                {
                    continue;
                }

            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button51_Click(object sender, EventArgs e)
        {
            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            StreamWriter excelWriterScoringMatrixNew1 = null;
            excelWriterScoringMatrixNew1 = System.IO.File.CreateText(textBox2.Text + "\\" + "ContentApprovalReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew1.WriteLine("URL" + "," + "Status");
            excelWriterScoringMatrixNew1.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            List Pagelist = _Lists.GetByTitle("2_Documents and Pages");
                            clientcontext.Load(Pagelist);
                            clientcontext.ExecuteQuery();

                            if (Pagelist.EnableModeration.ToString().ToLower() == "true")
                            {
                                excelWriterScoringMatrixNew1.WriteLine(lstSiteColl[j].ToString().Trim() + "," + "Success");
                                excelWriterScoringMatrixNew1.Flush();

                                Pagelist.EnableModeration = false;
                                Pagelist.Update();
                                clientcontext.ExecuteQuery();
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew1.Flush();
            excelWriterScoringMatrixNew1.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button52_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DefaultGroupsAsAssociateGroupsReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SPURL" + "," + "OwnerGroup" + "," + "MembersGroup" + "," + "VisitorsGroup");
            excelWriterScoringMatrixNew.Flush();

            StreamWriter excelWriterScoringMatrixNew1 = null;
            excelWriterScoringMatrixNew1 = System.IO.File.CreateText(textBox2.Text + "\\" + "NoGroupsReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew1.WriteLine("SPURL");
            excelWriterScoringMatrixNew1.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();

                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        Web owebt = clientcontext.Web;
                        clientcontext.Load(owebt);
                        clientcontext.ExecuteQuery();
                        clientcontext.Load(owebt, oweb => oweb.Title, oweb => oweb.HasUniqueRoleAssignments);//, clientcontext.Web.Title, clientcontext.Web.HasUniqueRoleAssignments);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            if (owebt.HasUniqueRoleAssignments)
                            {
                                string OwnerGroup = "TBD";
                                string MembersGroup = "TBD";
                                string VisitorsGroup = "TBD";

                                Microsoft.SharePoint.Client.GroupCollection AllGroups = clientcontext.Web.RoleAssignments.Groups;
                                clientcontext.Load(AllGroups);
                                clientcontext.ExecuteQuery();

                                if (AllGroups.Count < 1)
                                {
                                    excelWriterScoringMatrixNew1.WriteLine(lstSiteColl[j].ToString());
                                    excelWriterScoringMatrixNew1.Flush();
                                }
                                else
                                {

                                    Group ownergrp2 = null;

                                    foreach (Microsoft.SharePoint.Client.Group grp in AllGroups)
                                    {
                                        try
                                        {
                                            clientcontext.Load(grp);
                                            clientcontext.ExecuteQuery();

                                            if (grp.Title.EndsWith("Owners"))
                                            {
                                                try
                                                {
                                                    owebt.AssociatedOwnerGroup = grp;
                                                    owebt.AssociatedOwnerGroup.Update();
                                                    grp.Update();
                                                    owebt.Update();
                                                    clientcontext.ExecuteQuery();
                                                    OwnerGroup = "Success";

                                                    ownergrp2 = grp;
                                                    clientcontext.Load(ownergrp2);
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                    OwnerGroup = "Fail";
                                                }

                                                try
                                                {
                                                    RoleDefinition rd = clientcontext.Web.RoleDefinitions.GetByName("Site Admin");
                                                    RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(clientcontext);
                                                    rdb.Add(rd);
                                                    clientcontext.Web.RoleAssignments.Add(grp, rdb);
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                }

                                                try
                                                {
                                                    clientcontext.Web.RemovePermissionLevelFromGroup(grp.Title, "Full Control", false);
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                }
                                            }
                                            if (grp.Title.EndsWith("Members"))
                                            {
                                                try
                                                {
                                                    owebt.AssociatedMemberGroup = grp;
                                                    owebt.AssociatedMemberGroup.Update();
                                                    grp.Update();
                                                    owebt.Update();
                                                    clientcontext.ExecuteQuery();
                                                    MembersGroup = "Success";
                                                }
                                                catch (Exception ex)
                                                {
                                                    MembersGroup = "Fail";
                                                }
                                                try
                                                {
                                                    clientcontext.Web.RemovePermissionLevelFromGroup(grp.Title, "Edit", false);
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                }
                                            }
                                            if (grp.Title.EndsWith("Visitors"))
                                            {
                                                try
                                                {
                                                    owebt.AssociatedVisitorGroup = grp;
                                                    owebt.AssociatedVisitorGroup.Update();
                                                    grp.Update();
                                                    owebt.Update();
                                                    clientcontext.ExecuteQuery();
                                                    VisitorsGroup = "Success";
                                                }
                                                catch (Exception ex)
                                                {
                                                    VisitorsGroup = "Fail";
                                                }
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + grp.Title + " : GroupIssue" + "," + "NA" + "," + "NA");
                                            excelWriterScoringMatrixNew.Flush();
                                            continue;
                                        }
                                    }

                                    if (ownergrp2 != null)
                                    {
                                        foreach (Microsoft.SharePoint.Client.Group grp in AllGroups)
                                        {
                                            try
                                            {
                                                clientcontext.Load(grp);
                                                clientcontext.ExecuteQuery();

                                                try
                                                {
                                                    grp.Owner = ownergrp2;
                                                    grp.Update();
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception exc)
                                                {

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + grp.Title + " : OwnerGroupAssignIssue" + "," + "NA" + "," + "NA");
                                                excelWriterScoringMatrixNew.Flush();
                                                continue;
                                            }
                                        }
                                    }

                                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + OwnerGroup + "," + MembersGroup + "," + VisitorsGroup);
                                    excelWriterScoringMatrixNew.Flush();
                                }
                            }
                            else
                            {
                                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "InheritedPermissions" + "," + "" + "," + "");
                                excelWriterScoringMatrixNew.Flush();
                            }
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Failure" + "," + "" + "," + "");
                            excelWriterScoringMatrixNew.Flush();
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            excelWriterScoringMatrixNew1.Flush();
            excelWriterScoringMatrixNew1.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button53_Click(object sender, EventArgs e)
        {
            #region ImportedItems CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("ObjectID", typeof(string)), new DataColumn("ObjectType", typeof(string)), new DataColumn("Imported URL", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);
            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            DataTable dtJiveObjects = new DataTable();
            dtJiveObjects.Columns.AddRange(new DataColumn[6] { new DataColumn("ObjectId", typeof(string)), new DataColumn("ObjectType", typeof(string)), new DataColumn("Created Date", typeof(string)), new DataColumn("CreatedBy", typeof(string)), new DataColumn("Modified Date", typeof(string)), new DataColumn("ModifiedBy", typeof(string)) });

            string csvData1 = System.IO.File.ReadAllText(textBox3.Text);
            foreach (string row in csvData1.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtJiveObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region SBD User Mapping Reading

            DataTable UserData = new DataTable();
            List<string> _MappedUsersList = new List<string>();

            UserData = getcsvmetadatainfo(textBox4.Text);

            for (int Rscount = 1; Rscount <= UserData.Rows.Count - 1; Rscount++)
            {
                if (!string.IsNullOrEmpty(UserData.Rows[Rscount][1].ToString()))
                {
                    if (!_MappedUsersList.Contains(UserData.Rows[Rscount][1].ToString()))
                    {
                        _MappedUsersList.Add(UserData.Rows[Rscount][1].ToString());
                    }
                }
            }

            string[] _MappedUserCollection = _MappedUsersList.ToArray();

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SBDItemsModfiedReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("ObjectID" + "," + "ObjectType" + "," + "URL" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;
            string[] ListSplit = new string[] { "/Lists/" };
            string[] IDSplit = new string[] { "/DispForm.aspx?ID=" };
            string[] DocumentsSplit = new string[] { "/Shared Documents/" };
            string[] SitesSplit = new string[] { "/sites/" };
            string[] slashSplit = new string[] { "/" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string Modified = string.Empty;
                    string ModifiedBy = string.Empty;
                    string Created = string.Empty;
                    string CreatedBy = string.Empty;
                    string _SiteTitle = string.Empty;

                    if (drImported["Imported URL"].ToString().Trim().ToLower().Contains("/lists/"))
                    {
                        _SiteTitle = drImported["Imported URL"].ToString().Split(ListSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }
                    if (drImported["Imported URL"].ToString().Trim().ToLower().Contains("/shared documents/"))
                    {
                        _SiteTitle = drImported["Imported URL"].ToString().Split(DocumentsSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }

                    this.Text = (count).ToString() + " : " + _SiteTitle;
                    count++;

                    bool itemFound = false;

                    foreach (DataRow drJive in dtJiveObjects.Rows)
                    {
                        if ((drImported["ObjectID"].ToString().Trim() == drJive["ObjectId"].ToString().Trim()) && (drImported["ObjectType"].ToString().Trim().ToLower() == drJive["ObjectType"].ToString().Trim().ToLower()))
                        {
                            Modified = drJive["Modified Date"].ToString().Trim();
                            ModifiedBy = drJive["ModifiedBy"].ToString().Trim();
                            Created = drJive["Created Date"].ToString().Trim();
                            CreatedBy = drJive["CreatedBy"].ToString().Trim();

                            itemFound = true;
                            break;
                        }
                    }

                    if (itemFound)
                    {
                        string _FilePath = string.Empty;
                        string _ActualFilePath = string.Empty;
                        string _ListName = string.Empty;
                        string _ItemID = string.Empty;

                        AuthenticationManager authManager = new AuthenticationManager();
                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_SiteTitle, "sreekanth.grandhasila@sbdinc.com", "FN5O!CQa"))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();

                            #region Web.EnsureUser TEST

                            //            ////Greg.Keier@sbdinc.com
                            //            //User ModiUser = default(User);
                            //            //try
                            //            //{
                            //            //    ModiUser = clientcontext.Web.EnsureUser("Greg.Keier@sbdinc.com");
                            //            //    clientcontext.Load(ModiUser);
                            //            //    clientcontext.ExecuteQuery();
                            //            //}
                            //            //catch (Exception ex)
                            //            //{
                            //            //    ModiUser = clientcontext.Web.EnsureUser("sreekanth.grandhasila@sbdinc.com");
                            //            //    clientcontext.Load(ModiUser);
                            //            //    clientcontext.ExecuteQuery();
                            //            //} 

                            #endregion

                            switch (drImported["ObjectType"].ToString().ToLower().Trim())
                            {
                                case "folder":
                                    if (!drImported["Imported URL"].ToString().Trim().ToLower().EndsWith(".aspx"))
                                    {
                                        _FilePath = drImported["Imported URL"].ToString().Trim().Split(SitesSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                        _ActualFilePath = "/sites/" + _FilePath;
                                    }
                                    break;

                                case "otherfile":
                                    _FilePath = drImported["Imported URL"].ToString().Trim().Split(SitesSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    _ActualFilePath = "/sites/" + _FilePath;
                                    break;

                                case "row":
                                    _ItemID = drImported["Imported URL"].ToString().Trim().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    string lstName = drImported["Imported URL"].ToString().Trim().Split(ListSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    _ListName = lstName.Split(slashSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                                    break;

                                    //case "link":
                                    //    _FilePath = drImported["Imported URL"].ToString().Split(SitesSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    //    _ActualFilePath = "/sites/" + _FilePath;
                                    //    break;

                                    //case "otherfileversion":
                                    //    _FilePath = drImported["Imported URL"].ToString().Split(SitesSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                    //    _ActualFilePath = "/sites/" + _FilePath;
                                    //    break;
                            }

                            if (!string.IsNullOrEmpty(_ActualFilePath))
                            {
                                try
                                {
                                    ListItem _Item = oWeb.GetFileByUrl(_ActualFilePath).ListItemAllFields;
                                    clientcontext.Load(_Item);
                                    clientcontext.ExecuteQuery();

                                    List _List = _Item.ParentList;
                                    clientcontext.Load(_List);
                                    clientcontext.ExecuteQuery();

                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    try
                                    {
                                        User CreatedUser = default(User);
                                        try
                                        {
                                            CreatedUser = oWeb.EnsureUser(GetUserLoginName(_MappedUserCollection, CreatedBy));
                                            clientcontext.Load(CreatedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            CreatedUser = oWeb.EnsureUser("sreekanth.grandhasila@sbdinc.com");
                                            clientcontext.Load(CreatedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        FieldUserValue CreatedUserValue = new FieldUserValue();
                                        CreatedUserValue.LookupId = CreatedUser.Id;
                                        DateTime dtCreated = getdateformat(Created);

                                        User ModifiedUser = default(User);
                                        try
                                        {
                                            ModifiedUser = oWeb.EnsureUser(GetUserLoginName(_MappedUserCollection, ModifiedBy));
                                            clientcontext.Load(ModifiedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            ModifiedUser = oWeb.EnsureUser("sreekanth.grandhasila@sbdinc.com");
                                            clientcontext.Load(ModifiedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        FieldUserValue ModifiedUserValue = new FieldUserValue();
                                        ModifiedUserValue.LookupId = ModifiedUser.Id;
                                        //DateTime dtModified = getdateformat(Modified);
                                        DateTime dtModified = Convert.ToDateTime(_Item["Modified"]);

                                        //_Item["Created"] = dtCreated;
                                        _Item["Author"] = CreatedUserValue;
                                        _Item["Modified"] = dtModified;
                                        _Item["Editor"] = ModifiedUserValue;
                                        _Item.Update();
                                        clientcontext.ExecuteQuery();

                                        excelWriterScoringMatrixNew.WriteLine(drImported["ObjectID"].ToString().Trim() + "," + drImported["ObjectType"].ToString().Trim() + "," + drImported["Imported URL"].ToString().Trim() + "," + "Success");
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                    catch (Exception ex)
                                    {

                                    }

                                    _List.EnableVersioning = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();
                                }
                                catch (Exception ex)
                                {

                                }
                            }

                            if (!string.IsNullOrEmpty(_ItemID) && !string.IsNullOrEmpty(_ListName))
                            {
                                try
                                {
                                    List _List = clientcontext.Web.Lists.GetByTitle(_ListName);
                                    clientcontext.Load(_List);
                                    clientcontext.ExecuteQuery();

                                    try
                                    {
                                        ListItem _Item = _List.GetItemById(_ItemID);
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();

                                        User CreatedUser = default(User);
                                        try
                                        {
                                            CreatedUser = oWeb.EnsureUser(GetUserLoginName(_MappedUserCollection, CreatedBy));
                                            clientcontext.Load(CreatedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            CreatedUser = oWeb.EnsureUser("sreekanth.grandhasila@sbdinc.com");
                                            clientcontext.Load(CreatedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        FieldUserValue CreatedUserValue = new FieldUserValue();
                                        CreatedUserValue.LookupId = CreatedUser.Id;
                                        DateTime dtCreated = getdateformat(Created);

                                        User ModifiedUser = default(User);
                                        try
                                        {
                                            ModifiedUser = oWeb.EnsureUser(GetUserLoginName(_MappedUserCollection, ModifiedBy));
                                            clientcontext.Load(ModifiedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            ModifiedUser = oWeb.EnsureUser("sreekanth.grandhasila@sbdinc.com");
                                            clientcontext.Load(ModifiedUser);
                                            clientcontext.ExecuteQuery();
                                        }
                                        FieldUserValue ModifiedUserValue = new FieldUserValue();
                                        ModifiedUserValue.LookupId = ModifiedUser.Id;
                                        //DateTime dtModified = getdateformat(Modified);
                                        DateTime dtModified = Convert.ToDateTime(_Item["Modified"]);

                                        //_Item["Created"] = dtCreated;
                                        _Item["Author"] = CreatedUserValue;
                                        _Item["Modified"] = dtModified;
                                        _Item["Editor"] = ModifiedUserValue;
                                        _Item.Update();
                                        clientcontext.ExecuteQuery();

                                        excelWriterScoringMatrixNew.WriteLine(drImported["ObjectID"].ToString().Trim() + "," + drImported["ObjectType"].ToString().Trim() + "," + drImported["Imported URL"].ToString().Trim() + "," + "Success");
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                        }
                    }
                    else
                    {
                        excelWriterScoringMatrixNew.WriteLine(drImported["ObjectID"].ToString().Trim() + "," + drImported["ObjectType"].ToString().Trim() + "," + drImported["Imported URL"].ToString().Trim() + "," + "ItemIDNotFound");
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(drImported["ObjectID"].ToString().Trim() + "," + drImported["ObjectType"].ToString().Trim() + "," + drImported["Imported URL"].ToString().Trim() + "," + "Failure due to : " + ex.Message);
                    excelWriterScoringMatrixNew.Flush();

                    continue;
                }
            }
            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
        private DataTable getcsvmetadatainfo(string path)
        {
            DataTable dt = new DataTable();
            ArrayList sContents = new ArrayList();
            // dt = null;
            try
            {
                sContents = GetcsvTextContents(path);
                DataRow row1 = null;
                int rowcount = sContents.Count;
                //Getting Header
                string header_content = sContents[0].ToString();
                string[] dtheader = header_content.Split(',');
                for (int k = 0; k <= dtheader.Length - 1; k++)
                {
                    dt.Columns.Add(dtheader[k].ToString());
                }
                //Getting Rows
                for (int i = 1; i <= rowcount - 1; i++)
                {
                    try
                    {
                        string row_content = sContents[i].ToString();
                        string[] dtrow = row_content.Split(',');
                        row1 = dt.NewRow();
                        for (int j = 0; j <= dtrow.Length - 1; j++)
                        {
                            row1[j] = dtrow[j].ToString();
                        }
                        dt.Rows.Add(row1);
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }
                }
                sContents = null;
            }
            catch (Exception ex)
            {

            }
            return dt;
        }
        private ArrayList GetcsvTextContents(string path)
        {
            ArrayList return_txtrows = new ArrayList();
            try
            {
                using (StreamReader objReader = new StreamReader(path))
                {
                    while (!((objReader.EndOfStream)))
                    {
                        return_txtrows.Add(objReader.ReadLine().ToString());
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return return_txtrows;
        }
        public static string GetUserLoginName(IList<string> _MappedUsercollection, string eRoomUserLoginName)
        {
            try
            {
                for (int Mp = 0; Mp <= _MappedUsercollection.Count - 1; Mp++)
                {
                    if (eRoomUserLoginName == System.Text.RegularExpressions.Regex.Split(_MappedUsercollection[Mp].ToString(), "-->")[0].ToString())
                    {
                        eRoomUserLoginName = System.Text.RegularExpressions.Regex.Split(_MappedUsercollection[Mp].ToString(), "-->")[1].ToString();
                        break;
                    }
                }
                return eRoomUserLoginName;
            }
            catch (Exception ex)
            {
                return eRoomUserLoginName;
            }
        }
        private void button54_Click(object sender, EventArgs e)
        {
            if (openFileDialog3.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = openFileDialog3.FileName;
            }
        }
        private void button55_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteHistoryTitleFixReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ObjectID" + "," + "Title");
            excelWriterScoringMatrixNew.Flush();

            List<string> ListNames = new List<string>();
            ListNames.Add("2_Documents and Pages");
            ListNames.Add("Site Assets");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        foreach (string ls in ListNames)
                        {

                            try
                            {
                                List oList = oLists.GetByTitle(ls);
                                clientcontext.Load(oList);
                                clientcontext.ExecuteQuery();

                                try
                                {
                                    oList.EnableVersioning = false;
                                    oList.Update();
                                    clientcontext.ExecuteQuery();
                                }
                                catch (Exception ex)
                                {

                                }

                                CamlQuery camlQuery = new CamlQuery();
                                camlQuery.ViewXml = "<View Scope=\"RecursiveAll\"><RowLimit>500</RowLimit></View>";

                                ListItemCollection listItems = oList.GetItems(camlQuery);
                                clientcontext.Load(listItems);
                                clientcontext.ExecuteQuery();

                                foreach (ListItem _Item in listItems)
                                {
                                    clientcontext.Load(_Item);
                                    clientcontext.ExecuteQuery();

                                    try
                                    {
                                        DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                        FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];
                                        string _ItemName = _Item["FileLeafRef"].ToString();

                                        _Item["Title"] = _ItemName;
                                        _Item["Modified"] = Modified;
                                        _Item["Editor"] = ModifiedBy;
                                        _Item.Update();
                                        clientcontext.ExecuteQuery();

                                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + _Item.Id.ToString() + "," + _ItemName);
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                    catch (Exception ex)
                                    {
                                        excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Error" + "," + ex.Message);
                                        excelWriterScoringMatrixNew.Flush();

                                        continue;
                                    }
                                }

                                try
                                {
                                    oList.EnableVersioning = true;
                                    oList.Update();
                                    clientcontext.ExecuteQuery();
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }
        private void button56_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));
            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "UploadFilesCTReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "RicohCT" + "," + "DocumentCT");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            List oList = oLists.GetByTitle("1_Uploaded Files");
                            clientcontext.Load(oList);
                            clientcontext.ExecuteQuery();

                            string isRicohCTExist = "No";
                            string isDocumentCTExist = "No";

                            ContentTypeCollection _ViewColl = oList.ContentTypes;
                            clientcontext.Load(_ViewColl);
                            clientcontext.ExecuteQuery();

                            foreach (ContentType _CTItem in _ViewColl)
                            {
                                try
                                {
                                    clientcontext.Load(_CTItem);
                                    clientcontext.ExecuteQuery();

                                    if (_CTItem.Name == "RicohContentType")
                                    {
                                        isRicohCTExist = "Yes";
                                    }

                                    if (_CTItem.Name == "Document")
                                    {
                                        isDocumentCTExist = "Yes";
                                    }

                                }
                                catch (Exception ex)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "Error" + "," + ex.Message);
                                    excelWriterScoringMatrixNew.Flush();

                                    continue;
                                }
                            }

                            excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + isRicohCTExist + "," + isDocumentCTExist);
                            excelWriterScoringMatrixNew.Flush();

                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }

        private void button57_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig            

            StreamWriter excelWriterTagListCreation = null;
            excelWriterTagListCreation = System.IO.File.CreateText(textBox2.Text + "\\" + "ManageTagListCreationReport" + ".csv");
            excelWriterTagListCreation.WriteLine("SiteURL" + "," + "Status" + "," + "Details");
            excelWriterTagListCreation.Flush();

            StreamWriter excelWriterTagColumnCreation = null;
            excelWriterTagColumnCreation = System.IO.File.CreateText(textBox2.Text + "\\" + "TagColumnCreationErrorReport" + ".csv");
            excelWriterTagColumnCreation.WriteLine("SiteURL" + "," + "ListName" + "," + "Details");
            excelWriterTagColumnCreation.Flush();

            List<string> ListNames = new List<string>();

            //ListNames.Add("Team Files");
            ListNames.Add("1_Uploaded Files");
            ListNames.Add("2_Documents and Pages");
            ListNames.Add("Discussions");
            ListNames.Add("Events");
            ListNames.Add("Messages");
            ListNames.Add("Posts");
            //ListNames.Add("SiteHistory");
            ListNames.Add("Tasks");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web _Web = _cContext.Web;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        bool ManageTagtagsListExist = _Web.ListExists("Manage Tag");

                        if (!ManageTagtagsListExist)
                        {
                            try
                            {
                                ListCreationInformation creationInfo = new ListCreationInformation();
                                creationInfo.Title = "Manage Tag";
                                creationInfo.Description = "Manage Tag";
                                creationInfo.TemplateType = (int)ListTemplateType.GenericList;
                                List newList = _cContext.Web.Lists.Add(creationInfo);
                                _cContext.Load(newList);
                                _cContext.ExecuteQuery();

                                try
                                {
                                    List newList1 = _cContext.Web.Lists.GetByTitle("Manage Tag");
                                    _cContext.Load(newList1);
                                    _cContext.ExecuteQuery();

                                    FieldCollection oFieldColl = newList1.Fields;
                                    _cContext.Load(oFieldColl);
                                    _cContext.ExecuteQuery();

                                    Field field = oFieldColl.GetByTitle("Title");
                                    _cContext.Load(field);
                                    _cContext.ExecuteQuery();

                                    field.Indexed = true;
                                    field.Update();
                                    _cContext.ExecuteQuery();

                                    field.EnforceUniqueValues = true;
                                    field.Update();
                                    _cContext.ExecuteQuery();
                                }
                                catch (Exception ec1)
                                {

                                }

                                excelWriterTagListCreation.WriteLine(_cContext.Web.Url + "," + "Success" + "," + "NA");
                                excelWriterTagListCreation.Flush();
                            }
                            catch (Exception ex)
                            {
                                excelWriterTagListCreation.WriteLine(_cContext.Web.Url + "," + "ManageTag List Creation Failure" + "," + ex.Message.Replace(",", ""));
                                excelWriterTagListCreation.Flush();
                            }
                        }

                        foreach (string ls in ListNames)
                        {
                            try
                            {
                                List _List = null;
                                _List = _cContext.Web.Lists.GetByTitle(ls);
                                _cContext.Load(_List);
                                _cContext.ExecuteQuery();

                                bool tagsFileldExist = _List.FieldExistsByName("Tag");

                                if (!tagsFileldExist)
                                {
                                    try
                                    {
                                        List list = _Web.Lists.GetByTitle("Manage Tag");
                                        _cContext.Load(list);
                                        _cContext.ExecuteQuery();
                                        string schemaLookupField = "<Field Type='LookupMulti' Name='Tag' StaticName='Tag' DisplayName='Tag' List = '" + list.Id + "' ShowField = 'Title' Mult = 'TRUE'/>";
                                        Field lookupField = _List.Fields.AddFieldAsXml(schemaLookupField, false, AddFieldOptions.AddFieldInternalNameHint);
                                        _List.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        excelWriterTagColumnCreation.WriteLine(_cContext.Web.Url + "," + ls + "," + ex.Message.Replace(",", ""));
                                        excelWriterTagColumnCreation.Flush();

                                        continue;
                                    }
                                }

                                #region Views

                                //if (ls == "1_Uploaded Files")
                                //{
                                //    ViewCollection ViewColl = _List.Views;
                                //    _cContext.Load(ViewColl);
                                //    _cContext.ExecuteQuery();

                                //    Microsoft.SharePoint.Client.View v = ViewColl[0];
                                //    _cContext.Load(v);
                                //    _cContext.ExecuteQuery();

                                //    v.ViewFields.RemoveAll();
                                //    v.Update();
                                //    _cContext.ExecuteQuery();

                                //    v.ViewFields.Add("DocIcon");
                                //    v.ViewFields.Add("Title");
                                //    v.ViewFields.Add("LinkFilename");
                                //    v.ViewFields.Add("Created");
                                //    v.ViewFields.Add("Created By");
                                //    v.ViewFields.Add("Modified");
                                //    v.ViewFields.Add("Modified By");
                                //    v.ViewFields.Add("Tags");
                                //    v.ViewFields.Add("Categorization");
                                //    v.Update();
                                //    _cContext.ExecuteQuery();
                                //}

                                //if (ls == "2_Documents and Pages")
                                //{
                                //    try
                                //    {
                                //        ViewCollection ViewColl = _List.Views;
                                //        _cContext.Load(ViewColl);
                                //        _cContext.ExecuteQuery();

                                //        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                //        _cContext.Load(v);
                                //        _cContext.ExecuteQuery();

                                //        v.ViewFields.RemoveAll();
                                //        v.Update();
                                //        _cContext.ExecuteQuery();

                                //        v.ViewFields.Add("DocIcon");
                                //        v.ViewFields.Add("Title");
                                //        v.ViewFields.Add("LinkFilename");
                                //        v.ViewFields.Add("Created");
                                //        v.ViewFields.Add("Created By");
                                //        v.ViewFields.Add("Modified");
                                //        v.ViewFields.Add("Modified By");
                                //        v.ViewFields.Add("Tag");
                                //        v.ViewFields.Add("Categorization");
                                //        v.ViewFields.Add("CheckoutUser");
                                //        v.Update();
                                //        _cContext.ExecuteQuery();
                                //    }
                                //    catch (Exception ex)
                                //    { }
                                //}

                                //if (ls == "Posts")
                                //{
                                //    try
                                //    {
                                //        ViewCollection ViewColl = _List.Views;
                                //        _cContext.Load(ViewColl);
                                //        _cContext.ExecuteQuery();

                                //        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                //        _cContext.Load(v);
                                //        _cContext.ExecuteQuery();

                                //        v.ViewFields.RemoveAll();
                                //        v.Update();
                                //        _cContext.ExecuteQuery();

                                //        v.ViewFields.Add("LinkTitle");
                                //        v.ViewFields.Add("Created");
                                //        v.ViewFields.Add("Published");
                                //        v.ViewFields.Add("Category");
                                //        v.ViewFields.Add("NumComments");
                                //        v.ViewFields.Add("Edit");
                                //        v.ViewFields.Add("Categorization");
                                //        v.ViewFields.Add("LikesCount");
                                //        v.Update();
                                //        _cContext.ExecuteQuery();
                                //    }
                                //    catch (Exception ex)
                                //    { }
                                //} 

                                #endregion

                                #region Site Assets View

                                //if (ls == "Site Assets")
                                //{
                                //    try
                                //    {
                                //        ViewCollection ViewColl = _List.Views;
                                //        _cContext.Load(ViewColl);
                                //        _cContext.ExecuteQuery();

                                //        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                //        _cContext.Load(v);
                                //        _cContext.ExecuteQuery();

                                //        v.ViewFields.Add("DocIcon");
                                //        v.ViewFields.Add("Title");
                                //        v.ViewFields.Add("LinkFilename");
                                //        v.ViewFields.Add("Created");
                                //        v.ViewFields.Add("Created By");
                                //        v.ViewFields.Add("Modified");
                                //        v.ViewFields.Add("Modified By");
                                //        v.ViewFields.Add("CheckoutUser");
                                //        v.Update();
                                //        _cContext.ExecuteQuery();
                                //    }
                                //    catch (Exception ex)
                                //    { }
                                //} 

                                #endregion
                            }
                            catch (Exception ex)
                            {
                                excelWriterTagColumnCreation.WriteLine(_cContext.Web.Url + "," + ls + "," + "ListNotExist");
                                excelWriterTagColumnCreation.Flush();

                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterTagListCreation.WriteLine(lstSiteColl[j].ToString() + "," + "SiteNotAccessed" + "," + ex.Message.Replace(",", ""));
                    excelWriterTagListCreation.Flush();

                    continue;
                }
            }

            #endregion                       

            excelWriterTagColumnCreation.Flush();
            excelWriterTagColumnCreation.Close();

            excelWriterTagListCreation.Flush();
            excelWriterTagListCreation.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button58_Click(object sender, EventArgs e)
        {
            #region ImportedObjects CSV Reading

            DataTable dtImportedGUID = new DataTable();
            dtImportedGUID.Columns.AddRange(new DataColumn[2] { new DataColumn("SiteURL", typeof(string)), new DataColumn("GUID", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedGUID.Rows.Add();
                    int i = 0;
                    foreach (string cell in row.Split(','))
                    {
                        dtImportedGUID.Rows[dtImportedGUID.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            int count = 0;

            foreach (DataRow drImported in dtImportedGUID.Rows)
            {
                try
                {
                    string _SiteTitle = drImported["SiteURL"].ToString().Trim();
                    string _SiteGUID = drImported["GUID"].ToString().Trim();

                    FileInfo _CsvFilepath = null;

                    try
                    {
                        _CsvFilepath = new DirectoryInfo(textBox2.Text).GetFiles(_SiteGUID + "_UniqueTagsReport.csv", SearchOption.AllDirectories)[0];
                    }
                    catch
                    { }

                    if (_CsvFilepath != null)
                    {
                        #region ImportedObjects CSV Reading

                        DataTable dtImportedObjects = new DataTable();
                        dtImportedObjects.Columns.AddRange(new DataColumn[2] { new DataColumn("SiteURL", typeof(string)), new DataColumn("Tags", typeof(string)) });

                        //Read the contents of CSV file.  
                        string csvData1 = System.IO.File.ReadAllText(_CsvFilepath.FullName);

                        //Execute a loop over the rows.  
                        foreach (string row in csvData1.Split('\n'))
                        {
                            if (!string.IsNullOrEmpty(row))
                            {
                                dtImportedObjects.Rows.Add();
                                int i = 0;
                                //Execute a loop over the columns.  
                                foreach (string cell in row.Split(','))
                                {
                                    dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                                    i++;
                                }
                            }
                        }
                        #endregion

                        foreach (DataRow drGUID in dtImportedObjects.Rows)
                        {
                            try
                            {
                                string _TagsColl = drImported["Tags"].ToString().Trim();

                                if (!string.IsNullOrEmpty(_TagsColl))
                                {
                                    string[] tags = _TagsColl.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                                    this.Text = (count).ToString() + " : " + _SiteTitle;
                                    count++;

                                    AuthenticationManager authManager = new AuthenticationManager();
                                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_SiteTitle, "svc-jivemigration7@rsharepoint.onmicrosoft.com", "Nuq92882"))
                                    {
                                        Web _Web = _cContext.Web;
                                        _cContext.Load(_Web);
                                        _cContext.ExecuteQuery();

                                        bool ManageTagtagsListExist = _Web.ListExists("Manage Tag");

                                        if (!ManageTagtagsListExist)
                                        {
                                            List list = _Web.Lists.GetByTitle("Manage Tag");
                                            _cContext.Load(list);
                                            _cContext.ExecuteQuery();

                                            foreach (string ostr in tags)
                                            {
                                                try
                                                {
                                                    int _cId = GetLookupIDsManageTag(ostr, _cContext, _Web);

                                                    #region OLD

                                                    //CamlQuery cq = new CamlQuery();
                                                    //cq.ViewXml = "<View><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + ostr + "</Value></Eq></Where></View>";

                                                    //ListItemCollection oItems = list.GetItems(cq);
                                                    //_cContext.Load(oItems);
                                                    //_cContext.ExecuteQuery();

                                                    //if (oItems.Count == 0)
                                                    //{
                                                    //    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                                    //    ListItem oListItem = list.AddItem(itemCreateInfo);
                                                    //    oListItem["Title"] = ostr;
                                                    //    oListItem.Update();
                                                    //    _cContext.ExecuteQuery();
                                                    //} 

                                                    #endregion
                                                }
                                                catch (Exception ex)
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }


        }

        private void button59_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();

            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocumentListViewReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("URL" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();
                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString().Trim(), "sreekanth.grandhasila@sbdinc.com", "verinon@2018"))
                    {
                        clientcontext.Load(clientcontext.Web);
                        clientcontext.ExecuteQuery();

                        ListCollection _Lists = clientcontext.Web.Lists;
                        clientcontext.Load(_Lists);
                        clientcontext.ExecuteQuery();

                        try
                        {
                            List Pagelist = _Lists.GetByTitle("Documents");
                            clientcontext.Load(Pagelist);
                            clientcontext.ExecuteQuery();

                            ViewCollection ViewColl = Pagelist.Views;
                            clientcontext.Load(ViewColl);
                            clientcontext.ExecuteQuery();

                            Microsoft.SharePoint.Client.View v = ViewColl[0];
                            clientcontext.Load(v);
                            clientcontext.ExecuteQuery();

                            v.ViewFields.RemoveAll();
                            v.Update();
                            clientcontext.ExecuteQuery();

                            v.ViewFields.Add("DocIcon");
                            v.ViewFields.Add("LinkFilename");
                            v.ViewFields.Add("Modified");
                            v.ViewFields.Add("Modified By");
                            v.ViewFields.Add("Created By");
                            v.ViewFields.Add("File Size");
                            v.ViewFields.Add("Item Child Count");
                            v.Update();
                            clientcontext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }

        private void button60_Click(object sender, EventArgs e)
        {

            #region ImportedGUID CSV Reading

            DataTable dtImportedGUID = new DataTable();
            dtImportedGUID.Columns.AddRange(new DataColumn[2] { new DataColumn("SiteURL", typeof(string)), new DataColumn("GUID", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedGUID.Rows.Add();
                    int i = 0;
                    foreach (string cell in row.Split(','))
                    {
                        dtImportedGUID.Rows[dtImportedGUID.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            int count = 0;

            foreach (DataRow drImported in dtImportedGUID.Rows)
            {
                try
                {
                    string _SiteTitle = drImported["SiteURL"].ToString().Trim();
                    string _SiteGUID = drImported["GUID"].ToString().Trim();

                    FileInfo _CsvFilepath = null;

                    try
                    {
                        _CsvFilepath = new DirectoryInfo(textBox2.Text).GetFiles(_SiteGUID + "_TagsReport.csv", SearchOption.AllDirectories)[0];
                    }
                    catch
                    { }

                    if (_CsvFilepath != null)
                    {
                        #region ImportedObjects CSV Reading

                        DataTable dtImportedObjects = new DataTable();
                        dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("SiteURL", typeof(string)), new DataColumn("ListName", typeof(string)), new DataColumn("ItemID", typeof(string)), new DataColumn("Tags", typeof(string)) });

                        string csvData1 = System.IO.File.ReadAllText(_CsvFilepath.FullName);

                        foreach (string row in csvData1.Split('\n'))
                        {
                            if (!string.IsNullOrEmpty(row))
                            {
                                dtImportedObjects.Rows.Add();
                                int i = 0;

                                foreach (string cell in row.Split(','))
                                {
                                    dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                                    i++;
                                }
                            }
                        }
                        #endregion

                        StreamWriter excelWriterTagListCreation = null;
                        excelWriterTagListCreation = System.IO.File.CreateText(textBox2.Text + "\\" + _SiteGUID + "_TagReApplyReport" + ".csv");
                        excelWriterTagListCreation.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "Status");
                        excelWriterTagListCreation.Flush();

                        foreach (DataRow drGUID in dtImportedObjects.Rows)
                        {
                            try
                            {
                                string _SiteURL = drImported["SiteURL"].ToString().Trim();
                                string _ListName = drImported["ListName"].ToString().Trim();
                                string _ItemID = drImported["ItemID"].ToString().Trim();
                                string _TagsColl = drImported["Tags"].ToString().Trim();

                                if (!string.IsNullOrEmpty(_TagsColl))
                                {
                                    string[] tags = _TagsColl.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                                    this.Text = (count).ToString() + " : " + _SiteTitle;
                                    count++;

                                    AuthenticationManager authManager = new AuthenticationManager();
                                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_SiteURL, "svc-jivemigration7@rsharepoint.onmicrosoft.com", "Nuq92882"))
                                    {
                                        Web _Web = _cContext.Web;
                                        _cContext.Load(_Web);
                                        _cContext.ExecuteQuery();

                                        bool targetListExists = _Web.ListExists(_ListName);

                                        if (targetListExists)
                                        {
                                            List targetList = _Web.Lists.GetByTitle(_ListName);
                                            _cContext.Load(targetList);
                                            _cContext.ExecuteQuery();

                                            targetList.EnableVersioning = false;
                                            targetList.Update();
                                            _cContext.ExecuteQuery();

                                            bool tagsFileldExist = targetList.FieldExistsByName("Tag");

                                            if (!tagsFileldExist)
                                            {
                                                try
                                                {
                                                    ListItem oItem = targetList.GetItemById(_ItemID);
                                                    _cContext.Load(oItem);
                                                    _cContext.ExecuteQuery();

                                                    DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                                                    FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                                                    if (!string.IsNullOrEmpty(_TagsColl.Trim()))
                                                    {
                                                        string[] _ddTags = _TagsColl.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                                        try
                                                        {

                                                            FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_ddTags.Length];

                                                            for (int i = 0; i <= _ddTags.Length - 1; i++)
                                                            {
                                                                string newValue = _ddTags[i].ToString();

                                                                if (_ddTags[i].ToString().Contains("$"))
                                                                {
                                                                    newValue = _ddTags[i].ToString().Replace("$", ",");
                                                                }

                                                                int _cId = GetLookupIDsManageTag(newValue, _cContext, _Web);

                                                                if (_cId != 0)
                                                                {
                                                                    FieldLookupValue flv = new FieldLookupValue();
                                                                    flv.LookupId = _cId;

                                                                    lookupFieldValCollection.SetValue(flv, i);
                                                                }
                                                            }

                                                            if (lookupFieldValCollection.Length >= 1)
                                                            {
                                                                if (lookupFieldValCollection[0] != null)
                                                                    oItem["Tag"] = lookupFieldValCollection;
                                                            }

                                                            oItem.Update();
                                                            _cContext.Load(oItem);
                                                            _cContext.ExecuteQuery();

                                                            excelWriterTagListCreation.WriteLine(_SiteURL + "," + _ListName + "," + _ItemID + "," + "Success");
                                                            excelWriterTagListCreation.Flush();

                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            excelWriterTagListCreation.WriteLine(_SiteURL + "," + _ListName + "," + _ItemID + "," + "Error: " + ex.Message.Replace(",", ""));
                                                            excelWriterTagListCreation.Flush();
                                                        }

                                                        try
                                                        {
                                                            oItem["Modified"] = Modified;
                                                            oItem["Editor"] = ModifiedBy;
                                                            oItem.Update();
                                                            _cContext.ExecuteQuery();
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                }
                                            }

                                            targetList.EnableVersioning = true;
                                            targetList.Update();
                                            _cContext.ExecuteQuery();
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
        }

        private void button61_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("SpaceID", typeof(string)), new DataColumn("ObjectId", typeof(string)), new DataColumn("ObjectType", typeof(string)), new DataColumn("ImportedURL", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            DataTable dtJiveObjects = new DataTable();
            dtJiveObjects.Columns.AddRange(new DataColumn[6] { new DataColumn("PlaceID", typeof(string)), new DataColumn("ID", typeof(string)), new DataColumn("Type", typeof(string)), new DataColumn("Tags", typeof(string)), new DataColumn("Modified", typeof(string)), new DataColumn("ModifiedBy", typeof(string)) });

            string csvData1 = System.IO.File.ReadAllText(textBox3.Text);

            foreach (string row in csvData1.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtJiveObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "CategoriesFixReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("ObjectID" + "," + "SPURL" + "," + "Tags" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;

            string[] SiteSplit = new string[] { "/Lists/" };
            string[] IDSplit = new string[] { "?ID=" };
            string[] DocumentSplit = new string[] { "/Documents/" };
            string[] TagsSplit = new string[] { "|" };
            string[] FileURLSplit = new string[] { "/1_Uploaded Files/" };
            string[] PageURLSplit = new string[] { "/Pages/" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                count++;
                try
                {
                    string TagsColl = string.Empty;

                    string _SPSpaceID = string.Empty;
                    string _JivePlaceID = string.Empty;

                    string _objectID = string.Empty;
                    string _itemID = string.Empty;
                    string _FilePath = string.Empty;
                    string _objectURL = string.Empty;
                    string _importedURL = string.Empty;
                    string _ListName = string.Empty;
                    string _objList = string.Empty;

                    string _Modified = string.Empty;
                    string _ModifiedBY = string.Empty;

                    _SPSpaceID = drImported["SpaceID"].ToString().Trim();

                    _objectID = drImported["ObjectId"].ToString().Trim();
                    _objectURL = drImported["ImportedURL"].ToString().Trim();
                    _importedURL = drImported["ImportedURL"].ToString().Trim();
                    _objList = drImported["ObjectType"].ToString().Trim();

                    #region OLD TYPE

                    //if (_importedURL.Contains("/1_Uploaded Files/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "1_Uploaded Files";
                    //}
                    //if (_importedURL.Contains("/2_Documents and Pages/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "2_Documents and Pages";
                    //}
                    //if (_importedURL.Contains("/Discussions/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Discussions";
                    //}
                    //if (_importedURL.Contains("/Events/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Events";
                    //}
                    //if (_importedURL.Contains("/Messages/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Messages";
                    //}
                    //if (_importedURL.Contains("/Posts/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Posts";
                    //}
                    //if (_importedURL.Contains("/Site Assets/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Site Assets";
                    //}
                    //if (_importedURL.Contains("/SiteHistory/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "SiteHistory";
                    //}
                    //if (_importedURL.Contains("/Tasks/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Tasks";
                    //} 

                    #endregion

                    this.Text = (count).ToString() + " of " + dtImportedObjects.Rows.Count.ToString() + " : " + _objectURL;

                    bool itemFound = false;

                    foreach (DataRow drJive in dtJiveObjects.Rows)
                    {
                        //_JivePlaceID = drImported["PlaceID"].ToString().Trim();
                        if ((drImported["ObjectId"].ToString().Trim() == drJive["ID"].ToString().Trim()) && (drImported["ObjectType"].ToString().Trim() == drJive["Type"].ToString().Trim()))
                        {
                            TagsColl = drJive["Tags"].ToString().Trim();
                            _Modified = drJive["Modified"].ToString().Trim();
                            _ModifiedBY = drJive["ModifiedBy"].ToString().Trim();
                            itemFound = true;
                            break;
                        }
                    }

                    if (itemFound && !string.IsNullOrEmpty(TagsColl))
                    {
                        #region Get Site URL

                        if (_importedURL.Contains("/Lists/"))
                        {
                            _importedURL = drImported["ImportedURL"].ToString().Split(SiteSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                        }
                        if (_importedURL.Contains("/1_Uploaded Files/"))
                        {
                            _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                        }
                        if (_importedURL.Contains("/Pages/"))
                        {
                            _importedURL = drImported["ImportedURL"].ToString().Split(PageURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                        }

                        #endregion

                        #region Get List and FilePath/ItemID

                        switch (_objList)
                        {
                            case "Document":
                                _ListName = "2_Documents and Pages";
                                _FilePath = drImported["ImportedURL"].ToString().Split(DocumentSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "File":
                                _ListName = "1_Uploaded Files";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Blog":
                                _ListName = "Posts";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Discussion":
                                _ListName = "Discussions";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Event":
                                _ListName = "Events";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Task":
                                _ListName = "Tasks";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Idea":
                                _ListName = "Ideas";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;
                        }

                        #endregion

                        #region Tags Re-Apply

                        AuthenticationManager authManager = new AuthenticationManager();
                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, textBox6.Text, textBox5.Text))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();

                            List _List = null;

                            try
                            {
                                _List = clientcontext.Web.Lists.GetByTitle(_ListName);
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                if (_List.Title == "2_Documents and Pages")
                                {
                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    _List.ForceCheckout = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    try
                                    {
                                        clientcontext.Load(_List.RootFolder);
                                        clientcontext.ExecuteQuery();

                                        Folder docFolder = null;

                                        try
                                        {
                                            docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                            clientcontext.Load(docFolder);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        { }

                                        if (docFolder != null)
                                        {
                                            ListItem _Item = docFolder.Files.GetByUrl(_FilePath).ListItemAllFields;
                                            clientcontext.Load(_Item);
                                            clientcontext.ExecuteQuery();

                                            #region Document Modified, ModifiedBy

                                            DateTime Modified = new DateTime();
                                            FieldUserValue ModifiedBy = null;

                                            try
                                            {
                                                if (!string.IsNullOrEmpty(_Modified))
                                                {
                                                    Modified = getdateformat(_Modified);
                                                }
                                                else
                                                {
                                                    Modified = Convert.ToDateTime(_Item["Modified"]);
                                                }

                                                if (!string.IsNullOrEmpty(_ModifiedBY))
                                                {
                                                    User ModifiedUser = default(User);
                                                    try
                                                    {
                                                        ModifiedUser = clientcontext.Web.EnsureUser(_ModifiedBY);
                                                        clientcontext.Load(ModifiedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        ModifiedUser = clientcontext.Web.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                                        clientcontext.Load(ModifiedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }

                                                    ModifiedBy = new FieldUserValue();
                                                    ModifiedBy.LookupId = ModifiedUser.Id;
                                                }
                                                else
                                                {
                                                    ModifiedBy = (FieldUserValue)_Item["Editor"];
                                                }
                                            }
                                            catch (Exception ex)
                                            {

                                            }

                                            #endregion

                                            //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                            //FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                            if (!string.IsNullOrEmpty(TagsColl))
                                            {
                                                string[] _categories = TagsColl.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                                                try
                                                {
                                                    FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                                                    for (int i = 0; i <= _categories.Length - 1; i++)
                                                    {
                                                        string newValue = _categories[i].ToString();

                                                        if (_categories[i].ToString().Contains("$"))
                                                        {
                                                            newValue = _categories[i].ToString().Replace("$", ",");
                                                        }

                                                        int _cId = GetLookupIDsManageTag(newValue, clientcontext, oWeb);

                                                        if (_cId != 0)
                                                        {
                                                            FieldLookupValue flv = new FieldLookupValue();
                                                            flv.LookupId = _cId;

                                                            lookupFieldValCollection.SetValue(flv, i);
                                                        }
                                                    }

                                                    if (lookupFieldValCollection.Length >= 1)
                                                    {
                                                        if (lookupFieldValCollection[0] != null)
                                                            _Item["Tag"] = lookupFieldValCollection;
                                                    }

                                                    _Item.Update();
                                                    clientcontext.Load(_Item);
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "TAgApplyFailure");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }

                                                try
                                                {
                                                    _Item["Modified"] = Modified;
                                                    _Item["Editor"] = ModifiedBy;
                                                    _Item.Update();
                                                    clientcontext.ExecuteQuery();

                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "Success");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                                catch (Exception ex)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ModifyFailure");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        //excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Failure due to : " + ex.Message);
                                        //excelWriterScoringMatrixNew.Flush();
                                    }

                                    _List.EnableVersioning = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    _List.ForceCheckout = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();
                                }
                                else
                                {
                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    try
                                    {
                                        bool tagsFileldExist = _List.FieldExistsByName("Tag");

                                        if (tagsFileldExist)
                                        {

                                            ListItem _Item = _List.GetItemById(_itemID);
                                            clientcontext.Load(_Item);
                                            clientcontext.ExecuteQuery();

                                            #region Document Modified, ModifiedBy

                                            DateTime Modified = new DateTime();
                                            FieldUserValue ModifiedBy = null;

                                            try
                                            {
                                                if (!string.IsNullOrEmpty(_Modified))
                                                {
                                                    Modified = getdateformat(_Modified);
                                                }
                                                else
                                                {
                                                    Modified = Convert.ToDateTime(_Item["Modified"]);
                                                }

                                                if (!string.IsNullOrEmpty(_ModifiedBY))
                                                {
                                                    User ModifiedUser = default(User);
                                                    try
                                                    {
                                                        ModifiedUser = clientcontext.Web.EnsureUser(_ModifiedBY);
                                                        clientcontext.Load(ModifiedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        ModifiedUser = clientcontext.Web.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                                        clientcontext.Load(ModifiedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }

                                                    ModifiedBy = new FieldUserValue();
                                                    ModifiedBy.LookupId = ModifiedUser.Id;
                                                }
                                                else
                                                {
                                                    ModifiedBy = (FieldUserValue)_Item["Editor"];
                                                }
                                            }
                                            catch (Exception ex)
                                            {

                                            }

                                            #endregion

                                            //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                            //FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                            if (!string.IsNullOrEmpty(TagsColl))
                                            {
                                                string[] _categories = TagsColl.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                                                try
                                                {
                                                    FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                                                    for (int i = 0; i <= _categories.Length - 1; i++)
                                                    {
                                                        string newValue = _categories[i].ToString();

                                                        if (_categories[i].ToString().Contains("$"))
                                                        {
                                                            newValue = _categories[i].ToString().Replace("$", ",");
                                                        }

                                                        int _cId = GetLookupIDsManageTag(newValue, clientcontext, oWeb);

                                                        if (_cId != 0)
                                                        {
                                                            FieldLookupValue flv = new FieldLookupValue();
                                                            flv.LookupId = _cId;

                                                            lookupFieldValCollection.SetValue(flv, i);
                                                        }
                                                    }

                                                    if (lookupFieldValCollection.Length >= 1)
                                                    {
                                                        if (lookupFieldValCollection[0] != null)
                                                            _Item["Tag"] = lookupFieldValCollection;
                                                    }

                                                    _Item.Update();
                                                    clientcontext.Load(_Item);
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "TAgApplyFailure");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }

                                                try
                                                {
                                                    _Item["Modified"] = Modified;
                                                    _Item["Editor"] = ModifiedBy;
                                                    _Item.Update();
                                                    clientcontext.ExecuteQuery();

                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "Success");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                                catch (Exception ex)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ModifyFailure");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception EX)
                                    {
                                        excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ItemIDFailure");
                                        excelWriterScoringMatrixNew.Flush();
                                    }

                                    _List.EnableVersioning = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();
                                }
                            }
                        }

                        #endregion
                    }
                    else
                    {
                        excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ItemIDNotFoundinJive");
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button62_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            List<string> ListNames = new List<string>();

            //ListNames.Add("Team Files");
            //ListNames.Add("Team Files");
            //ListNames.Add("Uploaded Files");
            //ListNames.Add("Documents and Pages");
            ListNames.Add("1_Uploaded Files");
            ListNames.Add("2_Documents and Pages");
            ListNames.Add("Discussions");
            ListNames.Add("Events");
            ListNames.Add("Messages");
            ListNames.Add("Posts");
            ListNames.Add("Site Assets");
            //ListNames.Add("SiteHistory");
            ListNames.Add("Tasks");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), textBox6.Text, textBox5.Text))
                    {
                        Web _Web = _cContext.Web;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        StreamWriter excelWriterTagsReport = null;
                        excelWriterTagsReport = System.IO.File.CreateText(textBox2.Text + "\\" + _Web.Id.ToString() + "_TagsApplyReport" + ".csv");
                        excelWriterTagsReport.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "Tags");
                        excelWriterTagsReport.Flush();

                        List _List = null;

                        List<string> _UniqueTags = new List<string>();
                        string _strUniqueTags = string.Empty;

                        foreach (string ls in ListNames)
                        {
                            try
                            {
                                _List = _cContext.Web.Lists.GetByTitle(ls);
                                _cContext.Load(_List);
                                _cContext.ExecuteQuery();

                                _List.EnableVersioning = false;
                                _List.Update();
                                _cContext.ExecuteQuery();

                                if (ls == "2_Documents and Pages")
                                {
                                    try
                                    {
                                        _List.ForceCheckout = false;
                                        _List.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                    }
                                }

                                bool tagsFileldExist = _List.FieldExistsByName("Tags");

                                bool NewtagFileldExist = _List.FieldExistsByName("Tag");

                                if (tagsFileldExist && NewtagFileldExist)
                                {
                                    CamlQuery camlQuery = new CamlQuery();
                                    camlQuery.ViewXml = "<View Scope='RecursiveAll'></View>";//<RowLimit>5000</RowLimit>

                                    ListItemCollection listItems = _List.GetItems(camlQuery);
                                    _cContext.Load(listItems);
                                    _cContext.ExecuteQuery();

                                    foreach (ListItem oItem in listItems)
                                    {
                                        try
                                        {
                                            string Tags = string.Empty;

                                            _cContext.Load(oItem);
                                            _cContext.ExecuteQuery();

                                            string itemID = oItem.Id.ToString();

                                            DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                                            FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                                            TaxonomyFieldValueCollection taxFieldValues = oItem["Tags"] as TaxonomyFieldValueCollection;
                                            FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[taxFieldValues.Count];
                                            int i = 0;

                                            foreach (TaxonomyFieldValue tv in taxFieldValues)
                                            {
                                                int _cId = GetLookupIDsManageTag(tv.Label.ToString(), _cContext, _Web);

                                                if (_cId != 0)
                                                {
                                                    Tags += tv.Label.ToString() + "|";

                                                    FieldLookupValue flv = new FieldLookupValue();
                                                    flv.LookupId = _cId;
                                                    lookupFieldValCollection.SetValue(flv, i);
                                                    i++;
                                                }
                                            }

                                            if (lookupFieldValCollection.Length >= 1)
                                            {
                                                if (lookupFieldValCollection[0] != null)
                                                    oItem["Tag"] = lookupFieldValCollection;

                                                oItem.Update();
                                                _cContext.Load(oItem);
                                                _cContext.ExecuteQuery();

                                                try
                                                {
                                                    oItem["Modified"] = Modified;
                                                    oItem["Editor"] = ModifiedBy;
                                                    oItem.Update();
                                                    _cContext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {

                                                }

                                                excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + ls + "," + itemID + "," + Tags);
                                                excelWriterTagsReport.Flush();
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            continue;
                                        }
                                    }
                                }

                                _List.EnableVersioning = true;
                                _List.Update();
                                _cContext.ExecuteQuery();

                                if (ls == "2_Documents and Pages")
                                {
                                    try
                                    {
                                        _List.ForceCheckout = true;
                                        _List.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + ls + "," + "ERROR" + "," + ex.Message.Replace(",", ""));
                                excelWriterTagsReport.Flush();
                                continue;
                            }
                        }

                        excelWriterTagsReport.Flush();
                        excelWriterTagsReport.Close();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            #endregion   

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button63_Click(object sender, EventArgs e)
        {

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[2] { new DataColumn("TermGUID", typeof(string)), new DataColumn("Term", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "TagRemoveByIDReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("TermGUID" + "," + "Term" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            try
            {
                AuthenticationManager authManager = new AuthenticationManager();
                using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant("https://rsharepoint.sharepoint.com/sites/rworldgroups", "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                {
                    // Get the TaxonomySession
                    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientcontext);

                    // Get the term store by name
                    TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");

                    // Get the term group by Name
                    TermGroup termGroup = termStore.Groups.GetByName("RicohTags");

                    // Get the term set by Name
                    TermSet termSet = termGroup.TermSets.GetByName("TagsTermSet");

                    // Get all the terms 
                    TermCollection termColl = termSet.Terms;
                    clientcontext.Load(termColl);
                    clientcontext.ExecuteQuery();

                    int j = 0;

                    foreach (DataRow drImported in dtImportedObjects.Rows)
                    {
                        j++;

                        string termGUID = drImported["TermGUID"].ToString().Trim();
                        string termName = drImported["Term"].ToString().Trim();

                        this.Text = j.ToString() + " : " + termGUID;

                        try
                        {
                            Guid termID = termGUID.Trim().ToGuid();
                            Term tm = termColl.GetById(termID);
                            clientcontext.Load(tm);
                            clientcontext.ExecuteQuery();

                            string originalTerm = tm.Name;

                            //if (tm.Name == termName)
                            {
                                tm.DeleteObject();
                                termStore.CommitAll();
                                clientcontext.ExecuteQuery();

                                excelWriterScoringMatrixNew.WriteLine(termGUID + "," + originalTerm + "," + "Success");
                                excelWriterScoringMatrixNew.Flush();
                            }
                            //else
                            //{
                            //    excelWriterScoringMatrixNew.WriteLine(termGUID + "," + termName + "," + "Original Term: "+ tm.Name);
                            //    excelWriterScoringMatrixNew.Flush();
                            //}
                        }
                        catch (Exception ex)
                        {
                            excelWriterScoringMatrixNew.WriteLine(termGUID + "," + termName + "," + "Failure : " + ex.Message.Replace(",", ""));
                            excelWriterScoringMatrixNew.Flush();
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Completed.";
            MessageBox.Show("Process completed Successfully.");
        }

        private void button64_Click(object sender, EventArgs e)
        {
            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("SpaceID", typeof(string)), new DataColumn("JiveURL", typeof(string)), new DataColumn("SPURL", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            DataTable dtJiveObjects = new DataTable();
            dtJiveObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("PlaceID", typeof(string)), new DataColumn("JiveURL", typeof(string)), new DataColumn("Tags", typeof(string)) });

            string csvData1 = System.IO.File.ReadAllText(textBox3.Text);

            foreach (string row in csvData1.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtJiveObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteHistoryTagsFixReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SpaceID " + "," + "JiveURL" + "," + "SPURL" + "," + "Tags");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                count++;
                try
                {
                    string TagsColl = string.Empty;

                    string _SpaceID = string.Empty;
                    string _JiveURL = string.Empty;
                    string _importedURL = string.Empty;

                    _SpaceID = drImported["SpaceID"].ToString().Trim();
                    _JiveURL = drImported["JiveURL"].ToString().Trim();
                    _importedURL = drImported["SPURL"].ToString().Trim();

                    this.Text = (count).ToString() + " of " + dtImportedObjects.Rows.Count.ToString() + " : " + _JiveURL;

                    bool itemFound = false;

                    foreach (DataRow drJive in dtJiveObjects.Rows)
                    {
                        if ((drImported["SpaceID"].ToString().Trim() == drJive["PlaceID"].ToString().Trim()) && (drImported["JiveURL"].ToString().Trim() == drJive["JiveURL"].ToString().Trim()))
                        {
                            TagsColl = drJive["Tags"].ToString().Trim();
                            itemFound = true;
                            break;
                        }
                    }

                    if (itemFound && !string.IsNullOrEmpty(TagsColl))
                    {
                        #region Tags Re-Apply

                        AuthenticationManager authManager = new AuthenticationManager();
                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();

                            List _List = null;

                            try
                            {
                                _List = clientcontext.Web.Lists.GetByTitle("SiteHistory");
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                //_List.EnableVersioning = false;
                                //_List.Update();
                                //clientcontext.ExecuteQuery();

                                try
                                {
                                    bool tagsFileldExist = _List.FieldExistsByName("Tag");

                                    if (tagsFileldExist)
                                    {
                                        CamlQuery camlQuery = new CamlQuery();
                                        camlQuery.ViewXml = "<View><RowLimit>1</RowLimit></View>";

                                        ListItemCollection listItems = _List.GetItems(camlQuery);
                                        clientcontext.Load(listItems);
                                        clientcontext.ExecuteQuery();

                                        ListItem _Item = listItems[0];
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();

                                        //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                        //FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                        if (!string.IsNullOrEmpty(TagsColl))
                                        {
                                            string[] _categories = TagsColl.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                                            try
                                            {
                                                FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                                                for (int i = 0; i <= _categories.Length - 1; i++)
                                                {
                                                    string newValue = _categories[i].ToString();

                                                    if (_categories[i].ToString().Contains("$"))
                                                    {
                                                        newValue = _categories[i].ToString().Replace("$", ",");
                                                    }

                                                    int _cId = GetLookupIDsManageTag(newValue, clientcontext, oWeb);

                                                    if (_cId != 0)
                                                    {
                                                        FieldLookupValue flv = new FieldLookupValue();
                                                        flv.LookupId = _cId;

                                                        lookupFieldValCollection.SetValue(flv, i);
                                                    }
                                                }

                                                if (lookupFieldValCollection.Length >= 1)
                                                {
                                                    if (lookupFieldValCollection[0] != null)
                                                        _Item["Tag"] = lookupFieldValCollection;
                                                }

                                                _Item.SystemUpdate();
                                                clientcontext.Load(_Item);
                                                clientcontext.ExecuteQuery();

                                                excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + TagsColl);
                                                excelWriterScoringMatrixNew.Flush();
                                            }
                                            catch (Exception ex)
                                            {
                                                excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + TagsColl);
                                                excelWriterScoringMatrixNew.Flush();
                                            }

                                            #region System Update Substitute

                                            //try
                                            //{
                                            //    _Item["Modified"] = Modified;
                                            //    _Item["Editor"] = ModifiedBy;
                                            //    _Item.Update();
                                            //    clientcontext.ExecuteQuery();

                                            //    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "Success");
                                            //    excelWriterScoringMatrixNew.Flush();
                                            //}
                                            //catch (Exception ex)
                                            //{
                                            //    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ModifyFailure");
                                            //    excelWriterScoringMatrixNew.Flush();
                                            //} 

                                            #endregion
                                        }
                                    }
                                }
                                catch (Exception EX)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + "ItemIDFailure");
                                    excelWriterScoringMatrixNew.Flush();
                                }

                                //_List.EnableVersioning = true;
                                //_List.Update();
                                //clientcontext.ExecuteQuery();
                            }
                        }

                        #endregion

                    }
                    else
                    {
                        excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + "ItemIDNotFoundinJive");
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button65_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig            

            StreamWriter excelWriterTagListCreation = null;
            excelWriterTagListCreation = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteError_TagsColumnDeletion" + ".csv");
            excelWriterTagListCreation.WriteLine("SiteURL" + "," + "Status" + "," + "Details");
            excelWriterTagListCreation.Flush();

            StreamWriter excelWriterTagColumnCreation = null;
            excelWriterTagColumnCreation = System.IO.File.CreateText(textBox2.Text + "\\" + "ListError_TagsColumnDeletion" + ".csv");
            excelWriterTagColumnCreation.WriteLine("SiteURL" + "," + "ListName" + "," + "Details");
            excelWriterTagColumnCreation.Flush();

            List<string> ListNames = new List<string>();

            //ListNames.Add("Uploaded Files");
            //ListNames.Add("Documents and Pages");
            //ListNames.Add("1_Uploaded Files");
            //ListNames.Add("2_Documents and Pages");
            ListNames.Add("Discussions");
            ListNames.Add("Events");
            ListNames.Add("Messages");
            ListNames.Add("Posts");
            ListNames.Add("SiteHistory");
            ListNames.Add("Tasks");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration3@rsharepoint.onmicrosoft.com", "Goj72326"))
                    {
                        Web _Web = _cContext.Web;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        foreach (string ls in ListNames)
                        {
                            try
                            {
                                List _List = null;
                                _List = _cContext.Web.Lists.GetByTitle(ls);
                                _cContext.Load(_List);
                                _cContext.ExecuteQuery();

                                bool tagsFileldExist = _List.FieldExistsByName("Tags");

                                if (tagsFileldExist)
                                {
                                    try
                                    {
                                        FieldCollection fieldColl = _List.Fields;
                                        _cContext.Load(fieldColl);
                                        _cContext.ExecuteQuery();

                                        Field tagsField = fieldColl.GetByInternalNameOrTitle("Tags");
                                        _cContext.Load(tagsField);
                                        _cContext.ExecuteQuery();

                                        tagsField.DeleteObject();
                                        _cContext.ExecuteQuery();

                                    }
                                    catch (Exception ex)
                                    {
                                        excelWriterTagColumnCreation.WriteLine(_cContext.Web.Url + "," + ls + "," + "Error : " + ex.Message.Replace(",", ""));
                                        excelWriterTagColumnCreation.Flush();

                                        continue;
                                    }
                                }

                                #region Views

                                if (ls == "1_Uploaded Files")
                                {
                                    try
                                    {
                                        ViewCollection ViewColl = _List.Views;
                                        _cContext.Load(ViewColl);
                                        _cContext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        _cContext.Load(v);
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.RemoveAll();
                                        v.Update();
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.Add("DocIcon");
                                        v.ViewFields.Add("Title");
                                        v.ViewFields.Add("LinkFilename");
                                        v.ViewFields.Add("Created");
                                        v.ViewFields.Add("Created By");
                                        v.ViewFields.Add("Modified");
                                        v.ViewFields.Add("Modified By");
                                        v.ViewFields.Add("Tag");
                                        v.ViewFields.Add("Categorization");
                                        v.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }

                                if (ls == "2_Documents and Pages")
                                {
                                    try
                                    {
                                        ViewCollection ViewColl = _List.Views;
                                        _cContext.Load(ViewColl);
                                        _cContext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        _cContext.Load(v);
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.RemoveAll();
                                        v.Update();
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.Add("DocIcon");
                                        v.ViewFields.Add("Title");
                                        v.ViewFields.Add("LinkFilename");
                                        v.ViewFields.Add("Created");
                                        v.ViewFields.Add("Created By");
                                        v.ViewFields.Add("Modified");
                                        v.ViewFields.Add("Modified By");
                                        v.ViewFields.Add("Tag");
                                        v.ViewFields.Add("Categorization");
                                        v.ViewFields.Add("CheckoutUser");
                                        v.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    { }
                                }

                                if (ls == "Posts")
                                {
                                    try
                                    {
                                        ViewCollection ViewColl = _List.Views;
                                        _cContext.Load(ViewColl);
                                        _cContext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        _cContext.Load(v);
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.RemoveAll();
                                        v.Update();
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.Add("LinkTitle");
                                        v.ViewFields.Add("Created");
                                        v.ViewFields.Add("Published");
                                        v.ViewFields.Add("Category");
                                        v.ViewFields.Add("NumComments");
                                        v.ViewFields.Add("Edit");
                                        v.ViewFields.Add("Categorization");
                                        v.ViewFields.Add("LikesCount");
                                        v.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    { }
                                }

                                #endregion

                                #region Site Assets View

                                if (ls == "Site Assets")
                                {
                                    try
                                    {
                                        ViewCollection ViewColl = _List.Views;
                                        _cContext.Load(ViewColl);
                                        _cContext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        _cContext.Load(v);
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.Add("DocIcon");
                                        v.ViewFields.Add("Title");
                                        v.ViewFields.Add("LinkFilename");
                                        v.ViewFields.Add("Created");
                                        v.ViewFields.Add("Created By");
                                        v.ViewFields.Add("Modified");
                                        v.ViewFields.Add("Modified By");
                                        v.ViewFields.Add("CheckoutUser");
                                        v.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    { }
                                }

                                #endregion

                            }
                            catch (Exception ex)
                            {
                                //excelWriterTagColumnCreation.WriteLine(_cContext.Web.Url + "," + ls + "," + "ListNotExist");
                                //excelWriterTagColumnCreation.Flush();

                                continue;
                            }
                        }

                        ListCollection oListColl = _cContext.Web.Lists;
                        _cContext.Load(oListColl);
                        _cContext.ExecuteQuery();

                        foreach (List _List in oListColl)
                        {
                            _cContext.Load(_List);
                            _cContext.ExecuteQuery();

                            _cContext.Load(_List.RootFolder);
                            _cContext.ExecuteQuery();

                            string listPath = _List.RootFolder.ServerRelativeUrl;

                            if (listPath.ToLower().EndsWith("/pages") || listPath.ToLower().EndsWith("/1_uploaded files"))
                            {
                                #region Views

                                bool tagsFileldExist = _List.FieldExistsByName("Tags");

                                if (tagsFileldExist)
                                {
                                    try
                                    {
                                        FieldCollection fieldColl = _List.Fields;
                                        _cContext.Load(fieldColl);
                                        _cContext.ExecuteQuery();

                                        Field tagsField = fieldColl.GetByInternalNameOrTitle("Tags");
                                        _cContext.Load(tagsField);
                                        _cContext.ExecuteQuery();

                                        tagsField.DeleteObject();
                                        _cContext.ExecuteQuery();

                                    }
                                    catch (Exception ex)
                                    {
                                        excelWriterTagColumnCreation.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + "Error : " + ex.Message.Replace(",", ""));
                                        excelWriterTagColumnCreation.Flush();
                                        continue;
                                    }
                                }

                                if (listPath.ToLower().EndsWith("/1_uploaded files"))
                                {
                                    try
                                    {
                                        ViewCollection ViewColl = _List.Views;
                                        _cContext.Load(ViewColl);
                                        _cContext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        _cContext.Load(v);
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.RemoveAll();
                                        v.Update();
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.Add("DocIcon");
                                        v.ViewFields.Add("Title");
                                        v.ViewFields.Add("LinkFilename");
                                        v.ViewFields.Add("Created");
                                        v.ViewFields.Add("Created By");
                                        v.ViewFields.Add("Modified");
                                        v.ViewFields.Add("Modified By");
                                        v.ViewFields.Add("Tag");
                                        v.ViewFields.Add("Categorization");
                                        v.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }

                                if (listPath.ToLower().EndsWith("/pages"))
                                {
                                    try
                                    {
                                        ViewCollection ViewColl = _List.Views;
                                        _cContext.Load(ViewColl);
                                        _cContext.ExecuteQuery();

                                        Microsoft.SharePoint.Client.View v = ViewColl[0];
                                        _cContext.Load(v);
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.RemoveAll();
                                        v.Update();
                                        _cContext.ExecuteQuery();

                                        v.ViewFields.Add("DocIcon");
                                        v.ViewFields.Add("Title");
                                        v.ViewFields.Add("LinkFilename");
                                        v.ViewFields.Add("Created");
                                        v.ViewFields.Add("Created By");
                                        v.ViewFields.Add("Modified");
                                        v.ViewFields.Add("Modified By");
                                        v.ViewFields.Add("Tag");
                                        v.ViewFields.Add("Categorization");
                                        v.ViewFields.Add("CheckoutUser");
                                        v.Update();
                                        _cContext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    { }
                                }

                                #endregion
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterTagListCreation.WriteLine(lstSiteColl[j].ToString() + "," + "SiteNotAccessed" + "," + ex.Message.Replace(",", ""));
                    excelWriterTagListCreation.Flush();

                    continue;
                }
            }

            #endregion                       

            excelWriterTagColumnCreation.Flush();
            excelWriterTagColumnCreation.Close();

            excelWriterTagListCreation.Flush();
            excelWriterTagListCreation.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button66_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteHistoryItemCTRemove" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "TagColumnDeleted" + "," + "ItemCTRemoved" + "," + "TagColumnAdded");
            excelWriterScoringMatrixNew.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " : " + lstSiteColl[j].ToString();

                string TagColumnDeleted = "No";
                string ItemCTRemoved = "No";
                string TagColumnAdded = "No";

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var context = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        context.Load(context.Web);
                        context.ExecuteQuery();

                        ListCollection listColl = context.Web.Lists;
                        context.Load(listColl);
                        context.ExecuteQuery();

                        List list = null;

                        try
                        {
                            list = listColl.GetByTitle("SiteHistory");
                            //list = listColl.GetByTitle("1_Uploaded Files");
                            context.Load(list);
                            context.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {

                        }

                        if (list != null)
                        {
                            #region Delete "Tag" Column

                            bool tagFileldExist = list.FieldExistsByName("Tag");

                            if (tagFileldExist)
                            {
                                try
                                {
                                    FieldCollection fieldColl = list.Fields;
                                    context.Load(fieldColl);
                                    context.ExecuteQuery();

                                    Field tagsField = fieldColl.GetByInternalNameOrTitle("Tag");
                                    context.Load(tagsField);
                                    context.ExecuteQuery();

                                    tagsField.DeleteObject();
                                    context.ExecuteQuery();

                                    TagColumnDeleted = "Yes";
                                }
                                catch (Exception ex)
                                {
                                }
                            }

                            #endregion

                            #region Item CT Delete

                            try
                            {
                                ContentTypeCollection contentTypeColl = list.ContentTypes;
                                context.Load(contentTypeColl);
                                context.ExecuteQuery();

                                ContentType defaultcontentType = null;
                                ContentType ricohcontentType = null;

                                foreach (ContentType eachcontenttype in contentTypeColl)
                                {
                                    context.Load(eachcontenttype);
                                    context.ExecuteQuery();

                                    if (eachcontenttype.Name == "Group_ContentType" || eachcontenttype.Name == "Space_ContentType" || eachcontenttype.Name == "Project_ContentType")
                                    //if (eachcontenttype.Name == "RicohContentType")
                                    {
                                        ricohcontentType = eachcontenttype;
                                        context.Load(ricohcontentType);
                                        context.ExecuteQuery();

                                        ricohcontentType.ReadOnly = false;
                                        ricohcontentType.Update(false);
                                        context.Load(ricohcontentType);
                                        context.ExecuteQuery();
                                    }

                                    if (eachcontenttype.Name == "Item")
                                    //if (eachcontenttype.Name == "Document")
                                    {
                                        defaultcontentType = eachcontenttype;
                                        context.Load(defaultcontentType);
                                        context.ExecuteQuery();

                                        defaultcontentType.DeleteObject();
                                        context.ExecuteQuery();

                                        ItemCTRemoved = "Yes";
                                    }
                                }
                            }
                            catch (Exception ex)
                            { }

                            #endregion

                            #region Add "Tag" Column

                            bool tagFileldExist1 = list.FieldExistsByName("Tag");

                            if (!tagFileldExist1)
                            {
                                try
                                {
                                    List olist = context.Web.Lists.GetByTitle("Manage Tag");
                                    context.Load(olist);
                                    context.ExecuteQuery();
                                    string schemaLookupField = "<Field Type='LookupMulti' Name='Tag' StaticName='Tag' DisplayName='Tag' List = '" + olist.Id + "' ShowField = 'Title' Mult = 'TRUE'/>";
                                    Field lookupField = list.Fields.AddFieldAsXml(schemaLookupField, true, AddFieldOptions.AddFieldInternalNameHint);
                                    list.Update();
                                    context.ExecuteQuery();

                                    TagColumnAdded = "Yes";
                                }
                                catch (Exception ex)
                                {
                                }
                            }

                            #endregion
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }

                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + TagColumnDeleted + "," + ItemCTRemoved + "," + TagColumnAdded);
                excelWriterScoringMatrixNew.Flush();
            }

            #endregion

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button67_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("SpaceID", typeof(string)), new DataColumn("JiveURL", typeof(string)), new DataColumn("SPURL", typeof(string)), new DataColumn("Tags", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            //DataTable dtJiveObjects = new DataTable();
            //dtJiveObjects.Columns.AddRange(new DataColumn[3] { new DataColumn("PlaceID", typeof(string)), new DataColumn("JiveURL", typeof(string)), new DataColumn("Tags", typeof(string)) });

            //string csvData1 = System.IO.File.ReadAllText(textBox3.Text);

            //foreach (string row in csvData1.Split('\n'))
            //{
            //    if (!string.IsNullOrEmpty(row))
            //    {
            //        dtJiveObjects.Rows.Add();
            //        int i = 0;

            //        foreach (string cell in row.Split(','))
            //        {
            //            dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
            //            i++;
            //        }
            //    }
            //}

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "SiteHistoryTagsFixReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SpaceID " + "," + "JiveURL" + "," + "SPURL" + "," + "Tags");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                count++;
                try
                {
                    string _SpaceID = string.Empty;
                    string _JiveURL = string.Empty;
                    string _importedURL = string.Empty;
                    string TagsColl = string.Empty;

                    _SpaceID = drImported["SpaceID"].ToString().Trim();
                    _JiveURL = drImported["JiveURL"].ToString().Trim();
                    _importedURL = drImported["SPURL"].ToString().Trim();
                    TagsColl = drImported["Tags"].ToString().Trim();

                    this.Text = (count).ToString() + " of " + dtImportedObjects.Rows.Count.ToString() + " : " + _JiveURL;

                    bool itemFound = true;

                    //foreach (DataRow drJive in dtJiveObjects.Rows)
                    //{
                    //    if ((drImported["SpaceID"].ToString().Trim() == drJive["PlaceID"].ToString().Trim()) && (drImported["JiveURL"].ToString().Trim() == drJive["JiveURL"].ToString().Trim()))
                    //    {
                    //        TagsColl = drJive["Tags"].ToString().Trim();
                    //        itemFound = true;
                    //        break;
                    //    }
                    //}

                    if (itemFound && !string.IsNullOrEmpty(TagsColl))
                    {
                        #region Tags Re-Apply

                        AuthenticationManager authManager = new AuthenticationManager();
                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();

                            List _List = null;

                            try
                            {
                                _List = clientcontext.Web.Lists.GetByTitle("SiteHistory");
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                //_List.EnableVersioning = false;
                                //_List.Update();
                                //clientcontext.ExecuteQuery();

                                try
                                {
                                    bool tagsFileldExist = _List.FieldExistsByName("Tag");

                                    if (tagsFileldExist)
                                    {
                                        CamlQuery camlQuery = new CamlQuery();
                                        camlQuery.ViewXml = "<View><RowLimit>1</RowLimit></View>";

                                        ListItemCollection listItems = _List.GetItems(camlQuery);
                                        clientcontext.Load(listItems);
                                        clientcontext.ExecuteQuery();

                                        ListItem _Item = listItems[0];
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();

                                        //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                        //FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                        if (!string.IsNullOrEmpty(TagsColl))
                                        {
                                            string[] _categories = TagsColl.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                                            try
                                            {
                                                FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                                                for (int i = 0; i <= _categories.Length - 1; i++)
                                                {
                                                    string newValue = _categories[i].ToString();

                                                    if (_categories[i].ToString().Contains("$"))
                                                    {
                                                        newValue = _categories[i].ToString().Replace("$", ",");
                                                    }

                                                    int _cId = GetLookupIDsManageTag(newValue, clientcontext, oWeb);

                                                    if (_cId != 0)
                                                    {
                                                        FieldLookupValue flv = new FieldLookupValue();
                                                        flv.LookupId = _cId;

                                                        lookupFieldValCollection.SetValue(flv, i);
                                                    }
                                                }

                                                if (lookupFieldValCollection.Length >= 1)
                                                {
                                                    if (lookupFieldValCollection[0] != null)
                                                        _Item["Tag"] = lookupFieldValCollection;
                                                }

                                                _Item.SystemUpdate();
                                                clientcontext.Load(_Item);
                                                clientcontext.ExecuteQuery();

                                                excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + TagsColl);
                                                excelWriterScoringMatrixNew.Flush();
                                            }
                                            catch (Exception ex)
                                            {
                                                excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + TagsColl);
                                                excelWriterScoringMatrixNew.Flush();
                                            }
                                        }
                                    }
                                }
                                catch (Exception EX)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + "NoItems");
                                    excelWriterScoringMatrixNew.Flush();
                                }

                                //_List.EnableVersioning = true;
                                //_List.Update();
                                //clientcontext.ExecuteQuery();
                            }
                        }

                        #endregion
                    }
                    else
                    {
                        excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + "NoTagsFound");
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button68_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("SpaceID", typeof(string)), new DataColumn("JiveURL", typeof(string)), new DataColumn("SPURL", typeof(string)), new DataColumn("PermissionType", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "AllRegisterdUsersPermissionApplyReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SpaceID " + "," + "JiveURL" + "," + "SPURL" + "," + "JivePermissionType" + "," + "o365PermissionType");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                count++;
                try
                {
                    string _SpaceID = string.Empty;
                    string _JiveURL = string.Empty;
                    string _importedURL = string.Empty;
                    string Permission = string.Empty;

                    _SpaceID = drImported["SpaceID"].ToString().Trim();
                    _JiveURL = drImported["JiveURL"].ToString().Trim();
                    _importedURL = drImported["SPURL"].ToString().Trim();
                    Permission = drImported["PermissionType"].ToString().Trim();

                    this.Text = (count).ToString() + " of " + dtImportedObjects.Rows.Count.ToString() + " : " + _importedURL;

                    if (!string.IsNullOrEmpty(Permission))
                    {
                        #region Tags Re-Apply

                        AuthenticationManager authManager = new AuthenticationManager();
                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                        {
                            try
                            {
                                Web oWeb = clientcontext.Web;
                                clientcontext.Load(oWeb);
                                clientcontext.ExecuteQuery();

                                RoleDefinition _cRoleDef = null;

                                RoleDefinitionCollection _newRoleDefs = clientcontext.Web.RoleDefinitions;
                                clientcontext.Load(_newRoleDefs);
                                clientcontext.ExecuteQuery();

                                try
                                {
                                    switch (Permission)
                                    {
                                        case "IT-Level":
                                        case "Channel Reporting":
                                        case "Create Video/Dis/Photo":
                                        case "Create Projects":
                                        case "Create":
                                            {
                                                _cRoleDef = _newRoleDefs.GetByName("Contribute");
                                            }
                                            break;

                                        case "Read":
                                        case "ReadOnly":
                                        case "RICOH Recognize":
                                        case "Contribute + Discuss":
                                        case "Contribute + Discuss + Projects":
                                        case "Discuss":
                                        case "iDiscuss":
                                            {
                                                _cRoleDef = _newRoleDefs.GetByName("Read");
                                            }
                                            break;

                                        case "Administer":
                                        case "Admin":
                                        case "Full Control":
                                        case "Site Admin":
                                        case "Administer + Moderate":
                                        case "Admin + Moderate":
                                        case "Administrator + Moderate":
                                        case "Administrator":
                                            {
                                                _cRoleDef = _newRoleDefs.GetByName("Site Admin");
                                            }
                                            break;

                                        default:
                                            {
                                                _cRoleDef = _newRoleDefs.GetByName("View Only");
                                            }
                                            break;
                                    }

                                    clientcontext.Load(_cRoleDef);
                                    clientcontext.ExecuteQuery();

                                    string SPpermission = _cRoleDef.Name;

                                    User CreatedUser = default(User);
                                    try
                                    {
                                        CreatedUser = clientcontext.Web.EnsureUser("Everyone except external users");
                                        clientcontext.Load(CreatedUser);
                                        clientcontext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        CreatedUser = clientcontext.Web.EnsureUser("Rworldadmin@rsharepoint.onmicrosoft.com");
                                        clientcontext.Load(CreatedUser);
                                        clientcontext.ExecuteQuery();
                                    }

                                    Principal _User = clientcontext.CastTo<Principal>(CreatedUser);
                                    RoleDefinitionBindingCollection _rdbColl = new RoleDefinitionBindingCollection(clientcontext);
                                    _rdbColl.Add(_cRoleDef);
                                    clientcontext.Web.RoleAssignments.Add(_User, _rdbColl);
                                    clientcontext.ExecuteQuery();

                                    excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + Permission + "," + SPpermission);
                                    excelWriterScoringMatrixNew.Flush();
                                }
                                catch (Exception ex)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + Permission + "," + "PermissionApplyError: " + ex.Message.Replace(",", ""));
                                    excelWriterScoringMatrixNew.Flush();
                                }
                            }
                            catch (Exception ex)
                            {
                                excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + Permission + "," + "SiteLoadingError: " + ex.Message.Replace(",", ""));
                                excelWriterScoringMatrixNew.Flush();
                            }
                        }
                    }

                    #endregion

                    else
                    {
                        excelWriterScoringMatrixNew.WriteLine(_SpaceID + "," + _JiveURL + "," + _importedURL + "," + "NoPermissionFoundinJive" + "," + "NA");
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button69_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            StreamWriter excelWriterSite_GUID_Report = null;
            excelWriterSite_GUID_Report = System.IO.File.CreateText(textBox2.Text + "\\" + "Site_GUID_Report" + ".csv");
            excelWriterSite_GUID_Report.WriteLine("SiteURL" + "," + "GUID" + "," + "Generated" + "," + "RowsCount" + "," + "Error?samesame");
            excelWriterSite_GUID_Report.Flush();

            StreamWriter excelWriterErrorLogReport = null;
            excelWriterErrorLogReport = System.IO.File.CreateText(textBox2.Text + "\\" + "ErrorLog" + ".csv");
            excelWriterErrorLogReport.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "Tags");
            excelWriterErrorLogReport.Flush();

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), textBox6.Text, textBox5.Text))
                    {
                        Web _Web = _cContext.Web;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        FileInfo _CsvFilepath = null;

                        //excelWriterSite_GUID_Report

                        try
                        {
                            _CsvFilepath = new DirectoryInfo(textBox2.Text).GetFiles(_Web.Id.ToString() + "_TagsApplyReport.csv", SearchOption.AllDirectories)[0];
                        }
                        catch
                        {
                        }

                        if (_CsvFilepath != null)
                        {
                            DataTable dtImportedObjects = new DataTable();
                            dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("SiteURL", typeof(string)), new DataColumn("ListName", typeof(string)), new DataColumn("ItemID", typeof(string)), new DataColumn("Tags", typeof(string)) });

                            string csvData1 = System.IO.File.ReadAllText(_CsvFilepath.FullName);

                            //Execute a loop over the rows.  
                            foreach (string row in csvData1.Split('\n'))
                            {
                                if (!string.IsNullOrEmpty(row))
                                {
                                    dtImportedObjects.Rows.Add();
                                    int i = 0;
                                    //Execute a loop over the columns.  
                                    foreach (string cell in row.Split(','))
                                    {
                                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                                        i++;
                                    }
                                }
                            }

                            string errors = "No";

                            foreach (DataRow drImported in dtImportedObjects.Rows)
                            {
                                try
                                {
                                    string samesamesolutions = drImported["Tags"].ToString().Trim();
                                    string erroeMessage = drImported["ItemID"].ToString().Trim();

                                    if (erroeMessage.ToLower() == "error" || samesamesolutions.ToLower().Contains("solutions") || samesamesolutions.ToLower().Contains("samesame"))
                                    {
                                        excelWriterErrorLogReport.WriteLine(drImported["SiteURL"].ToString().Trim() + "," + drImported["ListName"].ToString().Trim() + "," + drImported["ItemID"].ToString().Trim() + "," + drImported["Tags"].ToString().Trim());
                                        excelWriterErrorLogReport.Flush();
                                        errors = "Yes";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }
                            }

                            if (dtImportedObjects.Rows.Count > 1)
                            {
                                excelWriterSite_GUID_Report.WriteLine(_cContext.Web.Url + "," + _Web.Id.ToString() + "," + "Yes" + "," + dtImportedObjects.Rows.Count.ToString() + "," + errors);
                                excelWriterSite_GUID_Report.Flush();
                            }
                            else
                            {
                                excelWriterSite_GUID_Report.WriteLine(_cContext.Web.Url + "," + _Web.Id.ToString() + "," + "Yes" + "," + "0" + "," + errors);
                                excelWriterSite_GUID_Report.Flush();
                            }
                        }
                        else
                        {
                            excelWriterSite_GUID_Report.WriteLine(_cContext.Web.Url + "," + _Web.Id.ToString() + "," + "NoReport" + "," + "--" + "," + "--");
                            excelWriterSite_GUID_Report.Flush();
                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterSite_GUID_Report.WriteLine(lstSiteColl[j].ToString() + "," + "Error in Site Loading" + "," + "SiteNotAccessed" + "," + "NA" + "," + "NA");
                    excelWriterSite_GUID_Report.Flush();
                    continue;
                }
            }

            #endregion   

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button70_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[2] { new DataColumn("SiteURL", typeof(string)), new DataColumn("ItemID", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion           

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocURLReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ItemID" + "," + "DocURL");
            excelWriterScoringMatrixNew.Flush();

            int count = 1;

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                try
                {
                    string docID = drImported["ItemID"].ToString().Trim();
                    string _importedURL = drImported["SiteURL"].ToString().Trim();

                    this.Text = (count).ToString() + " : " + _importedURL;
                    count++;

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        List _List = null;
                        string listName = string.Empty;
                        string _FilePath = string.Empty;

                        listName = "2_Documents and Pages";

                        try
                        {
                            _List = clientcontext.Web.Lists.GetByTitle(listName);
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        { }

                        if (_List != null)
                        {
                            if (_List.Title == "2_Documents and Pages")
                            {

                                try
                                {
                                    clientcontext.Load(_List.RootFolder);
                                    clientcontext.ExecuteQuery();

                                    Folder docFolder = null;

                                    try
                                    {
                                        docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                        clientcontext.Load(docFolder);
                                        clientcontext.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    { }

                                    if (docFolder != null)
                                    {
                                        ListItem _Item = _List.GetItemById(docID);
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();

                                        string NameofFile = _Item["FileLeafRef"].ToString();

                                        excelWriterScoringMatrixNew.WriteLine(_importedURL + "," + docID + "," + _importedURL + "/Pages/Documents/" + NameofFile);
                                        excelWriterScoringMatrixNew.Flush();
                                    }
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button71_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }
            #endregion           

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "DocURLReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("SiteURL" + "," + "ItemID" + "," + "DocName" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 1;

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    string _importedURL = lstSiteColl[j].ToString().Trim();

                    this.Text = (count).ToString() + " : " + _importedURL;
                    count++;

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, "svc-jivemigration1@rsharepoint.onmicrosoft.com", "Vak52950"))
                    {
                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        List _List = null;
                        string listName = string.Empty;
                        string _FilePath = string.Empty;

                        listName = "2_Documents and Pages";

                        try
                        {
                            _List = clientcontext.Web.Lists.GetByTitle(listName);
                            clientcontext.Load(_List);
                            clientcontext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        { }

                        if (_List != null)
                        {
                            if (_List.Title == "2_Documents and Pages")
                            {
                                try
                                {
                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    //_List.ForceCheckout = false;
                                    //_List.Update();
                                    //clientcontext.ExecuteQuery();

                                    CamlQuery camlQuery = new CamlQuery();
                                    camlQuery.ViewXml = "<View Scope='RecursiveAll'></View>";//<RowLimit>5000</RowLimit>

                                    ListItemCollection listItems = _List.GetItems(camlQuery);
                                    clientcontext.Load(listItems);
                                    clientcontext.ExecuteQuery();

                                    foreach (ListItem _Item in listItems)
                                    {
                                        clientcontext.Load(_Item);
                                        clientcontext.ExecuteQuery();
                                        string _ItemName = string.Empty;
                                        string _ItemID = string.Empty;

                                        //bool checkedout = false;

                                        //try
                                        //{
                                        //    if (_Item.File.CheckOutType == CheckOutType.None)
                                        //    {
                                        //        checkedout = true;
                                        //    }
                                        //}
                                        //catch (Exception er)
                                        //{

                                        //}

                                        //if (checkedout)
                                        {
                                            try
                                            {
                                                DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                                FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];
                                                _ItemName = _Item["FileLeafRef"].ToString();
                                                _ItemID = _Item.Id.ToString();

                                                _Item["Title"] = _ItemName;
                                                _Item["Modified"] = Modified;
                                                _Item["Editor"] = ModifiedBy;
                                                _Item.Update();
                                                clientcontext.ExecuteQuery();

                                                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + _ItemID + "," + _ItemName + "," + "Success");
                                                excelWriterScoringMatrixNew.Flush();
                                            }
                                            catch (Exception ex)
                                            {
                                                string mess = ex.Message.Replace("\n\n", "");
                                                excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + _ItemID + "," + _ItemName + "," + "ItemError: " + mess.Replace(",", ""));
                                                excelWriterScoringMatrixNew.Flush();

                                                continue;
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "NA" + "," + "NA" + "," + "ListError: " + ex.Message.Replace(",", ""));
                                    excelWriterScoringMatrixNew.Flush();
                                }

                                _List.EnableVersioning = true;
                                _List.Update();
                                clientcontext.ExecuteQuery();

                                _List.ForceCheckout = true;
                                _List.Update();
                                clientcontext.ExecuteQuery();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    excelWriterScoringMatrixNew.WriteLine(lstSiteColl[j].ToString() + "," + "NA" + "," + "NA" + "," + "SiteLoadingError: " + ex.Message.Replace(",", ""));
                    excelWriterScoringMatrixNew.Flush();
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");


        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button72_Click(object sender, EventArgs e)
        {

            #region ImportedObjects CSV Reading

            DataTable dtImportedObjects = new DataTable();
            dtImportedObjects.Columns.AddRange(new DataColumn[4] { new DataColumn("SpaceID", typeof(string)), new DataColumn("ObjectId", typeof(string)), new DataColumn("ObjectType", typeof(string)), new DataColumn("ImportedURL", typeof(string)) });

            string csvData = System.IO.File.ReadAllText(textBox1.Text);

            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtImportedObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtImportedObjects.Rows[dtImportedObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            #region JiveObjects CSV Reading

            DataTable dtJiveObjects = new DataTable();
            dtJiveObjects.Columns.AddRange(new DataColumn[6] { new DataColumn("PlaceID", typeof(string)), new DataColumn("ID", typeof(string)), new DataColumn("Type", typeof(string)), new DataColumn("Tags", typeof(string)), new DataColumn("Modified", typeof(string)), new DataColumn("ModifiedBy", typeof(string)) });

            string csvData1 = System.IO.File.ReadAllText(textBox3.Text);

            foreach (string row in csvData1.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    dtJiveObjects.Rows.Add();
                    int i = 0;

                    foreach (string cell in row.Split(','))
                    {
                        dtJiveObjects.Rows[dtJiveObjects.Rows.Count - 1][i] = cell;
                        i++;
                    }
                }
            }

            #endregion

            StreamWriter excelWriterScoringMatrixNew = null;
            excelWriterScoringMatrixNew = System.IO.File.CreateText(textBox2.Text + "\\" + "CategoriesFixReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".csv");
            excelWriterScoringMatrixNew.WriteLine("ObjectID" + "," + "SPURL" + "," + "Tags" + "," + "Status");
            excelWriterScoringMatrixNew.Flush();

            int count = 0;

            string[] SiteSplit = new string[] { "/Lists/" };
            string[] IDSplit = new string[] { "?ID=" };
            string[] DocumentSplit = new string[] { "/Documents/" };
            string[] TagsSplit = new string[] { "|" };
            string[] FileURLSplit = new string[] { "/1_Uploaded Files/" };
            string[] PageURLSplit = new string[] { "/Pages/" };

            foreach (DataRow drImported in dtImportedObjects.Rows)
            {
                count++;
                try
                {
                    string TagsColl = string.Empty;

                    string _SPSpaceID = string.Empty;
                    string _JivePlaceID = string.Empty;

                    string _objectID = string.Empty;
                    string _itemID = string.Empty;
                    string _FilePath = string.Empty;
                    string _objectURL = string.Empty;
                    string _importedURL = string.Empty;
                    string _ListName = string.Empty;
                    string _objList = string.Empty;

                    string _Modified = string.Empty;
                    string _ModifiedBY = string.Empty;

                    _SPSpaceID = drImported["SpaceID"].ToString().Trim();

                    _objectID = drImported["ObjectId"].ToString().Trim();
                    _objectURL = drImported["ImportedURL"].ToString().Trim();
                    _importedURL = drImported["ImportedURL"].ToString().Trim();
                    _objList = drImported["ObjectType"].ToString().Trim();

                    #region OLD TYPE

                    //if (_importedURL.Contains("/1_Uploaded Files/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "1_Uploaded Files";
                    //}
                    //if (_importedURL.Contains("/2_Documents and Pages/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "2_Documents and Pages";
                    //}
                    //if (_importedURL.Contains("/Discussions/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Discussions";
                    //}
                    //if (_importedURL.Contains("/Events/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Events";
                    //}
                    //if (_importedURL.Contains("/Messages/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Messages";
                    //}
                    //if (_importedURL.Contains("/Posts/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Posts";
                    //}
                    //if (_importedURL.Contains("/Site Assets/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Site Assets";
                    //}
                    //if (_importedURL.Contains("/SiteHistory/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "SiteHistory";
                    //}
                    //if (_importedURL.Contains("/Tasks/"))
                    //{
                    //    _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    //    _ListName = "Tasks";
                    //} 

                    #endregion

                    this.Text = (count).ToString() + " of " + dtImportedObjects.Rows.Count.ToString() + " : " + _objectURL;

                    bool itemFound = false;

                    foreach (DataRow drJive in dtJiveObjects.Rows)
                    {
                        //_JivePlaceID = drImported["PlaceID"].ToString().Trim();
                        if ((drImported["ObjectId"].ToString().Trim() == drJive["ID"].ToString().Trim()) && (drImported["ObjectType"].ToString().Trim() == drJive["Type"].ToString().Trim()))
                        {
                            TagsColl = drJive["Tags"].ToString().Trim();
                            _Modified = drJive["Modified"].ToString().Trim();
                            _ModifiedBY = drJive["ModifiedBy"].ToString().Trim();
                            itemFound = true;
                            break;
                        }
                    }

                    if (itemFound && !string.IsNullOrEmpty(TagsColl))
                    {
                        #region Get Site URL

                        if (_importedURL.Contains("/Lists/"))
                        {
                            _importedURL = drImported["ImportedURL"].ToString().Split(SiteSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                        }
                        if (_importedURL.Contains("/1_Uploaded Files/"))
                        {
                            _importedURL = drImported["ImportedURL"].ToString().Split(FileURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                        }
                        if (_importedURL.Contains("/Pages/"))
                        {
                            _importedURL = drImported["ImportedURL"].ToString().Split(PageURLSplit, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                        }

                        #endregion

                        #region Get List and FilePath/ItemID

                        switch (_objList)
                        {
                            case "Document":
                                _ListName = "2_Documents and Pages";
                                _FilePath = drImported["ImportedURL"].ToString().Split(DocumentSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "File":
                                _ListName = "1_Uploaded Files";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Blog":
                                _ListName = "Posts";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Discussion":
                                _ListName = "Discussions";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Event":
                                _ListName = "Events";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Task":
                                _ListName = "Tasks";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;

                            case "Idea":
                                _ListName = "Ideas";
                                _itemID = drImported["ImportedURL"].ToString().Split(IDSplit, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                                break;
                        }

                        #endregion

                        #region Tags Re-Apply

                        AuthenticationManager authManager = new AuthenticationManager();
                        using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(_importedURL, textBox6.Text, textBox5.Text))
                        {
                            Web oWeb = clientcontext.Web;
                            clientcontext.Load(oWeb);
                            clientcontext.ExecuteQuery();

                            List _List = null;

                            try
                            {
                                _List = clientcontext.Web.Lists.GetByTitle(_ListName);
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            { }

                            if (_List != null)
                            {
                                if (_List.Title == "2_Documents and Pages")
                                {
                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    _List.ForceCheckout = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    try
                                    {
                                        clientcontext.Load(_List.RootFolder);
                                        clientcontext.ExecuteQuery();

                                        Folder docFolder = null;

                                        try
                                        {
                                            docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                            clientcontext.Load(docFolder);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        { }

                                        if (docFolder != null)
                                        {
                                            ListItem _Item = docFolder.Files.GetByUrl(_FilePath).ListItemAllFields;
                                            clientcontext.Load(_Item);
                                            clientcontext.ExecuteQuery();

                                            #region Document Modified, ModifiedBy

                                            DateTime Modified = new DateTime();
                                            FieldUserValue ModifiedBy = null;

                                            try
                                            {
                                                if (!string.IsNullOrEmpty(_Modified))
                                                {
                                                    Modified = getdateformat(_Modified);
                                                }
                                                else
                                                {
                                                    Modified = Convert.ToDateTime(_Item["Modified"]);
                                                }

                                                if (!string.IsNullOrEmpty(_ModifiedBY))
                                                {
                                                    User ModifiedUser = default(User);
                                                    try
                                                    {
                                                        ModifiedUser = clientcontext.Web.EnsureUser(_ModifiedBY);
                                                        clientcontext.Load(ModifiedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        ModifiedUser = clientcontext.Web.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                                        clientcontext.Load(ModifiedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }

                                                    ModifiedBy = new FieldUserValue();
                                                    ModifiedBy.LookupId = ModifiedUser.Id;
                                                }
                                                else
                                                {
                                                    ModifiedBy = (FieldUserValue)_Item["Editor"];
                                                }
                                            }
                                            catch (Exception ex)
                                            {

                                            }

                                            #endregion

                                            //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                            //FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                            if (!string.IsNullOrEmpty(TagsColl))
                                            {
                                                string[] _categories = TagsColl.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                                                try
                                                {
                                                    FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                                                    for (int i = 0; i <= _categories.Length - 1; i++)
                                                    {
                                                        string newValue = _categories[i].ToString();

                                                        if (_categories[i].ToString().Contains("$"))
                                                        {
                                                            newValue = _categories[i].ToString().Replace("$", ",");
                                                        }

                                                        int _cId = GetLookupIDsManageTag(newValue, clientcontext, oWeb);

                                                        if (_cId != 0)
                                                        {
                                                            FieldLookupValue flv = new FieldLookupValue();
                                                            flv.LookupId = _cId;

                                                            lookupFieldValCollection.SetValue(flv, i);
                                                        }
                                                    }

                                                    if (lookupFieldValCollection.Length >= 1)
                                                    {
                                                        if (lookupFieldValCollection[0] != null)
                                                            _Item["Tag"] = lookupFieldValCollection;
                                                    }

                                                    _Item.Update();
                                                    clientcontext.Load(_Item);
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "TAgApplyFailure");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }

                                                try
                                                {
                                                    _Item["Modified"] = Modified;
                                                    _Item["Editor"] = ModifiedBy;
                                                    _Item.Update();
                                                    clientcontext.ExecuteQuery();

                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "Success");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                                catch (Exception ex)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ModifyFailure");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        //excelWriterScoringMatrixNew.WriteLine(drImported["DID"].ToString() + "," + drImported["URL"].ToString() + "," + "Failure due to : " + ex.Message);
                                        //excelWriterScoringMatrixNew.Flush();
                                    }

                                    _List.EnableVersioning = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    _List.ForceCheckout = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();
                                }
                                else
                                {
                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    try
                                    {
                                        bool tagsFileldExist = _List.FieldExistsByName("Tag");

                                        if (tagsFileldExist)
                                        {

                                            ListItem _Item = _List.GetItemById(_itemID);
                                            clientcontext.Load(_Item);
                                            clientcontext.ExecuteQuery();

                                            #region Document Modified, ModifiedBy

                                            DateTime Modified = new DateTime();
                                            FieldUserValue ModifiedBy = null;

                                            try
                                            {
                                                if (!string.IsNullOrEmpty(_Modified))
                                                {
                                                    Modified = getdateformat(_Modified);
                                                }
                                                else
                                                {
                                                    Modified = Convert.ToDateTime(_Item["Modified"]);
                                                }

                                                if (!string.IsNullOrEmpty(_ModifiedBY))
                                                {
                                                    User ModifiedUser = default(User);
                                                    try
                                                    {
                                                        ModifiedUser = clientcontext.Web.EnsureUser(_ModifiedBY);
                                                        clientcontext.Load(ModifiedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        ModifiedUser = clientcontext.Web.EnsureUser("RworldAdmin@rsharepoint.onmicrosoft.com");
                                                        clientcontext.Load(ModifiedUser);
                                                        clientcontext.ExecuteQuery();
                                                    }

                                                    ModifiedBy = new FieldUserValue();
                                                    ModifiedBy.LookupId = ModifiedUser.Id;
                                                }
                                                else
                                                {
                                                    ModifiedBy = (FieldUserValue)_Item["Editor"];
                                                }
                                            }
                                            catch (Exception ex)
                                            {

                                            }

                                            #endregion

                                            //DateTime Modified = Convert.ToDateTime(_Item["Modified"]);
                                            //FieldUserValue ModifiedBy = (FieldUserValue)_Item["Editor"];

                                            if (!string.IsNullOrEmpty(TagsColl))
                                            {
                                                string[] _categories = TagsColl.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                                                try
                                                {
                                                    FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[_categories.Length];

                                                    for (int i = 0; i <= _categories.Length - 1; i++)
                                                    {
                                                        string newValue = _categories[i].ToString();

                                                        if (_categories[i].ToString().Contains("$"))
                                                        {
                                                            newValue = _categories[i].ToString().Replace("$", ",");
                                                        }

                                                        int _cId = GetLookupIDsManageTag(newValue, clientcontext, oWeb);

                                                        if (_cId != 0)
                                                        {
                                                            FieldLookupValue flv = new FieldLookupValue();
                                                            flv.LookupId = _cId;

                                                            lookupFieldValCollection.SetValue(flv, i);
                                                        }
                                                    }

                                                    if (lookupFieldValCollection.Length >= 1)
                                                    {
                                                        if (lookupFieldValCollection[0] != null)
                                                            _Item["Tag"] = lookupFieldValCollection;
                                                    }

                                                    _Item.Update();
                                                    clientcontext.Load(_Item);
                                                    clientcontext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "TAgApplyFailure");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }

                                                try
                                                {
                                                    _Item["Modified"] = Modified;
                                                    _Item["Editor"] = ModifiedBy;
                                                    _Item.Update();
                                                    clientcontext.ExecuteQuery();

                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "Success");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                                catch (Exception ex)
                                                {
                                                    excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ModifyFailure");
                                                    excelWriterScoringMatrixNew.Flush();
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception EX)
                                    {
                                        excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ItemIDFailure");
                                        excelWriterScoringMatrixNew.Flush();
                                    }

                                    _List.EnableVersioning = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();
                                }
                            }
                        }

                        #endregion
                    }
                    else
                    {
                        excelWriterScoringMatrixNew.WriteLine(_objectID + "," + _objectURL + "," + TagsColl + "," + "ItemIDNotFoundinJive");
                        excelWriterScoringMatrixNew.Flush();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterScoringMatrixNew.Flush();
            excelWriterScoringMatrixNew.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button73_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig

            List<string> ListNames = new List<string>();

            //ListNames.Add("Team Files");
            //ListNames.Add("Team Files");
            //ListNames.Add("Uploaded Files");
            //ListNames.Add("Documents and Pages");
            ListNames.Add("1_Uploaded Files");
            ListNames.Add("2_Documents and Pages");
            ListNames.Add("Discussions");
            ListNames.Add("Events");
            ListNames.Add("Messages");
            ListNames.Add("Posts");
            ListNames.Add("Site Assets");
            //ListNames.Add("SiteHistory");
            ListNames.Add("Tasks");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {

                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), textBox6.Text, textBox5.Text))
                    {
                        Web _Web = _cContext.Web;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        _cContext.RequestTimeout = -1;

                        StreamWriter excelWriterTagsReport = null;
                        excelWriterTagsReport = System.IO.File.CreateText(textBox2.Text + "\\" + _Web.Id.ToString() + "_TagsApplyReport" + ".csv");
                        excelWriterTagsReport.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "Tags");
                        excelWriterTagsReport.Flush();

                        ListCollection oListColl = _cContext.Web.Lists;
                        _cContext.Load(oListColl);
                        _cContext.ExecuteQuery();

                        bool ManageTagtagsListExist = _Web.ListExists("Manage Tag");

                        if (!ManageTagtagsListExist)
                        {
                            try
                            {
                                ListCreationInformation creationInfo = new ListCreationInformation();
                                creationInfo.Title = "Manage Tag";
                                creationInfo.Description = "Manage Tag";
                                creationInfo.TemplateType = (int)ListTemplateType.GenericList;
                                List newList = _cContext.Web.Lists.Add(creationInfo);
                                _cContext.Load(newList);
                                _cContext.ExecuteQuery();

                                try
                                {
                                    List newList1 = _cContext.Web.Lists.GetByTitle("Manage Tag");
                                    _cContext.Load(newList1);
                                    _cContext.ExecuteQuery();

                                    FieldCollection oFieldColl = newList1.Fields;
                                    _cContext.Load(oFieldColl);
                                    _cContext.ExecuteQuery();

                                    Field field = oFieldColl.GetByTitle("Title");
                                    _cContext.Load(field);
                                    _cContext.ExecuteQuery();

                                    field.Indexed = true;
                                    field.Update();
                                    _cContext.ExecuteQuery();

                                    field.EnforceUniqueValues = true;
                                    field.Update();
                                    _cContext.ExecuteQuery();
                                }
                                catch (Exception ec1)
                                {

                                }
                            }
                            catch (Exception ex)
                            {
                            }
                        }

                        foreach (List _List in oListColl)
                        {
                            _cContext.Load(_List);
                            _cContext.ExecuteQuery();

                            _cContext.Load(_List.RootFolder);
                            _cContext.ExecuteQuery();

                            string listPath = _List.RootFolder.ServerRelativeUrl;

                            if (listPath.ToLower().EndsWith("/pages") || listPath.ToLower().EndsWith("/1_uploaded files"))
                            // if (listPath.ToLower().EndsWith("/1_uploaded files"))
                            {
                                try
                                {
                                    _cContext.Load(_List);
                                    _cContext.ExecuteQuery();

                                    bool tagsFileldExist = _List.FieldExistsByName("Tag");

                                    if (!tagsFileldExist)
                                    {
                                        try
                                        {
                                            List list = _Web.Lists.GetByTitle("Manage Tag");
                                            _cContext.Load(list);
                                            _cContext.ExecuteQuery();
                                            string schemaLookupField = "<Field Type='LookupMulti' Name='Tag' StaticName='Tag' DisplayName='Tag' List = '" + list.Id + "' ShowField = 'Title' Mult = 'TRUE'/>";
                                            Field lookupField = _List.Fields.AddFieldAsXml(schemaLookupField, false, AddFieldOptions.AddFieldInternalNameHint);
                                            _List.Update();
                                            _cContext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    }

                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    _cContext.ExecuteQuery();

                                    if (listPath.ToLower().EndsWith("/pages"))
                                    {
                                        try
                                        {
                                            _List.ForceCheckout = false;
                                            _List.Update();
                                            _cContext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    }

                                    bool tagsFileldExist1 = _List.FieldExistsByName("Tags");
                                    bool NewtagFileldExist1 = _List.FieldExistsByName("Tag");

                                    if (tagsFileldExist1 && NewtagFileldExist1)
                                    {

                                        #region NEW

                                        //ListItemCollectionPosition itemPosition = null;
                                        //while (true)
                                        //{
                                        //    CamlQuery camlQuery = new CamlQuery();
                                        //    camlQuery.ListItemCollectionPosition = itemPosition;
                                        //    camlQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit>500</RowLimit></View>";
                                        //    ListItemCollection collListItem = _List.GetItems(camlQuery);
                                        //    _cContext.Load(collListItem);
                                        //    _cContext.ExecuteQuery();

                                        //    itemPosition = collListItem.ListItemCollectionPosition;

                                        //    foreach (ListItem oItem in collListItem)
                                        //    {
                                        //        try
                                        //        {
                                        //            string Tags = string.Empty;

                                        //            _cContext.Load(oItem);
                                        //            _cContext.ExecuteQuery();

                                        //            string itemID = oItem.Id.ToString();

                                        //            DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                                        //            FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                                        //            TaxonomyFieldValueCollection taxFieldValues = oItem["Tags"] as TaxonomyFieldValueCollection;

                                        //            if (taxFieldValues.Count > 0)
                                        //            {
                                        //                FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[taxFieldValues.Count];
                                        //                int i = 0;

                                        //                foreach (TaxonomyFieldValue tv in taxFieldValues)
                                        //                {
                                        //                    int _cId = GetLookupIDsManageTag(tv.Label.ToString(), _cContext, _Web);

                                        //                    if (_cId != 0)
                                        //                    {
                                        //                        Tags += tv.Label.ToString() + "|";

                                        //                        FieldLookupValue flv = new FieldLookupValue();
                                        //                        flv.LookupId = _cId;
                                        //                        lookupFieldValCollection.SetValue(flv, i);
                                        //                        i++;
                                        //                    }
                                        //                }

                                        //                if (lookupFieldValCollection.Length >= 1)
                                        //                {
                                        //                    if (lookupFieldValCollection[0] != null)
                                        //                        oItem["Tag"] = lookupFieldValCollection;

                                        //                    oItem.Update();
                                        //                    _cContext.Load(oItem);
                                        //                    _cContext.ExecuteQuery();

                                        //                    try
                                        //                    {
                                        //                        oItem["Modified"] = Modified;
                                        //                        oItem["Editor"] = ModifiedBy;
                                        //                        oItem.Update();
                                        //                        _cContext.ExecuteQuery();
                                        //                    }
                                        //                    catch (Exception ex)
                                        //                    {

                                        //                    }

                                        //                    excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + itemID + "," + Tags);
                                        //                    excelWriterTagsReport.Flush();
                                        //                }
                                        //            }
                                        //        }
                                        //        catch (Exception ex)
                                        //        {
                                        //            continue;
                                        //        }
                                        //    }

                                        //    if (itemPosition == null)
                                        //    {
                                        //        break;
                                        //    }

                                        //    //Console.WriteLine("\n" + itemPosition.PagingInfo + "\n");
                                        //}

                                        #endregion

                                        #region OTHER METHOD  

                                        int start = 0;
                                        int end = 0;

                                        if (listPath.ToLower().EndsWith("/1_uploaded files"))
                                        {
                                            start = 3621;
                                            end = 5500;
                                        }

                                        if (listPath.ToLower().EndsWith("/pages"))
                                        {
                                            start = 0;
                                            end = 50;
                                        }

                                        for (int w = start; w <= end; w++)
                                        {
                                            this.Text = start + "*to*" + end + " @" + (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + "; Item Count :" + w.ToString() + " -- " + lstSiteColl[j].ToString();

                                            ListItem oItem = null;

                                            try
                                            {
                                                oItem = _List.GetItemById(w.ToString());
                                                _cContext.Load(oItem);
                                                _cContext.ExecuteQuery();
                                            }
                                            catch (Exception ex)
                                            {
                                                continue;
                                            }

                                            if (oItem != null)
                                            {
                                                try
                                                {
                                                    string Tags = string.Empty;

                                                    _cContext.Load(oItem);
                                                    _cContext.ExecuteQuery();

                                                    string itemID = oItem.Id.ToString();

                                                    DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                                                    FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                                                    TaxonomyFieldValueCollection taxFieldValues = oItem["Tags"] as TaxonomyFieldValueCollection;

                                                    if (taxFieldValues.Count > 0)
                                                    {
                                                        FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[taxFieldValues.Count];
                                                        int i = 0;

                                                        foreach (TaxonomyFieldValue tv in taxFieldValues)
                                                        {
                                                            int _cId = GetLookupIDsManageTag(tv.Label.ToString(), _cContext, _Web);

                                                            if (_cId != 0)
                                                            {
                                                                Tags += tv.Label.ToString() + "|";

                                                                FieldLookupValue flv = new FieldLookupValue();
                                                                flv.LookupId = _cId;
                                                                lookupFieldValCollection.SetValue(flv, i);
                                                                i++;
                                                            }
                                                        }

                                                        if (lookupFieldValCollection.Length >= 1)
                                                        {
                                                            if (lookupFieldValCollection[0] != null)
                                                                oItem["Tag"] = lookupFieldValCollection;

                                                            oItem.Update();
                                                            _cContext.Load(oItem);
                                                            _cContext.ExecuteQuery();

                                                            try
                                                            {
                                                                oItem["Modified"] = Modified;
                                                                oItem["Editor"] = ModifiedBy;
                                                                oItem.Update();
                                                                _cContext.ExecuteQuery();
                                                            }
                                                            catch (Exception ex)
                                                            {

                                                            }

                                                            excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + itemID + "," + Tags);
                                                            excelWriterTagsReport.Flush();
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    continue;
                                                }
                                            }
                                        }

                                        #endregion

                                        #region OLD

                                        //CamlQuery camlQuery = new CamlQuery();
                                        //camlQuery.ViewXml = "<View><RowLimit>500</RowLimit></View>";//<View Scope='RecursiveAll'></View>//<RowLimit>5000</RowLimit>

                                        //ListItemCollection listItems = _List.GetItems(camlQuery);
                                        //_cContext.Load(listItems);
                                        //_cContext.ExecuteQuery();

                                        //foreach (ListItem oItem in listItems)
                                        //{
                                        //    try
                                        //    {
                                        //        string Tags = string.Empty;

                                        //        _cContext.Load(oItem);
                                        //        _cContext.ExecuteQuery();

                                        //        string itemID = oItem.Id.ToString();

                                        //        DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                                        //        FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                                        //        TaxonomyFieldValueCollection taxFieldValues = oItem["Tags"] as TaxonomyFieldValueCollection;

                                        //        if (taxFieldValues.Count > 0)
                                        //        {
                                        //            FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[taxFieldValues.Count];
                                        //            int i = 0;

                                        //            foreach (TaxonomyFieldValue tv in taxFieldValues)
                                        //            {
                                        //                int _cId = GetLookupIDsManageTag(tv.Label.ToString(), _cContext, _Web);

                                        //                if (_cId != 0)
                                        //                {
                                        //                    Tags += tv.Label.ToString() + "|";

                                        //                    FieldLookupValue flv = new FieldLookupValue();
                                        //                    flv.LookupId = _cId;
                                        //                    lookupFieldValCollection.SetValue(flv, i);
                                        //                    i++;
                                        //                }
                                        //            }

                                        //            if (lookupFieldValCollection.Length >= 1)
                                        //            {
                                        //                if (lookupFieldValCollection[0] != null)
                                        //                    oItem["Tag"] = lookupFieldValCollection;

                                        //                oItem.Update();
                                        //                _cContext.Load(oItem);
                                        //                _cContext.ExecuteQuery();

                                        //                try
                                        //                {
                                        //                    oItem["Modified"] = Modified;
                                        //                    oItem["Editor"] = ModifiedBy;
                                        //                    oItem.Update();
                                        //                    _cContext.ExecuteQuery();
                                        //                }
                                        //                catch (Exception ex)
                                        //                {

                                        //                }

                                        //                excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + itemID + "," + Tags);
                                        //                excelWriterTagsReport.Flush();
                                        //            }
                                        //        }
                                        //    }
                                        //    catch (Exception ex)
                                        //    {
                                        //        continue;
                                        //    }
                                        //} 

                                        #endregion
                                    }

                                    _List.EnableVersioning = true;
                                    _List.Update();
                                    _cContext.ExecuteQuery();

                                    if (listPath.ToLower().EndsWith("/pages"))
                                    {
                                        try
                                        {
                                            _List.ForceCheckout = true;
                                            _List.Update();
                                            _cContext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + "ERROR" + "," + ex.Message.Replace(",", ""));
                                    excelWriterTagsReport.Flush();
                                    continue;
                                }
                            }
                        }



                        excelWriterTagsReport.Flush();
                        excelWriterTagsReport.Close();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }


            #endregion

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button74_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            #region Remainig   

            StreamWriter excelWriterTagsReport = null;
            excelWriterTagsReport = System.IO.File.CreateText(textBox2.Text + "\\" + "_TagSolutionsIsuueReport" + ".csv");
            excelWriterTagsReport.WriteLine("SiteURL" + "," + "Tags");
            excelWriterTagsReport.Flush();


            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var _cContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "svc-jivemigration@rsharepoint.onmicrosoft.com", "Lot62215"))
                    {
                        Web _Web = _cContext.Web;
                        _cContext.Load(_Web);
                        _cContext.ExecuteQuery();

                        _cContext.RequestTimeout = -1;

                        ListCollection oListColl = _cContext.Web.Lists;
                        _cContext.Load(oListColl);
                        _cContext.ExecuteQuery();

                        bool ManageTagtagsListExist = _Web.ListExists("Manage Tag");

                        if (ManageTagtagsListExist)
                        {
                            try
                            {
                                List newList = _cContext.Web.Lists.GetByTitle("Manage Tag");
                                _cContext.Load(newList);
                                _cContext.ExecuteQuery();

                                CamlQuery camlQuery = new CamlQuery();
                                camlQuery.ViewXml = "<View><Query><Where><Or><Eq><FieldRef Name='Title'/><Value Type='Text'>solutions</Value></Eq><Eq><FieldRef Name='Title'/><Value Type='Text'>samesame</Value></Eq></Or></Where></Query></View>";
                                ListItemCollection collListItem = newList.GetItems(camlQuery);
                                _cContext.Load(collListItem);
                                _cContext.ExecuteQuery();

                                var totalListItems = collListItem.Count;

                                if (totalListItems > 0)
                                {
                                    for (var counter = totalListItems - 1; counter > -1; counter--)
                                    {
                                        collListItem[counter].DeleteObject();
                                        _cContext.ExecuteQuery();
                                    }

                                    excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + totalListItems.ToString());
                                    excelWriterTagsReport.Flush();
                                }
                            }
                            catch (Exception eq)
                            {
                            }
                        }

                        #region EXTRA CODE

                        //if (solutionsTagExist)
                        //{
                        //    foreach (List _List in oListColl)
                        //    {
                        //        _cContext.Load(_List);
                        //        _cContext.ExecuteQuery();

                        //        //_List.EnableVersioning = false;
                        //        //_List.Update();
                        //        //_cContext.ExecuteQuery();

                        //        try
                        //        {
                        //            bool tagsFileldExist = _List.FieldExistsByName("Tag");

                        //            if (tagsFileldExist)
                        //            {
                        //                //_List.EnableVersioning = false;
                        //                //_List.Update();
                        //                //_cContext.ExecuteQuery();

                        //                #region NEW

                        //                try
                        //                {
                        //                    ListItemCollectionPosition itemPosition = null;

                        //                    while (true)
                        //                    {
                        //                        CamlQuery camlQuery = new CamlQuery();
                        //                        camlQuery.ListItemCollectionPosition = itemPosition;
                        //                        //camlQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit>500</RowLimit></View>";
                        //                        camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Tag'/><Value Type='Lookup'>India</Value></Eq></Where></Query></View>";
                        //                        ListItemCollection collListItem = _List.GetItems(camlQuery);
                        //                        _cContext.Load(collListItem);
                        //                        _cContext.ExecuteQuery();

                        //                        itemPosition = collListItem.ListItemCollectionPosition;

                        //                        foreach (ListItem oItem in collListItem)
                        //                        {
                        //                            try
                        //                            {
                        //                                FieldLookupValue[] values = oItem["Tag"] as FieldLookupValue[];

                        //                                if (values.Count() == 1)
                        //                                {
                        //                                    foreach (FieldLookupValue value in values)
                        //                                    {
                        //                                        if (value.LookupValue.ToString().ToLower() == "solutions" || value.LookupValue.ToString().ToLower() == "samesame")
                        //                                        {
                        //                                            string TagValue = value.LookupValue.ToString();

                        //                                            try
                        //                                            {
                        //                                                #region Delete Tag

                        //                                                //_cContext.Load(oItem);
                        //                                                //_cContext.ExecuteQuery();

                        //                                                //string itemID = oItem.Id.ToString();

                        //                                                //DateTime Modified = Convert.ToDateTime(oItem["Modified"]);
                        //                                                //FieldUserValue ModifiedBy = (FieldUserValue)oItem["Editor"];

                        //                                                //TaxonomyFieldValueCollection taxFieldValues = oItem["Tags"] as TaxonomyFieldValueCollection;

                        //                                                //if (taxFieldValues.Count > 0)
                        //                                                //{
                        //                                                //    FieldLookupValue[] lookupFieldValCollection = new FieldLookupValue[taxFieldValues.Count];
                        //                                                //    int i = 0;

                        //                                                //    foreach (TaxonomyFieldValue tv in taxFieldValues)
                        //                                                //    {
                        //                                                //        int _cId = GetLookupIDsManageTag(tv.Label.ToString(), _cContext, _Web);

                        //                                                //        if (_cId != 0)
                        //                                                //        {
                        //                                                //            Tags += tv.Label.ToString() + "|";

                        //                                                //            FieldLookupValue flv = new FieldLookupValue();
                        //                                                //            flv.LookupId = _cId;
                        //                                                //            lookupFieldValCollection.SetValue(flv, i);
                        //                                                //            i++;
                        //                                                //        }
                        //                                                //    }

                        //                                                //    if (lookupFieldValCollection.Length >= 1)
                        //                                                //    {
                        //                                                //        if (lookupFieldValCollection[0] != null)
                        //                                                //            oItem["Tag"] = lookupFieldValCollection;

                        //                                                //        oItem.Update();
                        //                                                //        _cContext.Load(oItem);
                        //                                                //        _cContext.ExecuteQuery();

                        //                                                //        try
                        //                                                //        {
                        //                                                //            oItem["Modified"] = Modified;
                        //                                                //            oItem["Editor"] = ModifiedBy;
                        //                                                //            oItem.Update();
                        //                                                //            _cContext.ExecuteQuery();
                        //                                                //        }
                        //                                                //        catch (Exception ex)
                        //                                                //        {

                        //                                                //        }

                        //                                                //        excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + itemID + "," + Tags);
                        //                                                //        excelWriterTagsReport.Flush();
                        //                                                //    }

                        //                                                #endregion

                        //                                                excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + oItem.Id.ToString() + "," + TagValue);
                        //                                                excelWriterTagsReport.Flush();
                        //                                            }
                        //                                            catch (Exception es)
                        //                                            {
                        //                                                excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + oItem.Id.ToString() + "," + TagValue);
                        //                                                excelWriterTagsReport.Flush();
                        //                                            }
                        //                                        }
                        //                                    }
                        //                                }
                        //                            }
                        //                            catch (Exception ex)
                        //                            {
                        //                                continue;
                        //                            }
                        //                        }

                        //                        if (itemPosition == null)
                        //                        {
                        //                            break;
                        //                        }
                        //                    }
                        //                }
                        //                catch (Exception exc)
                        //                {

                        //                }

                        //                #endregion

                        //                //_List.EnableVersioning = true;
                        //                //_List.Update();
                        //                //_cContext.ExecuteQuery();
                        //            }
                        //        }
                        //        catch (Exception ex)
                        //        {
                        //            excelWriterTagsReport.WriteLine(_cContext.Web.Url + "," + _List.Title.ToString() + "," + "ERROR" + "," + ex.Message.Replace(",", ""));
                        //            excelWriterTagsReport.Flush();
                        //            continue;
                        //        }
                        //    }
                        //} 

                        #endregion

                        excelWriterTagsReport.Flush();
                        excelWriterTagsReport.Close();
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }


            #endregion

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button75_Click(object sender, EventArgs e)
        {

            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion  

            StreamWriter excelWriterDateIssueReport = null;
            excelWriterDateIssueReport = System.IO.File.CreateText(textBox2.Text + "\\" + "_DateIssueReport" + ".csv");
            excelWriterDateIssueReport.WriteLine("SiteURL" + "," + "ListName" + "," + "ItemID" + "," + "CreatedChange" + "," + "ModifiedChange");
            excelWriterDateIssueReport.Flush();

            List<string> ListNames = new List<string>();

            ListNames.Add("1_Uploaded Files");
            ListNames.Add("2_Documents and Pages");
            ListNames.Add("Discussions");
            ListNames.Add("Events");
            ListNames.Add("Messages");
            ListNames.Add("Posts");
            ListNames.Add("SiteHistory");
            ListNames.Add("Tasks");
            ListNames.Add("Comments");

            for (int j = 0; j <= lstSiteColl.Count - 1; j++)
            {
                this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                try
                {
                    AuthenticationManager authManager = new AuthenticationManager();
                    using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant(lstSiteColl[j].ToString(), "infosys.service1@RICOHEUROPEPLC.onmicrosoft.com", "NByH2354"))
                    {
                        Web oWeb = clientcontext.Web;
                        clientcontext.Load(oWeb);
                        clientcontext.ExecuteQuery();

                        ListCollection oLists = clientcontext.Web.Lists;
                        clientcontext.Load(oLists);
                        clientcontext.ExecuteQuery();

                        foreach (string _List1 in ListNames)
                        {
                            try
                            {
                                List _List = clientcontext.Web.Lists.GetByTitle(_List1);
                                clientcontext.Load(_List);
                                clientcontext.ExecuteQuery();

                                if (_List.Title == "2_Documents and Pages")
                                {
                                    string CreatedChange = "No";
                                    string ModifiedChange = "No";

                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    _List.ForceCheckout = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    try
                                    {
                                        clientcontext.Load(_List.RootFolder);
                                        clientcontext.ExecuteQuery();

                                        Folder docFolder = null;

                                        try
                                        {
                                            docFolder = _List.RootFolder.Folders.GetByUrl("Documents");
                                            clientcontext.Load(docFolder);
                                            clientcontext.ExecuteQuery();
                                        }
                                        catch (Exception ex)
                                        { }

                                        if (docFolder != null)
                                        {
                                            FileCollection oItemColl = docFolder.Files;
                                            clientcontext.Load(oItemColl);
                                            clientcontext.ExecuteQuery();

                                            foreach (Microsoft.SharePoint.Client.File _FileItem in oItemColl)
                                            {
                                                try
                                                {
                                                    ListItem _oItem = _FileItem.ListItemAllFields;
                                                    clientcontext.Load(_oItem);
                                                    clientcontext.ExecuteQuery();

                                                    string itemID = _oItem.Id.ToString();

                                                    //if (itemID != "83")
                                                    {

                                                        DateTime Modified = Convert.ToDateTime(_oItem["Modified"]);
                                                        FieldUserValue ModifiedBy = (FieldUserValue)_oItem["Editor"];
                                                        DateTime Created = Convert.ToDateTime(_oItem["Created"]);
                                                        FieldUserValue CreatedBy = (FieldUserValue)_oItem["Author"];

                                                        //DateTime dtModified = Checkdateformat(Modified);

                                                        DateTime OriginalModifiedDate = Modified;
                                                        int imortedModifiedMonth = Modified.Month;
                                                        int imortedModifiedDay = Modified.Day;
                                                        int imortedModifiedYear = Modified.Year;

                                                        if (imortedModifiedDay <= 12)
                                                        {
                                                            string actualFormat = Modified.Day.ToString() + "/" + Modified.Month.ToString() + "/" + Modified.Year.ToString() + " " + Modified.TimeOfDay.ToString();
                                                            OriginalModifiedDate = getdateformat(actualFormat);

                                                            ModifiedChange = "Yes";
                                                        }

                                                        //DateTime dtCreated = Checkdateformat(Created);

                                                        DateTime OriginalCreatedDate = Created;
                                                        int imortedCreatedMonth = Created.Month;
                                                        int imortedCreatedDay = Created.Day;
                                                        int imortedCreatedYear = Created.Year;

                                                        if (imortedCreatedDay <= 12)
                                                        {
                                                            string actualFormat = Created.Day.ToString() + "/" + Created.Month.ToString() + "/" + Created.Year.ToString() + " " + Created.TimeOfDay.ToString();
                                                            OriginalCreatedDate = getdateformat(actualFormat);

                                                            CreatedChange = "Yes";
                                                        }

                                                        _oItem["Created"] = OriginalCreatedDate;
                                                        _oItem["Author"] = CreatedBy;
                                                        _oItem["Modified"] = OriginalModifiedDate;
                                                        _oItem["Editor"] = ModifiedBy;
                                                        _oItem.Update();
                                                        clientcontext.ExecuteQuery();

                                                        if (CreatedChange == "Yes" || ModifiedChange == "Yes")
                                                        {
                                                            excelWriterDateIssueReport.WriteLine(clientcontext.Web.Url + "," + _List.Title.ToString() + "," + itemID + "," + CreatedChange + "," + ModifiedChange);
                                                            excelWriterDateIssueReport.Flush();
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                    }

                                    _List.EnableVersioning = true;
                                    _List.ForceCheckout = true;
                                    _List.ContentTypesEnabled = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();
                                }
                                else
                                {
                                    string CreatedChange = "No";
                                    string ModifiedChange = "No";

                                    _List.EnableVersioning = false;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();

                                    CamlQuery camlQuery = new CamlQuery();
                                    camlQuery.ViewXml = "<View></View>";
                                    ListItemCollection collListItem = _List.GetItems(camlQuery);
                                    clientcontext.Load(collListItem);
                                    clientcontext.ExecuteQuery();

                                    foreach (ListItem _oItem in collListItem)
                                    {
                                        try
                                        {
                                            clientcontext.Load(_oItem);
                                            clientcontext.ExecuteQuery();

                                            string itemID = _oItem.Id.ToString();

                                            //if (itemID != "4")
                                            {
                                                DateTime Modified = Convert.ToDateTime(_oItem["Modified"]);
                                                FieldUserValue ModifiedBy = (FieldUserValue)_oItem["Editor"];
                                                DateTime Created = Convert.ToDateTime(_oItem["Created"]);
                                                FieldUserValue CreatedBy = (FieldUserValue)_oItem["Author"];

                                                //DateTime dtModified = Checkdateformat(Modified);

                                                DateTime OriginalModifiedDate = Modified;
                                                int imortedModifiedMonth = Modified.Month;
                                                int imortedModifiedDay = Modified.Day;
                                                int imortedModifiedYear = Modified.Year;

                                                if (imortedModifiedDay <= 12)
                                                {
                                                    string actualFormat = Modified.Day.ToString() + "/" + Modified.Month.ToString() + "/" + Modified.Year.ToString() + " " + Modified.TimeOfDay.ToString();
                                                    OriginalModifiedDate = getdateformat(actualFormat);

                                                    ModifiedChange = "Yes";
                                                }

                                                //DateTime dtCreated = Checkdateformat(Created);

                                                DateTime OriginalCreatedDate = Created;
                                                int imortedCreatedMonth = Created.Month;
                                                int imortedCreatedDay = Created.Day;
                                                int imortedCreatedYear = Created.Year;

                                                if (imortedCreatedDay <= 12)
                                                {
                                                    string actualFormat = Created.Day.ToString() + "/" + Created.Month.ToString() + "/" + Created.Year.ToString() + " " + Created.TimeOfDay.ToString();
                                                    OriginalCreatedDate = getdateformat(actualFormat);

                                                    CreatedChange = "Yes";
                                                }

                                                _oItem["Created"] = OriginalCreatedDate;
                                                _oItem["Author"] = CreatedBy;
                                                _oItem["Modified"] = OriginalModifiedDate;
                                                _oItem["Editor"] = ModifiedBy;
                                                _oItem.Update();
                                                clientcontext.ExecuteQuery();

                                                if (CreatedChange == "Yes" || ModifiedChange == "Yes")
                                                {
                                                    excelWriterDateIssueReport.WriteLine(clientcontext.Web.Url + "," + _List.Title.ToString() + "," + itemID + "," + CreatedChange + "," + ModifiedChange);
                                                    excelWriterDateIssueReport.Flush();
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            continue;
                                        }
                                    }

                                    _List.EnableVersioning = true;
                                    _List.Update();
                                    clientcontext.ExecuteQuery();
                                }
                            }
                            catch (Exception ew)
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            excelWriterDateIssueReport.Flush();
            excelWriterDateIssueReport.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }

        private void button76_Click(object sender, EventArgs e)
        {
            #region Site Collection URLS CSV Reading

            List<string> lstSiteColl = new List<string>();
            StreamReader sr = new StreamReader(System.IO.File.OpenRead(textBox1.Text));

            while (!sr.EndOfStream)
            {
                try
                {
                    lstSiteColl.Add(sr.ReadLine().Trim());
                }
                catch
                {
                    continue;
                }
            }

            #endregion

            StreamWriter excelWriterTagsCleanup = null;
            excelWriterTagsCleanup = System.IO.File.CreateText(textBox2.Text + "\\" + "TagsTermCleanUpReport" + ".csv");
            excelWriterTagsCleanup.WriteLine("Term" + "," + "TermID" + "," + "Status");
            excelWriterTagsCleanup.Flush();

            AuthenticationManager authManager = new AuthenticationManager();
            using (var clientcontext = authManager.GetSharePointOnlineAuthenticatedContextTenant("https://rsharepoint.sharepoint.com/sites/rworld", textBox6.Text, textBox5.Text))
            {
                try
                {
                    // Get the TaxonomySession
                    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientcontext);

                    // Get the term store by name
                    TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_3uoEd4FJufp7hiqHvWFqhw==");

                    // Get the term group by Name
                    TermGroup termGroup = termStore.Groups.GetByName("RicohTags");

                    // Get the term set by Name
                    TermSet termSet = termGroup.TermSets.GetByName("TagsTermSet");

                    // Get all the terms 
                    TermCollection termColl = termSet.Terms;
                    clientcontext.Load(termColl);
                    clientcontext.ExecuteQuery();

                    for (int j = 0; j <= lstSiteColl.Count - 1; j++)
                    {
                        this.Text = (j + 1).ToString() + " of " + (lstSiteColl.Count).ToString() + " : " + lstSiteColl[j].ToString();

                        try
                        {
                            Term tm = termColl.GetByName(lstSiteColl[j].ToString());
                            clientcontext.Load(tm);
                            clientcontext.ExecuteQuery();

                            string tmID = tm.Id.ToString();

                            try
                            {
                                tm.DeleteObject();
                                termStore.CommitAll();
                                clientcontext.ExecuteQuery();

                                excelWriterTagsCleanup.WriteLine(lstSiteColl[j].ToString() + "," + tmID + "," + "Deleted");
                                excelWriterTagsCleanup.Flush();
                            }
                            catch (Exception er)
                            {
                                excelWriterTagsCleanup.WriteLine(lstSiteColl[j].ToString() + "," + tmID + "," + "NotDeleted");
                                excelWriterTagsCleanup.Flush();
                            }

                        }
                        catch (Exception ex)
                        {
                            string mess = ex.Message.Replace("\n\n", "");
                            excelWriterTagsCleanup.WriteLine(lstSiteColl[j].ToString() + "," + "--" + "," + "TermNotFound");
                            excelWriterTagsCleanup.Flush();

                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                }
            }

            excelWriterTagsCleanup.Flush();
            excelWriterTagsCleanup.Close();

            this.Text = "Process completed successfully.";
            MessageBox.Show("Process Completed");
        }
    }
}
