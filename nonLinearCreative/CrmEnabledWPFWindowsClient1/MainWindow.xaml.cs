using CrmEnabledWPFWindowsClient1.LoginWindow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using nonLinear;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Reflection;

namespace CrmEnabledWPFWindowsClient1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Button to login to CRM and create a CrmService Client 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            //Load Excel
            #region Load Excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;

            Marketing oMarketing = new Marketing();
            List<Marketing> oListMarketing = new List<Marketing>();

            Excel.Range range;
            workbook = excelApp.Workbooks.Open("C:\\Users\\Sana\\Desktop\\NoNLinear\\Book1.xlsx");
            worksheet = (Excel.Worksheet)workbook.Sheets["ENTERPRISE WEEKLY CRM IMPORTS"];

            int column = 0;
            int row = 0;

            range = worksheet.UsedRange;
            DataTable dt = new DataTable();
            dt.Columns.Add("Project Name");
            dt.Columns.Add("Division");
            dt.Columns.Add("Lead Source");
            dt.Columns.Add("Social Network");
            dt.Columns.Add("Interaction Type");
            dt.Columns.Add("Content");
            dt.Columns.Add("First Name");
            dt.Columns.Add("Last Name");
            dt.Columns.Add("Email");
            dt.Columns.Add("Company Name");
            dt.Columns.Add("Total site visits");
            dt.Columns.Add("Num of visits");
            dt.Columns.Add("Three visitor");
            dt.Columns.Add("Seven visitor");
            dt.Columns.Add("Job Title");
            dt.Columns.Add("Function");
            dt.Columns.Add("Seniority");
            dt.Columns.Add("Industry");
            dt.Columns.Add("Add to book list");
            dt.Columns.Add("City");
            dt.Columns.Add("Province");
            dt.Columns.Add("Street");
            dt.Columns.Add("Date Added");
            dt.Columns.Add("Postal Code");
            dt.Columns.Add("Notes");
            dt.Columns.Add("Interaction Date");
            dt.Columns.Add("LinkedIn Lead Profile URL");
            dt.Columns.Add("Owner");
            dt.Columns.Add("Created by");
            dt.Columns.Add("CMS");
            dt.Columns.Add("Add to LI campaign");
            dt.Columns.Add("Sent Book");
            dt.Columns.Add("Book Return Reason");
            dt.Columns.Add("Book Sent Date");
            dt.Columns.Add("Extra Books Sent");
            dt.Columns.Add("Website");
            dt.Columns.Add("Sitecore User");

            for (row = 2; row <= range.Rows.Count; row++)
            {
                DataRow dr = dt.NewRow();
                for (column = 1; column <= range.Columns.Count; column++)
                {
                    if (column == 38)
                        break;
                    try
                    {
                        if (column == 24 || column == 27 || column == 35)
                        {
                            object value = (range.Cells[row, column] as Excel.Range).Value2;
                            DateTime dts = DateTime.Now;

                            if (value != null)
                            {
                                if (value is double)
                                {
                                    dts = DateTime.FromOADate((double)value);
                                }
                                else
                                {
                                    DateTime.TryParse((string)value, out dts);
                                }
                            }
                            //  MessageBox.Show(dts.ToString());
                            dr[column - 1] = dts.ToString();

                        }
                        else if (column == 11 || column == 12)
                        {
                            object value = (range.Cells[row, column] as Excel.Range).Value2;
                            double d;
                            if (value != null)
                            {
                                if (value is double)
                                {
                                    d = ((double)value);
                                    dr[column - 1] = d.ToString();
                                    //        MessageBox.Show(d.ToString());
                                }
                            }
                            else
                            {
                                dr[column - 1] = string.Empty;
                            }
                        }
                        else if (column == 25)
                        {
                            object value = (range.Cells[row, column] as Excel.Range).Value2;
                            double d;
                            if (value != null)
                            {
                                if (value is double)
                                {
                                    d = ((double)value);
                                    dr[column - 1] = d.ToString();
                                    //          MessageBox.Show(d.ToString());
                                }
                            }
                            else
                            {
                                dr[column - 1] = (range.Cells[row, column] as Excel.Range).Value2;
                            }

                        }
                        else
                        {
                            //    MessageBox.Show((range.Cells[row, column] as Excel.Range).Value2);
                            dr[column - 1] = (range.Cells[row, column] as Excel.Range).Value2;

                        }
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.Message);

                    }

                }

                dt.Rows.Add(dr);
                Console.Write(dt.Rows.Count + "\n");
                dt.AcceptChanges();

                if (dt.Rows.Count == 200)
                    break;
            }
            workbook.Close(true, Missing.Value, Missing.Value);
            excelApp.Quit();

            int i = 0;
            foreach (DataRow rows in dt.Rows)
            {
                oMarketing = new Marketing();
                //Project Name
                oMarketing.ProjectName = dt.Rows[i][0].ToString();
                //"Division"
                oMarketing.Division = dt.Rows[i][1].ToString();
                //"Lead Source"
                oMarketing.LeadSource = dt.Rows[i][2].ToString();
                //"Social Network"
                oMarketing.SocialNetwork = dt.Rows[i][3].ToString();
                //"Interaction Type"
                oMarketing.InteractionType = dt.Rows[i][4].ToString();
                //"Content"
                oMarketing.Content = dt.Rows[i][5].ToString();
                //"First Name"
                oMarketing.FirstName = dt.Rows[i][6].ToString();
                //"Last Name"
                oMarketing.LastName = dt.Rows[i][7].ToString();
                //"Email"
                oMarketing.Email = dt.Rows[i][8].ToString();
                //"Company Name"
                oMarketing.CompanyName = dt.Rows[i][9].ToString();
                //"Total site visits"
                oMarketing.TotalSiteVisits = dt.Rows[i][10].ToString();
                //"Num of visits"
                oMarketing.NumOfVisits = dt.Rows[i][11].ToString();
                //"Three visitor"
                oMarketing.ThreeVisitor = dt.Rows[i][12].ToString();
                //"Seven visitor"
                oMarketing.SevenVisitor = dt.Rows[i][13].ToString();
                //"Job Title"
                oMarketing.JobTitle = dt.Rows[i][14].ToString();
                //"Function"
                oMarketing.Function = dt.Rows[i][15].ToString();
                ///"Seniority"
                oMarketing.Seniority = dt.Rows[i][16].ToString();
                //"Industry"
                oMarketing.Industry = dt.Rows[i][17].ToString();
                //"Add to book list"
                oMarketing.AddToBookList = dt.Rows[i][18].ToString();
                //"City"
                oMarketing.City = dt.Rows[i][19].ToString();
                //"Province"
                oMarketing.Province = dt.Rows[i][20].ToString();
                oMarketing.Country = dt.Rows[i][21].ToString();
                //Street
                oMarketing.Street = dt.Rows[i][22].ToString();
                //Date Added
                oMarketing.DateAdded = dt.Rows[i][23].ToString();
                //"Postal Code"
                oMarketing.PostalCode = dt.Rows[i][24].ToString();
                // Notes
                oMarketing.Notes = dt.Rows[i][25].ToString();
                //Interaction Date
                oMarketing.InteractionDate = dt.Rows[i][26].ToString();
                //"LinkedIn Lead Profile URL"
                oMarketing.LinkedInLeadProfileURL = dt.Rows[i][27].ToString();
                //"Owner"
                oMarketing.Owner = dt.Rows[i][28].ToString();
                //"Created by"
                oMarketing.CreatedBy = dt.Rows[i][29].ToString();
                //"CMS"
                oMarketing.CMS = dt.Rows[i][30].ToString();
                //"Add to LI campaign"
                oMarketing.AddToLICampaign = dt.Rows[i][31].ToString();
                //"Sent Book"
                oMarketing.SentBook = dt.Rows[i][32].ToString();
                //"Book Return Reason"
                oMarketing.BookReturnReason = dt.Rows[i][33].ToString();
                //Book Sent Date
                oMarketing.BookSentDate = dt.Rows[i][34].ToString();
                //"Extra Books Sent"
                oMarketing.ExtraBooksSent = dt.Rows[i][35].ToString();
                //"Website"
                oMarketing.Website = dt.Rows[i][36].ToString();
                //"Sitecore User"
                //  oMarketing.SitecoreUser = dt.Rows[i][37].ToString();

                oListMarketing.Add(oMarketing);
                i++;
            }
            #endregion

            #region Login Control
            // Establish the Login control
            CrmLogin ctrl = new CrmLogin();
            // Wire Event to login response. 
            ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;
            // Show the dialog. 
            ctrl.ShowDialog();

            // Handel return. 
            if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                MessageBox.Show("Good Connect");
            else
                MessageBox.Show("BadConnect");

            #endregion

            #region LeadScore
            // if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
            // {
            //     CrmServiceClient svcClients = ctrl.CrmConnectionMgr.CrmSvc;
            //     if (svcClients.IsReady)
            //     {

            //         //Load Lead
            //         // Get data from CRM . 
            //         string FetchXML =
            //              @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                             <entity name='lead'>
            //                                 <attribute name='firstname' />
            //                                 <attribute name='lastname' />
            //                                 <attribute name='contactid' />
            //                                 <attribute name = 'companyname' />
            //                               </entity>
            //                             </fetch>";

            //         // 
            //         //<attribute name='leadid' />
            //         //  <order attribute='firstname' descending='false' />


            //         var Result = svcClients.GetEntityDataByFetchSearchEC(FetchXML).Entities.ToList();
            //         List<Lead> oLead = new List<Lead>();
            //         foreach (Lead c in Result)
            //         {
            //             oLead.Add(c);
            //         }


            //         foreach (Marketing osMarketing in oListMarketing)
            //         {
            //           //  osMarketing.FirstName = "Ken";
            //           //  osMarketing.LastName = "Hoff";
            //           //  osMarketing.CompanyName = "Langley School Board No 35";

            //              // Load data from the Excelsheet
            //              Lead o = oLead.FirstOrDefault(x => ((x.FirstName != null && x.FirstName.ToUpper().Equals(osMarketing.FirstName.ToUpper())) &&
            //                                                 (x.LastName != null && x.LastName.ToUpper().Equals(osMarketing.LastName.ToUpper())) &&
            //                                                 (x.CompanyName != null && x.CompanyName.ToUpper().Equals(osMarketing.CompanyName.ToUpper()))
            //              ));

            //         if(o!= null)
            //         { 
            //                     try
            //                     {
            //                         Entity LeadENt = new Entity("lead");
            //                         LeadENt.Attributes.Add("leadid", o.LeadId);
            //                         LeadENt.Attributes.Add("new_marketingleadsource", new OptionSetValue(this.SetLeadSource(osMarketing.LeadSource)));
            //                         svcClients.OrganizationServiceProxy.Update(LeadENt);
            //                         Console.Write("Found Lead  - " +osMarketing.FirstName + "," + osMarketing.LastName + "\n");
            //                     }
            //                     catch (Exception e1)
            //                     {
            //                         Console.Write("Not Lead  - " + e1.Message + "\n");
            //                         Console.Write("Not osMarketing- " + osMarketing + "\n");

            //                     }
            //           }else
            //             {
            //                 Console.Write("Not Found Lead  - " + osMarketing.FirstName + "," + osMarketing.LastName + "\n");
            //             }
            //        }
            //    }
            //}
            #endregion

            #region CRMServiceClient
            if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
            {
                CrmServiceClient svcClient = ctrl.CrmConnectionMgr.CrmSvc;
                if (svcClient.IsReady)
                {

                    //Load Lead
                    // Get data from CRM . 
                    string FetchXML =
                         @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='lead'>
                                            <attribute name='firstname' />
                                            <attribute name='lastname' />
                                            <attribute name='contactid' />
                                            <attribute name = 'companyname' />
                                          </entity>
                                        </fetch>";

                    // 
                    //<attribute name='leadid' />
                    //  <order attribute='firstname' descending='false' />


                    var Result = svcClient.GetEntityDataByFetchSearchEC(FetchXML).Entities.ToList();
                    List<Lead> oLead = new List<Lead>();
                    foreach (Lead c in Result)
                    {
                        oLead.Add(c);
                    }

                    int Index = 1;


                    foreach (Marketing osMarketing in oListMarketing)
                    {
                        // Load data from the Excelsheet
                        Lead o = oLead.FirstOrDefault(x => ((x.FirstName != null && x.FirstName.ToUpper().Equals(osMarketing.FirstName.ToUpper())) &&
                                                            (x.LastName != null && x.LastName.ToUpper().Equals(osMarketing.LastName.ToUpper())) &&
                                                            (x.CompanyName != null && x.CompanyName.ToUpper().Equals(osMarketing.CompanyName.ToUpper()))
                         ));

                        //if (Result != null)
                        //{
                        //    MessageBox.Show(string.Format("Found {0} records\nFirst Record name is {1}", Result.Entities.Count, Result.Entities.FirstOrDefault().GetAttributeValue<string>("name")));
                        //}
                        DateTime thisDay = DateTime.Today;

                        if (o == null)
                        {
                            try
                            {
                                CreateRequest req = new CreateRequest();
                                Entity LeadENt = new Entity("lead");
                                LeadENt.Attributes.Add("subject", osMarketing.ProjectName);
                                LeadENt.Attributes.Add("new_marketingleadsource", new OptionSetValue(this.SetLeadSource(osMarketing.LeadSource)));
                                LeadENt.Attributes.Add("nlc_division", new OptionSetValue(857710001));
                                LeadENt.Attributes.Add("firstname", osMarketing.FirstName);
                                LeadENt.Attributes.Add("lastname", osMarketing.LastName);
                                LeadENt.Attributes.Add("companyname", osMarketing.CompanyName);
                                LeadENt.Attributes.Add("nlc_jobfunction", new OptionSetValue(this.SetFunction(osMarketing.Function)));
                                LeadENt.Attributes.Add("nlc_seniority", new OptionSetValue(this.SetSeniority(osMarketing.Seniority)));
                                LeadENt.Attributes.Add("nlc_industry", new OptionSetValue(this.SetIndustry(osMarketing.Industry)));

                                if (osMarketing.JobTitle.Length > 100)
                                    LeadENt.Attributes.Add("jobtitle", osMarketing.JobTitle.Substring(0, 99));
                                else
                                    LeadENt.Attributes.Add("jobtitle", osMarketing.JobTitle);

                                //Leigh-Ann Redmond
                                LeadENt.Attributes.Add("ownerid", new EntityReference("systemuser", new Guid("b441b00b-eb7b-e511-80d9-3863bb35af70")));

                                //Address
                                string address = osMarketing.Street + " " + osMarketing.City + " " + osMarketing.Province + " " + osMarketing.Country;
                                if (!string.IsNullOrEmpty(address))
                                    LeadENt.Attributes.Add("address1_composite", address);
                                if (!string.IsNullOrEmpty(osMarketing.City))
                                    LeadENt.Attributes.Add("address1_city", osMarketing.City);
                                if (!string.IsNullOrEmpty(osMarketing.Country))
                                    LeadENt.Attributes.Add("address1_country", osMarketing.Country);
                                if (!string.IsNullOrEmpty(osMarketing.Street))
                                    LeadENt.Attributes.Add("address1_line1", osMarketing.Street);
                                if (!string.IsNullOrEmpty(osMarketing.Province))
                                    LeadENt.Attributes.Add("address1_stateorprovince", osMarketing.Province);

                                try
                                {

                                    if (!string.IsNullOrEmpty(osMarketing.BookSentDate))
                                        LeadENt.Attributes.Add("new_booksentdate", Convert.ToDateTime(osMarketing.BookSentDate));
                                }
                                catch { }

                                if (!string.IsNullOrEmpty(osMarketing.Email))
                                {
                                    LeadENt.Attributes.Add("emailaddress1", osMarketing.Email);
                                }

                                if (!string.IsNullOrEmpty(osMarketing.AddToLICampaign))
                                {

                                    LeadENt.Attributes.Add("new_dateaddedtolicampaign", thisDay);
                                }
                                if (!string.IsNullOrEmpty(osMarketing.Notes))
                                    LeadENt.Attributes.Add("new_marketingnote", osMarketing.Notes);

                                if (!string.IsNullOrEmpty(osMarketing.AddToBookList))
                                {
                                    if (osMarketing.AddToBookList.Equals("Y"))
                                        LeadENt.Attributes.Add("new_addtobooklist", new OptionSetValue(100000000));
                                    else
                                        LeadENt.Attributes.Add("new_addtobooklist", new OptionSetValue(100000001));
                                }

                                if (!string.IsNullOrEmpty(osMarketing.AddToLICampaign))
                                {
                                    if (osMarketing.AddToLICampaign.Equals("Y"))
                                        LeadENt.Attributes.Add("new_addtolicampaign", new OptionSetValue(100000000));
                                    else
                                        LeadENt.Attributes.Add("new_addtolicampaign", new OptionSetValue(100000001));
                                }

                                //if (!string.IsNullOrEmpty(osMarketing.so))
                                //{
                                //    if (osMarketing.AddToLICampaign.Equals("Y"))
                                //        LeadENt.Attributes.Add("new_addtosourcecompaign", new OptionSetValue(100000001));
                                //    else
                                //        LeadENt.Attributes.Add("new_addtomarketinglist", new OptionSetValue(100000000));
                                //}

                                if (!string.IsNullOrEmpty(osMarketing.SentBook))
                                {
                                    if (osMarketing.SentBook.Equals("Y"))
                                        LeadENt.Attributes.Add("new_sentbook", new OptionSetValue(100000000));
                                    else
                                        LeadENt.Attributes.Add("new_sentbook", new OptionSetValue(100000001));
                                }

                                req.Target = LeadENt;
                                CreateResponse ress = (CreateResponse)svcClient.OrganizationServiceProxy.Execute(req);
                                //  CreateResponse res = (CreateResponse)svcClient.ExecuteCrmOrganizationRequest(req, "MyAccountCreate");

                                Lead loadLead = new Lead();
                                loadLead.CompanyName = osMarketing.CompanyName;
                                loadLead.FirstName = osMarketing.FirstName;
                                loadLead.LastName = osMarketing.LastName;
                                loadLead.Id = ress.id;

                                oLead.Add(loadLead);
                                Console.Write("New Lead - " + ress.id.ToString() + " ------ " + "Index : -" + Index + "\n");

                                Index++;
                                //  MessageBox.Show(ress.id.ToString());
                            }
                            catch (Exception e1)
                            {
                                Console.Write("New Lead  - " + e1.Message + "\n");
                                Console.Write("New Lead  osMarketing- " + osMarketing + "\n");

                            }
                        }
                        else {
                            #region Load lead Interaction                          
                            //string FetchXMLLeadInteraction = "<fetch mapping='logical'>" +
                            //         "<entity name='lead'> " +
                            //            "<attribute name='firstname'/>" +
                            //            "<attribute name='lastname'/>" +
                            //            "<attribute name = 'companyname'/>" +
                            //               "<filter type='and'>" +
                            //                  "<condition attribute='firstname' operator='eq' value='" + o.FirstName + "'/> " +
                            //                  "<condition attribute='lastname' operator='eq' value= '" + o.LastName + "' /> " +
                            //                  "<condition attribute='companyname' operator='eq' value= '" + o.CompanyName + "' /> " +
                            //               "</filter>" +
                            //         "</entity>" +
                            //       "</fetch> ";
                            //var Results = svcClient.GetEntityDataByFetchSearchEC(FetchXMLLeadInteraction).Entities.ToList();                        
                            //foreach (Lead c in Results)
                            //{
                            //    o.Id = c.Id;
                            //    break;
                            //}

                            try
                            {

                                CreateRequest req = new CreateRequest();
                                Entity accENt = new Entity("nlc_leadinteractions");

                                // accENt.Attributes.Add("nlc_lead", new EntityReference("lead", new Guid("89590dc1-a231-e611-80ec-5065f38a6b31")));
                                accENt.Attributes.Add("nlc_lead", new EntityReference("lead", o.Id));

                                //Leigh-Ann Redmond
                                accENt.Attributes.Add("ownerid", new EntityReference("systemuser", new Guid("b441b00b-eb7b-e511-80d9-3863bb35af70")));

                                //  accENt.Attributes.Add("nlc_lead", new EntityReference("lead", o.Id));
                                if (!string.IsNullOrEmpty(osMarketing.InteractionType))
                                    accENt.Attributes.Add("new_marketinginteractiontype", new OptionSetValue(this.SetInteractionType(osMarketing.InteractionType)));

                                try
                                {
                                    if (!string.IsNullOrEmpty(osMarketing.InteractionDate))
                                        accENt.Attributes.Add("nlc_interactiondate", Convert.ToDateTime(osMarketing.InteractionDate));
                                }
                                catch { }

                                if (!string.IsNullOrEmpty(osMarketing.TotalSiteVisits))
                                    accENt.Attributes.Add("nlc_sitevisits", osMarketing.TotalSiteVisits);

                                if (!string.IsNullOrEmpty(osMarketing.NumOfVisits))
                                    accENt.Attributes.Add("nlc_sitevisits", osMarketing.NumOfVisits);

                                if (!string.IsNullOrEmpty(osMarketing.Content))
                                    accENt.Attributes.Add("nlc_content", osMarketing.Content);

                                if (!string.IsNullOrEmpty(osMarketing.CompanyName))
                                    accENt.Attributes.Add("nlc_name", osMarketing.CompanyName);

                                if (!string.IsNullOrEmpty(osMarketing.SocialNetwork))
                                    accENt.Attributes.Add("nlc_socialnetwork", new OptionSetValue(this.SetSocialNetwork(osMarketing.SocialNetwork)));

                                if (!string.IsNullOrEmpty(osMarketing.Email))
                                    accENt.Attributes.Add("new_email", osMarketing.Email);

                                try
                                {

                                    if (!string.IsNullOrEmpty(osMarketing.BookSentDate))
                                        accENt.Attributes.Add("new_booksentdate", Convert.ToDateTime(osMarketing.BookSentDate));
                                }
                                catch { }

                                if (!string.IsNullOrEmpty(osMarketing.BookReturnReason))
                                    accENt.Attributes.Add("new_bookreturnreason", osMarketing.BookReturnReason);

                                if (!string.IsNullOrEmpty(osMarketing.Country))
                                    accENt.Attributes.Add("new_country", osMarketing.Country);

                                if (!string.IsNullOrEmpty(osMarketing.DateAdded))
                                    accENt.Attributes.Add("new_dateadded", osMarketing.DateAdded);

                                if (!string.IsNullOrEmpty(osMarketing.Notes))
                                    accENt.Attributes.Add("new_notes", osMarketing.Notes);

                                if (!string.IsNullOrEmpty(osMarketing.Notes))
                                    accENt.Attributes.Add("new_description", osMarketing.Notes);

                                if (!string.IsNullOrEmpty(osMarketing.AddToBookList))
                                {
                                    if (osMarketing.AddToBookList.Equals("Y"))
                                        accENt.Attributes.Add("new_addtobooklist", new OptionSetValue(100000000));
                                    else
                                        accENt.Attributes.Add("new_addtobooklist", new OptionSetValue(100000001));
                                }

                                if (!string.IsNullOrEmpty(osMarketing.ExtraBooksSent))
                                {
                                    if (osMarketing.ExtraBooksSent.Equals("Y"))
                                        accENt.Attributes.Add("new_extrabookssent", new OptionSetValue(100000000));
                                    else
                                        accENt.Attributes.Add("new_extrabookssent", new OptionSetValue(100000001));
                                }


                                if (!string.IsNullOrEmpty(osMarketing.ThreeVisitor))
                                {
                                    if (osMarketing.ExtraBooksSent.Equals("Y"))
                                        accENt.Attributes.Add("new_3visitor", new OptionSetValue(100000000));
                                    else
                                        accENt.Attributes.Add("new_3visitor", new OptionSetValue(100000001));
                                }

                                if (!string.IsNullOrEmpty(osMarketing.SevenVisitor))
                                {
                                    if (osMarketing.ExtraBooksSent.Equals("Y"))
                                        accENt.Attributes.Add("new_7visitor", new OptionSetValue(100000000));
                                    else
                                        accENt.Attributes.Add("new_7visitor", new OptionSetValue(100000001));
                                }

                                if (!string.IsNullOrEmpty(osMarketing.SentBook))
                                {
                                    if (osMarketing.ExtraBooksSent.Equals("Y"))
                                        accENt.Attributes.Add("new_sentbook", new OptionSetValue(100000000));
                                    else
                                        accENt.Attributes.Add("new_sentbook", new OptionSetValue(100000001));
                                }

                                if (!string.IsNullOrEmpty(osMarketing.AddToLICampaign))
                                {
                                    if (osMarketing.AddToLICampaign.Equals("Y"))
                                        accENt.Attributes.Add("new_addtolicampaign", new OptionSetValue(100000000));
                                    else
                                        accENt.Attributes.Add("new_addtolicampaign", new OptionSetValue(100000001));
                                }


                                req.Target = accENt;
                                CreateResponse res = (CreateResponse)svcClient.OrganizationServiceProxy.Execute(req);
                                //CreateResponse res = (CreateResponse)svcClient.ExecuteCrmOrganizationRequest(req, "MyAccountCreate");
                                //  MessageBox.Show(res.id.ToString());
                                Console.Write("New Lead Interaction - " + res.id.ToString() + "\n");
                            }
                            catch (Exception e2)
                            {
                                Console.Write("New Lead Interaction - " + e2.Message + "\n");
                                Console.Write("New Lead Interaction osMarketing- " + osMarketing + "\n");

                            }
                            #endregion
                        }

                        //                    // Core API using SDK OOTB 

                        //                     CreateRequest req = new CreateRequest();
                        //                    Entity accENt = new Entity("account");
                        //                    accENt.Attributes.Add("name", "TESTFOO");
                        //                    req.Target = accENt;
                        //                    CreateResponse res = (CreateResponse)svcClient.OrganizationServiceProxy.Execute(req);
                        //                    //CreateResponse res = (CreateResponse)svcClient.ExecuteCrmOrganizationRequest(req, "MyAccountCreate");
                        //                    MessageBox.Show(res.id.ToString());



                        //                    // Using Xrm.Tooling helpers. 
                        //                    Dictionary<string, CrmDataTypeWrapper> newFields = new Dictionary<string, CrmDataTypeWrapper>();
                        //                    // Create a new Record. - Account 
                        //                    newFields.Add("name", new CrmDataTypeWrapper("CrudTestAccount", CrmFieldType.String));
                        //                    Guid guAcctId = svcClient.CreateNewRecord("account", newFields);

                        //                    MessageBox.Show(string.Format("New Record Created {0}", guAcctId));

                    }
                }
            }
            #endregion


        }

        private int SetCountry(string leadSource)
        {
            int ret = 0;

            if (leadSource.ToUpper().Equals("Algeria".ToUpper()))
                ret = 100000000;
            else if (leadSource.ToUpper().Equals("Australia".ToUpper()))
                ret = 100000003;
            else if (leadSource.ToUpper().Equals("Bangladesh".ToUpper()))
                ret = 100000004;
            else if (leadSource.ToUpper().Equals("Brazil".ToUpper()))
                ret = 100000005;
            else if (leadSource.ToUpper().Equals("Canada".ToUpper()))
                ret = 100000006;
            else if (leadSource.ToUpper().Equals("China".ToUpper()))
                ret = 100000007;
            else if (leadSource.ToUpper().Equals("Costa Rica".ToUpper()))
                ret = 100000008;
            else if (leadSource.ToUpper().Equals("Egypt".ToUpper()))
                ret = 100000009;
            else if (leadSource.ToUpper().Equals("France".ToUpper()))
                ret = 100000010;
            else if (leadSource.ToUpper().Equals("Germany".ToUpper()))
                ret = 100000011;
            else if (leadSource.ToUpper().Equals("Indonesia".ToUpper()))
                ret = 100000012;
            else if (leadSource.ToUpper().Equals("Italy".ToUpper()))
                ret = 100000013;
            else if (leadSource.ToUpper().Equals("Ireland".ToUpper()))
                ret = 100000014;
            else if (leadSource.ToUpper().Equals("Kuwait".ToUpper()))
                ret = 100000015;
            else if (leadSource.ToUpper().Equals("Morocco".ToUpper()))
                ret = 100000016;
            else if (leadSource.ToUpper().Equals("Nigeria".ToUpper()))
                ret = 100000017;
            else if (leadSource.ToUpper().Equals("Pakistan".ToUpper()))
                ret = 100000018;
            else if (leadSource.ToUpper().Equals("Philippines".ToUpper()))
                ret = 100000019;
            else if (leadSource.ToUpper().Equals("Portugal".ToUpper()))
                ret = 100000020;
            else if (leadSource.ToUpper().Equals("South Africa".ToUpper()))
                ret = 100000021;
            else if (leadSource.ToUpper().Equals("Spain".ToUpper()))
                ret = 100000022;
            else if (leadSource.ToUpper().Equals("Switzerland".ToUpper()))
                ret = 100000023;
            else if (leadSource.ToUpper().Equals("United States".ToUpper()))
                ret = 100000002;
            else if (leadSource.ToUpper().Equals("United Kingdom".ToUpper()))
                ret = 100000001;
            else if (leadSource.ToUpper().Equals("GB".ToUpper()))
                ret = 100000001;
            else if (leadSource.ToUpper().Equals("BR".ToUpper()))
                ret = 100000005;
            else if (leadSource.ToUpper().Equals("CA".ToUpper()))
                ret = 100000006;
            else if (leadSource.ToUpper().Equals("US".ToUpper()))
                ret = 100000002;
            else if (leadSource.ToUpper().Equals("ES".ToUpper()))
                ret = 100000022;
            else if (leadSource.ToUpper().Equals("BE".ToUpper()))
                ret = 857710024;
            else
                ret = 857710024;

            return ret;
        }

        private int SetSocialNetwork(string leadSource)
        {
            int ret = 0;
            leadSource = leadSource.Trim();
            if (leadSource.ToUpper().Equals("Website".ToUpper()))
                ret = 857710003;
            else if (leadSource.ToUpper().Equals("LinkedIn".ToUpper()))
                ret = 857710001;
            else if (leadSource.ToUpper().Equals("Twitter".ToUpper()))
                ret = 857710002;
            else if (leadSource.ToUpper().Equals("Facebook".ToUpper()))
                ret = 857710000;
            else
                ret = 100000000;

            return ret;

        }

        private int SetIndustry(string leadSource)
        {
            int ret = 0;
            leadSource = leadSource.Trim();

            if (leadSource.Equals("Construction/Engineering".ToUpper()))
                ret = 857710045;
            else if (leadSource.ToUpper().Equals("Health Care".ToUpper()))
                ret = 857710041;
            else if (leadSource.ToUpper().Equals("Legal Services".ToUpper()))
                ret = 857710020;
            else if (leadSource.ToUpper().Equals("Accounting".ToUpper()))
                ret = 857710000;
            else if (leadSource.ToUpper().Equals("Agriculture and Non-petrol Natural Resource Extraction".ToUpper()))
                ret = 857710001;
            else if (leadSource.ToUpper().Equals("Broadcasting, Priniting and Publishing".ToUpper()))
                ret = 857710002;
            else if (leadSource.ToUpper().Equals("Brokers".ToUpper()))
                ret = 857710003;
            else if (leadSource.ToUpper().Equals("Building Supply Retail".ToUpper()))
                ret = 857710004;
            else if (leadSource.ToUpper().Equals("Business Services".ToUpper()))
                ret = 857710005;
            else if (leadSource.ToUpper().Equals("Consulting".ToUpper()))
                ret = 857710006;
            else if (leadSource.ToUpper().Equals("Consumer Services".ToUpper()))
                ret = 857710007;
            else if (leadSource.ToUpper().Equals("Design, Direction and Creative Management".ToUpper()))
                ret = 857710008;
            else if (leadSource.ToUpper().Equals("Distributors, Dispatchers and Processors".ToUpper()))
                ret = 857710009;
            else if (leadSource.ToUpper().Equals("Doctor's Offices and Clinics".ToUpper()))
                ret = 857710010;
            else if (leadSource.ToUpper().Equals("Durable Manufacturing".ToUpper()))
                ret = 857710011;
            else if (leadSource.ToUpper().Equals("Eating and Drinking Places".ToUpper()))
                ret = 857710012;
            else if (leadSource.ToUpper().Equals("Entertainment Retail".ToUpper()))
                ret = 857710013;
            else if (leadSource.ToUpper().Equals("Equipment Rental and Leasing".ToUpper()))
                ret = 857710014;
            else if (leadSource.ToUpper().Equals("Financial Services".ToUpper()))
                ret = 857710015;
            else if (leadSource.ToUpper().Equals("Food and Tobacco Processing".ToUpper()))
                ret = 857710016;
            else if (leadSource.ToUpper().Equals("Goverment - State / Provincial".ToUpper()))
                ret = 857710033;
            else if (leadSource.ToUpper().Equals("Government - Federal".ToUpper()))
                ret = 857710034;
            else if (leadSource.ToUpper().Equals("Government - Municipal".ToUpper()))
                ret = 857710035;
            else if (leadSource.ToUpper().Equals("High Tech".ToUpper()))
                ret = 857710036;
            else if (leadSource.ToUpper().Equals("Higher Education".ToUpper()))
                ret = 857710036;
            else if (leadSource.ToUpper().Equals("Inbound Capital Intensive Processing".ToUpper()))
                ret = 857710017;
            else if (leadSource.ToUpper().Equals("Inbound Repair and Services".ToUpper()))
                ret = 857710018;
            else if (leadSource.ToUpper().Equals("Insurance".ToUpper()))
                ret = 857710019;
            else if (leadSource.ToUpper().Equals("Manufacturing".ToUpper()))
                ret = 857710038;
            else if (leadSource.ToUpper().Equals("Mining".ToUpper()))
                ret = 857710039;
            else if (leadSource.ToUpper().Equals("Non-Durable Merchandise Retail".ToUpper()))
                ret = 857710021;
            else if (leadSource.ToUpper().Equals("Oil and Gas".ToUpper()))
                ret = 857710040;
            else if (leadSource.ToUpper().Equals("Outbound Consumer Service".ToUpper()))
                ret = 857710022;
            else if (leadSource.ToUpper().Equals("Petrochemical Extraction and Distribution".ToUpper()))
                ret = 857710023;
            else if (leadSource.ToUpper().Equals("Real Estate / Home Building".ToUpper()))
                ret = 857710042;
            else if (leadSource.ToUpper().Equals("Retail".ToUpper()))
                ret = 857710043;
            else if (leadSource.ToUpper().Equals("Service Retail".ToUpper()))
                ret = 857710024;
            else if (leadSource.ToUpper().Equals("SIG Affiliations".ToUpper()))
                ret = 857710025;
            else if (leadSource.ToUpper().Equals("Social Services".ToUpper()))
                ret = 857710026;
            else if (leadSource.ToUpper().Equals("Special Outbound Trade Contractors".ToUpper()))
                ret = 857710027;
            else if (leadSource.ToUpper().Equals("Specialty Realty".ToUpper()))
                ret = 857710028;
            else if (leadSource.ToUpper().Equals("Telecommunications".ToUpper()))
                ret = 857710044;
            else if (leadSource.ToUpper().Equals("Transportation".ToUpper()))
                ret = 857710029;
            else if (leadSource.ToUpper().Equals("Utility Creation and Distribution".ToUpper()))
                ret = 857710030;
            else if (leadSource.ToUpper().Equals("Vehicle Retail".ToUpper()))
                ret = 857710031;
            else if (leadSource.ToUpper().Equals("Wholesale".ToUpper()))
                ret = 857710032;
            else if (leadSource.ToUpper().Equals("Other".ToUpper()))
                ret = 100000000;
            else
                ret = 100000000;
            return ret;
        }

        private int SetSeniority(string leadSource)
        {
            int ret = 0;
            leadSource = leadSource.Trim();
            if (leadSource.ToUpper().Equals("C-level".ToUpper()))
                ret = 857710000;
            else if (leadSource.ToUpper().Equals("VP-level".ToUpper()))
                ret = 857710001;
            else if (leadSource.ToUpper().Equals("Director-level".ToUpper()))
                ret = 857710002;
            else if (leadSource.ToUpper().Equals("Manager".ToUpper()))
                ret = 857710003;
            else if (leadSource.ToUpper().Equals("Specialist/Officer".ToUpper()))
                ret = 857710004;
            else if (leadSource.ToUpper().Equals("Other Unknown".ToUpper()))
                ret = 857710005;
            else
                ret = 857710005;

            return ret;

        }

        //new_marketinginteractiontype
        private int SetFunction(string leadSource)
        {
            int ret = 0;
            leadSource = leadSource.Trim();
            if (leadSource.ToUpper().Equals("Communications(Internal even better)".ToUpper()))
                ret = 857710003;
            else if (leadSource.ToUpper().Equals("Marketing".ToUpper()))
                ret = 857710000;
            else if (leadSource.ToUpper().Equals("IT".ToUpper()))
                ret = 857710001;
            else if (leadSource.ToUpper().Equals("HR".ToUpper()))
                ret = 857710002;
            else if (leadSource.ToUpper().Equals("Other or Unknown".ToUpper()))
                ret = 857710004;
            else
                ret = 857710004;

            return ret;
        }

        //new_marketinginteractiontype
        private int SetInteractionType(string leadSource)
        {
            int ret = 0;
            leadSource = leadSource.Trim();
            if (leadSource.ToUpper().Equals("Followed us on LI".ToUpper()))
                ret = 100000000;
            else if (leadSource.ToUpper().Equals("Followed us on Twitter".ToUpper()))
                ret = 100000002;
            else if (leadSource.ToUpper().Equals("Liked LI content".ToUpper()))
                ret = 100000003;
            else if (leadSource.ToUpper().Equals("Liked Twitter content".ToUpper()))
                ret = 100000004;
            else if (leadSource.ToUpper().Equals("Followed us on LI".ToUpper()))
                ret = 100000005;
            else if (leadSource.ToUpper().Equals("Followed us on Twitter".ToUpper()))
                ret = 100000001;
            else if (leadSource.ToUpper().Equals("Liked LI content".ToUpper()))
                ret = 100000006;
            else if (leadSource.ToUpper().Equals("Liked Twitter content".ToUpper()))
                ret = 100000007;
            else if (leadSource.ToUpper().Equals("Shared LI content".ToUpper()))
                ret = 100000008;
            else if (leadSource.ToUpper().Equals("Shared Twitter content".ToUpper()))
                ret = 100000009;
            else if (leadSource.ToUpper().Equals("Digital Transformation Guide DL".ToUpper()))
                ret = 100000010;
            else if (leadSource.ToUpper().Equals("Intranet Guide Download".ToUpper()))
                ret = 100000011;
            else if (leadSource.ToUpper().Equals("Whitepaper Download".ToUpper()))
                ret = 100000012;
            else if (leadSource.ToUpper().Equals("Conference attendance".ToUpper()))
                ret = 100000013;
            else if (leadSource.ToUpper().Equals("Email marketing reply".ToUpper()))
                ret = 100000014;
            else if (leadSource.ToUpper().Equals("Lead subscribed to blog".ToUpper()))
                ret = 100000015;
            else if (leadSource.ToUpper().Equals("Recent Sitecore Rep Connection".ToUpper()))
                ret = 100000016;
            else if (leadSource.ToUpper().Equals("Webinar Sign up".ToUpper()))
                ret = 100000017;
            else if (leadSource.ToUpper().Equals("Attended full webinar".ToUpper()))
                ret = 100000018;
            else if (leadSource.ToUpper().Equals("Opened an email".ToUpper()))
                ret = 100000019;
            else if (leadSource.ToUpper().Equals("Clicked a link in an email".ToUpper()))
                ret = 100000020;
            else if (leadSource.ToUpper().Equals("Company 7+ web visits".ToUpper()))
                ret = 100000021;
            else if (leadSource.ToUpper().Equals("Company 3+ web visits".ToUpper()))
                ret = 100000022;
            else if (leadSource.ToUpper().Equals("Unsubscribe from email".ToUpper()))
                ret = 100000023;
            else if (leadSource.ToUpper().Equals("Downloaded Intranet Guide".ToUpper()))
                ret = 100000011;
            else
                ret = 100000024;

            return ret;
        }

        private int SetLeadSource(string leadSource)
        {
            int ret = 0;
            leadSource = leadSource.Trim();
            if (leadSource.ToUpper().Equals("Digital Transformation Guide DL".ToUpper()))
                ret = 100000000;
            else if (leadSource.ToUpper().Equals("Intranet Guide Download".ToUpper()))
                ret = 100000001;
            else if (leadSource.ToUpper().Equals("Whitepaper Download".ToUpper()))
                ret = 100000002;
            else if (leadSource.ToUpper().Equals("LinkedIn Liker".ToUpper()))
                ret = 100000003;
            else if (leadSource.ToUpper().Equals("Twitter Favourite".ToUpper()))
                ret = 100000004;
            else if (leadSource.ToUpper().Equals("Engagement Plan".ToUpper()))
                ret = 100000005;
            else if (leadSource.ToUpper().Equals("ConferenceAttendee".ToUpper()))
                ret = 100000006;
            else if (leadSource.ToUpper().Equals("LinkedIn Follower".ToUpper()))
                ret = 100000007;
            else if (leadSource.ToUpper().Equals("Twitter Follower".ToUpper()))
                ret = 100000008;
            else if (leadSource.ToUpper().Equals("Frequent visitor".ToUpper()))
                ret = 100000009;
            else if (leadSource.ToUpper().Equals("Company 3 + web visits".ToUpper()))
                ret = 100000010;
            else if (leadSource.ToUpper().Equals("LinkedIn Shared".ToUpper()))
                ret = 100000011;
            else if (leadSource.ToUpper().Equals("Twitter Retweet".ToUpper()))
                ret = 100000012;
            else if (leadSource.ToUpper().Equals("Inbound request by email or phone".ToUpper()))
                ret = 100000013;
            else if (leadSource.ToUpper().Equals("Blog Subscriber".ToUpper()))
                ret = 100000014;
            else if (leadSource.ToUpper().Equals("Analyst Referral".ToUpper()))
                ret = 100000015;
            else if (leadSource.ToUpper().Equals("Client Referral".ToUpper()))
                ret = 100000016;
            else if (leadSource.ToUpper().Equals("Sitecore Partner Referral".ToUpper()))
                ret = 100000017;
            else if (leadSource.ToUpper().Equals("Sitecore Partner".ToUpper()))
                ret = 100000018;
            else if (leadSource.ToUpper().Equals("Sitecore Referral".ToUpper()))
                ret = 100000019;
            else if (leadSource.ToUpper().Equals("Recent Sitecore Rep Connection".ToUpper()))
                ret = 100000020;
            else if (leadSource.ToUpper().Equals("Webinar Sign up".ToUpper()))
                ret = 100000021;
            else if (leadSource.ToUpper().Equals("Employee Referral".ToUpper()))
                ret = 100000022;
            else if (leadSource.ToUpper().Equals("Randy's NA Sitecore Users List".ToUpper()))
                ret = 100000023;
            else
            {
                ret = 1000;
            }

            return ret;
        }

        /// <summary>
        /// Raised when the login form process is completed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ctrl_ConnectionToCrmCompleted(object sender, EventArgs e)
        {
            if (sender is CrmLogin)
            {
                this.Dispatcher.Invoke(() =>
                {
                    ((CrmLogin)sender).Close();
                });
            }
        }

    }
}
