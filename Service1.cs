using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Protocols;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;


using System.Text.RegularExpressions;

using Azure.Core;

using System.Net;
using Newtonsoft.Json.Linq;

namespace Simplisys_OAuth2_Service
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer(); // name space(using System.Timers;)  


 


        public Service1()
        {
            InitializeComponent();
        }

        [Obsolete]
        protected override void OnStart(string[] args)
        {


            //WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval =30000; //Convert.ToDouble ( ConfigurationManager.AppSettings["EmailReadingIntervals"])  ;  // 300000; //number in milisecinds  //EmailReadingIntervals
            timer.Enabled = true;
        }

        [Obsolete]
        public void onDebug()
        {
            _ = processEmailAsync();

            //OnStart(null);
        }

        [Obsolete]
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {


            _ = processEmailAsync();

          
        }

        protected override void OnStop()
        {
            //WriteToFile("Service is stopped at " + DateTime.Now);

        }

        [Obsolete]
        public async Task processEmailAsync()
        {

            string strSubjectFirstPart, strSubjectSecondPart;

            string strFrom, strCC, strSubject, strBody, strHTMLBody, strAttachments, strMessageID, strSentDate, strReceivedDate, strDisplayName, strError;

            string strClientID, strClientDatabase, strClientConnectionString;

            string connectionString = ConfigurationManager.AppSettings["ConnectionString"];
            string strAttchmentPath = ConfigurationManager.AppSettings["IncomingAttachmentPath"];

            string strQuery = "select  acctmID , acctmAccountName , acctmConnectionString  from tblmAccount  where acctmAccountName ='LiveJepson' ";


            string strResult;





            string clientID, clientPassword, clientTenant, mailBoxID, strEmail = "", strEmailPassword = "";
            SqlConnection sqlConnection = new SqlConnection(connectionString);
            SqlCommand sqlCommand = new SqlCommand();


            strSubjectFirstPart = ConfigurationManager.AppSettings["SubjectFirstPart"];

            strSubjectSecondPart = ConfigurationManager.AppSettings["SubjectSecondPart"];

            //Read the databases one by one
            //BELOW PART OF THE CODE IS USED TO READ EMAILS 


            try
            {




                sqlConnection.Open();


                sqlCommand = new SqlCommand(strQuery, sqlConnection);

                SqlDataReader reader = sqlCommand.ExecuteReader();



                while (reader.Read())
                {

                    strClientID = reader["acctmID"].ToString();
                    strClientDatabase = reader["acctmAccountName"].ToString();
                    strClientConnectionString = reader["acctmConnectionString"].ToString();
                    SqlCommand sqlClientDBCommand = new SqlCommand();


                    if (strClientID.Length > 0)
                    {


                        if (strClientDatabase.Length > 0)
                        {
                            if (strClientConnectionString.Length > 0)
                            {
                                try
                                {

                                    SqlConnection sqlClientDBConnection = new SqlConnection(strClientConnectionString);

                                    sqlClientDBConnection.Open();
                                    strQuery = "SELECT  [mbOAuth2MboxID]";
                                    strQuery = strQuery + ",[mbOAuth2Name]";
                                    strQuery = strQuery + ",[mbOAuth2EmailAddress]";
                                    strQuery = strQuery + ",[mbOAuth2ClientID]";
                                    strQuery = strQuery + ",[mbOAuth2TenantID]";
                                    strQuery = strQuery + ",[mbOAuth2Password]";
                                    strQuery = strQuery + ",[mbOAuth2isDelete]";
                                    strQuery = strQuery + ",[mbOAuth2IncomingEnabled] ";
                                    strQuery = strQuery + ",[mbOAuth2OutgoingEnabled] ";
                                    strQuery = strQuery + ",[mbOAuth2AutoReplyEnabldEnabled] ";
                                    strQuery = strQuery + ",[mbOAuth2ErrorID] ";
                                    strQuery = strQuery + "  ,(select  mbIncomingPassword  from tblMailbox where mbid=[mbOAuth2MboxID])  as mbPassword";
                                    strQuery = strQuery + " FROM [tblMailboxMSoftOAuthDetails] where  mbOAuth2isDelete=0 and mbOAuth2IncomingEnabled = 1";
                                    sqlClientDBCommand.CommandText = strQuery;


                                    sqlClientDBCommand.Connection = sqlClientDBConnection;
                                    SqlDataReader clientDBreader = sqlClientDBCommand.ExecuteReader();


                                    while (clientDBreader.Read())
                                    {

                                        try
                                        {
                                            clientID = clientDBreader["mbOAuth2ClientID"].ToString();
                                            clientPassword = clientDBreader["mbOAuth2Password"].ToString();
                                            clientTenant = clientDBreader["mbOAuth2TenantID"].ToString();
                                            mailBoxID = clientDBreader["mbOAuth2MboxID"].ToString();
                                            strEmail = clientDBreader["mbOAuth2EmailAddress"].ToString();
                                            strEmailPassword = clientDBreader["mbPassword"].ToString();
                                            //Test -  https://graph.microsoft.com/v1.0/users/research@simplisysltd.onmicrosoft.com/mailFolders/Inbox/messages

                                            GraphServiceClient client = GetAuthenticatedClient(clientID, clientPassword, clientTenant);



                                            //var users = client.Users.Request().GetAsync().Result;


                                            //int userCount = users.Count;





                                            //try
                                            //{


                                            SecureString emailPassword = ConvertToSecureString(strEmailPassword);

                                            MGraphServices graphServices = new MGraphServices();

                                            GraphServiceClient graphServiceClient = await graphServices.connectToMailBox(clientID,
                                                clientTenant,
                                                strEmail,
                                                strEmailPassword);

                                            string testString="";

                                            //var _MailList = await graphServiceClient.Me.MailFolders.Inbox.Messages
                                            //     .Request().Expand("attachments").GetAsync();

                                            //var _MailList2 = await graphServiceClient.Me.MailFolders[strEmail].Messages.Request().GetAsync();


                                            //var _MailList = await graphServiceClient.Users[strEmail].Messages.Request().Expand("attachments").GetAsync();


                                            //var messages = await graphServiceClient.Me.MailFolders.Inbox.Messages
                                            //     .Request().Expand("attachments").GetAsync();

                                            var messages = await graphServiceClient.Me.MailFolders.Inbox.Messages
                                         .Request().Filter("isRead eq false").Expand("attachments").GetAsync();




                                            //var authResult = await pca.AcquireTokenByUsernamePassword(new string[] { "https://graph.microsoft.com/.default" }, strEmail, emailPassword).ExecuteAsync();


                                            //}
                                            //catch (Exception ex)
                                            //{
                                            //    strError = ex.Message.ToString();
                                            //}


                                            //Below line commented on 13 th Sep 2022

                                            //var messages1 = await client.Users[strEmail].Messages
                                            //                 .Request().Expand("attachments")
                                            //                 .GetAsync();


                                            //Above line commented on 13 th Sep 2022












                                            string strattachmentName, strsubject = "";

                                            Guid strIncID;

                                            Guid strIncidentEmailID;

                                            Guid inEmailID;
                                            string strServerAddress, strIncomingUsername, strIncomingEmailPassword, strMailBoxID;
                                            string strDateString, strDate, strYear, strMonth, strDay, strAttachmentName = "";


                                            SqlConnection sqlMessageIDDBConnection = new SqlConnection(strClientConnectionString);
                                            SqlCommand sqlMessageIDDBCommand = new SqlCommand();


                                            for (int i = 0; i < messages.CurrentPage.Count; i++)
                                            {



                                                string strCurrentPage = messages.CurrentPage.ToString();
                                                //READ IF THE MESSAGEID IS IN TBLINCOMINH EMAIL TABLE
                                                int countMessageID = 0;
                                                sqlMessageIDDBConnection.Open();

                                                try
                                                {



                                                    strQuery = "select count(inemlSubject ) from tblIncomingEmail where inemlMessageID ='" + messages.CurrentPage[i].InternetMessageId.ToString() + "' and inemlDateAdded >getdate()-1";
                                                    sqlMessageIDDBCommand.CommandText = strQuery;


                                                    sqlClientDBCommand.Connection = sqlMessageIDDBConnection;
                                                    SqlDataReader messageCountreader = sqlClientDBCommand.ExecuteReader();


                                                    while (messageCountreader.Read())
                                                    {
                                                        countMessageID = Convert.ToInt32( messageCountreader[0].ToString());
                                                    }
                                                    messageCountreader.Close();
                                                    
                                                }
                                                catch (Exception exc)
                                                {
                                                    strError = exc.Message.ToString();
                                                }
                                                finally
                                                {
                                                    sqlMessageIDDBConnection.Close();
                                                }



                                                if (strEmail == messages.CurrentPage[i].From.EmailAddress.Address.ToString())
                                                {
                                                    await client.Users[strEmail].Messages[messages.CurrentPage[i].Id.ToString()].Request().DeleteAsync();

                                                }

                                                //IF THE EMAIL IS ALREADY READ AND INSERTED IN THE TBLINCOMING EMAIL TABLE
                                                
                                                else if(countMessageID>0)
                                                {
                                                    //already read

                                                }


                                                else //(strEmail != messages.CurrentPage[i].From.EmailAddress.Address.ToString())
                                                {

                                                    if (messages[i].IsRead == false)
                                                    {
                                                        if (messages[i].IsRead is false)
                                                        {

                                                            //BELOW  CODE ADDED ON 31 MARCH 2023
                                                            try
                                                            {
                                                                await client.Users[strEmail].Messages[messages.CurrentPage[i].Id.ToString()].Request().Select("IsRead").UpdateAsync(messages[i]);

                                                            }
                                                            catch (Exception exc)
                                                            {
                                                             
                                                            }

                                                            try
                                                            {
                                                                messages.CurrentPage[i].IsRead = true;

                                                            }
                                                            catch (Exception exc)
                                                            {
                                                              
                                                            }


                                                            //ABOVE CODE ADDED ON 31 MARCH 2023

                                                            try
                                                            {
                                                                // await client.Users[strEmail].Messages[messages.CurrentPage[i].Id.ToString()].Request().DeleteAsync();

                                                                //var msg = await client.Users[strEmail].Messages[messages.CurrentPage[i].Id.ToString()].Request().GetAsync();

                                                                //BELOW LINE ADDED ON 31 MARCH 2023
                                                               
                                                                //ABOVE LINE ADDED ON 31 MARCH 2023
                                                                 

                                                                strFrom = messages.CurrentPage[i].From.EmailAddress.Address.ToString();
                                                                strCC = messages.CurrentPage[i].CcRecipients.ToString();



                                                                strSubject = messages.CurrentPage[i].Subject.ToString();
                                                                strHTMLBody = messages.CurrentPage[i].Body.Content.ToString();
                                                                strBody = messages.CurrentPage[i].BodyPreview.ToString();
                                                                strMessageID = messages.CurrentPage[i].InternetMessageId.ToString();
                                                                strSentDate = messages.CurrentPage[i].SentDateTime.ToString();
                                                                strReceivedDate = messages.CurrentPage[i].ReceivedDateTime.ToString();
                                                                strDisplayName = messages.CurrentPage[i].From.EmailAddress.Name.ToString();


                                                                 

                                                               //if (messages.CurrentPage[i].HasAttachments == true)
                                                               //{



                                                               strDate = messages.CurrentPage[i].SentDateTime.ToString();
                                                               // strDate = messages.CurrentPage[i].CreatedDateTime.ToString();


                                                                //try
                                                                //{
                                                                //    DateTime DT = new DateTime();
                                                                //    DT = Convert.ToDateTime(strDate);

                                                                //    DT = Convert.ToDateTime(DT.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'"));

                                                                //    strDate = DT.ToString();
                                                                //}
                                                                //catch (Exception ex)
                                                                //{
                                                                //    strError = ex.Message.ToString();
                                                                //}




                                                                //strDate = email.Date.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'");

                                                                strIncID = Guid.NewGuid();
                                                                strIncidentEmailID = Guid.NewGuid();
                                                                strMessageID = strMessageID.Replace('<', ' ');
                                                                strMessageID = strMessageID.Replace('>', ' ');

                                                                inEmailID = new Guid(mailBoxID);
                                                                string strFilePath = strAttchmentPath + "\\" + strClientDatabase + "\\Simplisys Incident Support\\" + strMessageID.Trim(); // Your code goes here

                                                                strFilePath = strFilePath.Trim();

                                                                try
                                                                {
                                                                    System.IO.Directory.CreateDirectory(strFilePath);
                                                                }
                                                                catch (Exception ex)
                                                                {
                                                                    strError = ex.Message.ToString();
                                                                }





                                                                //strResult = mynumber;



                                                                ////Prefix  set in tblsystemsettings
                                                                ///
                                                                try
                                                                {

                                                                    SqlConnection sqlConnectionChkInc = new SqlConnection(strClientConnectionString);
                                                                    SqlCommand sqlCommandChkInc = new SqlCommand();
                                                                    sqlConnectionChkInc.Open();

                                                                    string strQueryChkIncident = "select top 1  ssValue  from tblSystemSettings where ssKey='IncUpdateFrmEmailPrefix'  and ssIsDeleted=0";

                                                                    sqlCommandChkInc = new SqlCommand(strQueryChkIncident, sqlConnectionChkInc);

                                                                    SqlDataReader readerChkInc = sqlCommandChkInc.ExecuteReader();

                                                                    while (readerChkInc.Read())
                                                                    {
                                                                        strSubjectFirstPart = readerChkInc["ssValue"].ToString();

                                                                    }
                                                                    readerChkInc.Close();
                                                                    sqlConnectionChkInc.Close();

                                                                }
                                                                catch (Exception exc)
                                                                {
                                                                    string err = exc.Message.ToString();
                                                                }


                                                                /// Suffux set in tblsystemsettings
                                                                ///
                                                                try
                                                                {

                                                                    SqlConnection sqlConnectionChkInc = new SqlConnection(strClientConnectionString);
                                                                    SqlCommand sqlCommandChkInc = new SqlCommand();
                                                                    sqlConnectionChkInc.Open();

                                                                    string strQueryChkIncident = "select top 1  ssValue  from tblSystemSettings where ssKey='IncUpdateFrmEmailSuffix'  and ssIsDeleted=0";

                                                                    sqlCommandChkInc = new SqlCommand(strQueryChkIncident, sqlConnectionChkInc);

                                                                    SqlDataReader readerChkInc = sqlCommandChkInc.ExecuteReader();

                                                                    while (readerChkInc.Read())
                                                                    {
                                                                        strSubjectSecondPart = readerChkInc["ssValue"].ToString();

                                                                    }
                                                                    readerChkInc.Close();
                                                                    sqlConnectionChkInc.Close();
                                                                }
                                                                catch (Exception exc)
                                                                {
                                                                    string err = exc.Message.ToString();
                                                                }



                                                                strResult = Between(strSubject, strSubjectFirstPart, strSubjectSecondPart);


                                                                string var = strResult;
                                                                string mystr = Regex.Replace(var, @"\d", "");
                                                                string mynumber = Regex.Replace(var, @"\D", "");

                                                                // This is to remove I from the Incident number


                                                                string strIncidentNumber = "INCNUMBER";


                                                                if (strResult != "NoMatchFound")
                                                                {
                                                                    try
                                                                    {
                                                                        strResult = strResult.Trim();
                                                                        strResult = strResult.Remove(0, 1);

                                                                        SqlConnection sqlConnectionChkInc = new SqlConnection(strClientConnectionString);
                                                                        SqlCommand sqlCommandChkInc = new SqlCommand();
                                                                        sqlConnectionChkInc.Open();

                                                                        string strQueryChkIncident = "select incID, incNumber from tblincident where incNumber='" + strResult + "'";

                                                                        sqlCommandChkInc = new SqlCommand(strQueryChkIncident, sqlConnectionChkInc);

                                                                        SqlDataReader readerChkInc = sqlCommandChkInc.ExecuteReader();

                                                                        while (readerChkInc.Read())
                                                                        {
                                                                            strIncID = Guid.Parse(readerChkInc["incID"].ToString());

                                                                            strIncidentNumber = readerChkInc["IncNumber"].ToString();
                                                                        }
                                                                        readerChkInc.Close();
                                                                        sqlConnectionChkInc.Close();

                                                                    }
                                                                    catch (Exception exc)
                                                                    {
                                                                        string err = exc.Message.ToString();
                                                                    }
                                                                }





                                                                string strSQLProcedure = "";

                                                                if (strIncidentNumber == strResult)
                                                                {
                                                                    strSQLProcedure = "SPUpdateIncFrmEmail"; // "SPUpdateIncFrmEmail";

                                                                }
                                                                else
                                                                {

                                                                    strSQLProcedure = "SPInsertToIncomingEmail";

                                                                }



                                                                if (strBody.Length > 50)  //  if (strBody.Length > 254)
                                                                {


                                                                    try
                                                                    {
                                                                        strBody = HTMLToText(strHTMLBody);

                                                                    }
                                                                    catch (Exception ex)
                                                                    {
                                                                        strError = ex.Message.ToString();

                                                                    }

                                                                }





                                                                connectionString = ConfigurationManager.AppSettings["ConnectionString"];

                                                                using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                                                {
                                                                    using (SqlCommand cmd = new SqlCommand(strSQLProcedure, con))
                                                                    {
                                                                        try
                                                                        {

                                                                            cmd.CommandType = CommandType.StoredProcedure;

                                                                            cmd.Parameters.AddWithValue("@incID", strIncID);
                                                                            cmd.Parameters.AddWithValue("@inemlID", strIncidentEmailID);
                                                                            cmd.Parameters.AddWithValue("@inemlMailbox", inEmailID);



                                                                            cmd.Parameters.AddWithValue("@inemlSubject", strSubject);
                                                                            cmd.Parameters.AddWithValue("@inemlIsBodyHTML", 1);   //@inemlIsBodyHTML bit,
                                                                            cmd.Parameters.AddWithValue("@inemlHTMLBody", strHTMLBody); //@inemlHTMLBody nvarchar(max),  
                                                                            cmd.Parameters.AddWithValue("@inemlTextBody", strBody);//@inemlTextBody nvarchar(max),  
                                                                            strDateString = strSentDate;





                                                                            cmd.Parameters.AddWithValue("@inemlDateSent", Convert.ToDateTime(strDate)); //@inemlDateAdded datetime,  
                                                                            cmd.Parameters.AddWithValue("@inemlImportance", "Normal");//@inemlImportance nvarchar(50),  
                                                                            cmd.Parameters.AddWithValue("@inemlMessageID", strMessageID); //@inemlMessageID nvarchar(250),  
                                                                            cmd.Parameters.AddWithValue("@inemlFromAddress", strFrom);//@inemlFromAddress nvarchar(320),  
                                                                            cmd.Parameters.AddWithValue("@inemlFromDisplayName", strDisplayName); //@inemlFromDisplayName nvarchar(260),  
                                                                            cmd.Parameters.AddWithValue("@inemlBlockedBySizeLimit", 0); //@inemlBlockedBySizeLimit bit,
                                                                            cmd.Parameters.AddWithValue("@inemlBlockedByBlacklist", 0);//@inemlBlockedByBlacklist bit,  
                                                                            cmd.Parameters.AddWithValue("@inemlBlockedBySpam", 0); //@inemlBlockedBySpam bit,
                                                                            cmd.Parameters.AddWithValue("@inemlFailedDueToError", 0);//@inemlFailedDueToError bit,  
                                                                            cmd.Parameters.AddWithValue("@inemlFailedDueToErrorBeforeDeletion", 0);//@inemlFailedDueToErrorBeforeDeletion bit,
                                                                            cmd.Parameters.AddWithValue("@inemlIsDeleted", 0);//@inemlIsDeleted bit
                                                                            cmd.Parameters.AddWithValue("@inemlFilePath", strFilePath + "\\email.eml");


                                                                            con.Open();
                                                                            cmd.ExecuteNonQuery();
                                                                            // con.Close();


                                                                        }
                                                                        catch (Exception ex)
                                                                        {
                                                                            strError = ex.Message.ToString();
                                                                        }


                                                                    }
                                                                }


                                                                string strfile = "";

                                                                string strEmbeddedFilename, strEmbeddedContentID, strEmbeddedFileFullPath = "";



                                                                // ATTACHMENTS



                                                                //------------------------------


                                                                try
                                                                {


                                                                    if (messages.CurrentPage[i].Attachments != null)
                                                                    {
                                                                        string strAtt = messages.CurrentPage[i].Attachments.ToString();

                                                                        for (int j = 0; j < messages.CurrentPage[i].Attachments.Count; j++)
                                                                        {






                                                                            strattachmentName = messages.CurrentPage[i].Attachments[j].Name;

                                                                            if (messages.CurrentPage[i].Attachments[j].IsInline == true)
                                                                            {


                                                                                strEmbeddedFilename = strattachmentName;

                                                                                strEmbeddedFileFullPath = strFilePath + "\\" + strEmbeddedFilename;


                                                                                var item = (FileAttachment)messages.CurrentPage[i].Attachments[j]; // Cast from Attachment
                                                                                //var folder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                                                                                //var filePath = Path.Combine(strFilePath, item.Name);


                                                                                System.IO.File.WriteAllBytes(strEmbeddedFileFullPath, item.ContentBytes);



                                                                                strEmbeddedContentID = item.ContentId; // messages.CurrentPage[i].Attachments[j].Id;

                                                                                strEmbeddedContentID = strEmbeddedContentID.Replace('<', ' ');
                                                                                strEmbeddedContentID = strEmbeddedContentID.Replace('>', ' ');

                                                                                strEmbeddedContentID = strEmbeddedContentID.Trim();



                                                                                try
                                                                                {

                                                                                    using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                                                                    {
                                                                                        //SPInsertEmbeddedInToAttachments  
                                                                                        using (SqlCommand cmd = new SqlCommand("SPInsertEmbeddedInToAttachments", con))
                                                                                        // using (SqlCommand cmd = new SqlCommand("SPInsertInToAttachments", con))
                                                                                        {
                                                                                            cmd.CommandType = CommandType.StoredProcedure;

                                                                                            cmd.Parameters.AddWithValue("@incID", strIncID);     ////// @incID uniqueidentifier,
                                                                                            cmd.Parameters.AddWithValue("@inemlID", strIncidentEmailID);
                                                                                            cmd.Parameters.AddWithValue("@attDiskPath", strEmbeddedFileFullPath);        //////@attDiskPath varchar(500),
                                                                                            cmd.Parameters.AddWithValue("@attFileName", strEmbeddedFilename);            //////@attFileName varchar(150),
                                                                                            cmd.Parameters.AddWithValue("@inemlattContentID", strEmbeddedContentID);
                                                                                            cmd.Parameters.AddWithValue("@attNotes", "File added by incoming email from " + strFrom); //////@attNotes varchar(250),
                                                                                            cmd.Parameters.AddWithValue("@attSize", messages.CurrentPage[i].Attachments[j].Size);   //////@attSize bigint
                                                                                            cmd.Parameters.AddWithValue("@inemlattEmbeddedImage", 1);


                                                                                            con.Open();
                                                                                            cmd.ExecuteNonQuery();

                                                                                        }
                                                                                    }


                                                                                }
                                                                                catch (Exception exc)
                                                                                {
                                                                                    string err = exc.Message.ToString();
                                                                                }






                                                                            }
                                                                            else
                                                                            {


                                                                                strEmbeddedFileFullPath = strFilePath + "\\" + strattachmentName;


                                                                                var item = (FileAttachment)messages.CurrentPage[i].Attachments[j];

                                                                                System.IO.File.WriteAllBytes(strEmbeddedFileFullPath, item.ContentBytes);

                                                                                using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                                                                {
                                                                                    using (SqlCommand cmd = new SqlCommand("SPInsertInToAttachments", con))
                                                                                    {
                                                                                        cmd.CommandType = CommandType.StoredProcedure;

                                                                                        cmd.Parameters.AddWithValue("@incID", strIncID);     ////// @incID uniqueidentifier,
                                                                                        cmd.Parameters.AddWithValue("@inemlID", strIncidentEmailID);
                                                                                        cmd.Parameters.AddWithValue("@attDiskPath", strFilePath + "\\" + strattachmentName);        //////@attDiskPath varchar(500),
                                                                                        cmd.Parameters.AddWithValue("@attFileName", strattachmentName);            //////@attFileName varchar(150),
                                                                                        cmd.Parameters.AddWithValue("@inemlattContentID", "");
                                                                                        cmd.Parameters.AddWithValue("@attNotes", "File added by incoming email from " + strFrom); //////@attNotes varchar(250),
                                                                                        cmd.Parameters.AddWithValue("@attSize", messages.CurrentPage[i].Attachments[j].Size);   //////@attSize bigint
                                                                                        cmd.Parameters.AddWithValue("@inemlattEmbeddedImage", 0);
                                                                                        con.Open();
                                                                                        cmd.ExecuteNonQuery();

                                                                                    }
                                                                                }
                                                                            }


                                                                        }

                                                                    }

                                                                }
                                                                catch (Exception eee)
                                                                {

                                                                    string strerr = eee.Message.ToString();
                                                                }
                                                                // }



                                                                //  below code to insert into ToAddress

                                                                int toRecepientsCount = messages.CurrentPage[i].ToRecipients.Count();
                                                                for (int torcpt = 0; torcpt < toRecepientsCount; torcpt++)
                                                                {

                                                                    try
                                                                    {
                                                                        using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                                                        {
                                                                            using (SqlCommand cmd = new SqlCommand("spInsertIncomingEmailToAddress", con))
                                                                            {
                                                                                cmd.CommandType = CommandType.StoredProcedure;

                                                                                //@inemlID uniqueidentifier,
                                                                                //@inemlEmailAddress varchar(500),
                                                                                //@inemlDisplayName varchar(500)

                                                                                cmd.Parameters.AddWithValue("@inemlID", strIncidentEmailID);
                                                                                cmd.Parameters.AddWithValue("@inemlEmailAddress", messages.CurrentPage[i].ToRecipients.ElementAt(torcpt).EmailAddress.Address.ToString());
                                                                                cmd.Parameters.AddWithValue("@inemlDisplayName", messages.CurrentPage[i].ToRecipients.ElementAt(torcpt).EmailAddress.Name.ToString());


                                                                                con.Open();
                                                                                cmd.ExecuteNonQuery();

                                                                            }
                                                                        }
                                                                    }

                                                                    catch (Exception eee)
                                                                    {

                                                                        string strerr = eee.Message.ToString();

                                                                        //Insert record in to tblerror
                                                                    }
                                                                }




                                                                //  above code to insert into ToAddress
                                                                //-----------
                                                                // below code to insert into cc addres
                                                                int ccRecepientsCount = messages.CurrentPage[i].CcRecipients.Count();
                                                                for (int ccrcpt = 0; ccrcpt < ccRecepientsCount; ccrcpt++)
                                                                {

                                                                    try
                                                                    {
                                                                        using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                                                        {
                                                                            using (SqlCommand cmd = new SqlCommand("spInsertIncomingEmailCCAddress", con))
                                                                            {
                                                                                cmd.CommandType = CommandType.StoredProcedure;

                                                                                //@inemlID uniqueidentifier,
                                                                                //@inemlEmailAddress varchar(500),
                                                                                //@inemlDisplayName varchar(500)

                                                                                cmd.Parameters.AddWithValue("@inemlID", strIncidentEmailID);
                                                                                cmd.Parameters.AddWithValue("@inemlEmailAddress", messages.CurrentPage[i].CcRecipients.ElementAt(ccrcpt).EmailAddress.Address.ToString());
                                                                                cmd.Parameters.AddWithValue("@inemlDisplayName", messages.CurrentPage[i].CcRecipients.ElementAt(ccrcpt).EmailAddress.Name.ToString());


                                                                                con.Open();
                                                                                cmd.ExecuteNonQuery();

                                                                            }
                                                                        }
                                                                    }

                                                                    catch (Exception eee)
                                                                    {

                                                                        string strerr = eee.Message.ToString();

                                                                        //Insert record in to tblerror
                                                                    }
                                                                }





                                                                //  above code to insert into cc Address






                                                                if (strIncidentNumber != strResult)
                                                                {

                                                                    using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                                                    {
                                                                        using (SqlCommand cmd = new SqlCommand("SPUpdateObjectHstory", con))
                                                                        {
                                                                            cmd.CommandType = CommandType.StoredProcedure;

                                                                            cmd.Parameters.AddWithValue("@incID", strIncID);
                                                                            cmd.Parameters.AddWithValue("@inemlID", strIncidentEmailID);
                                                                            con.Open();
                                                                            cmd.ExecuteNonQuery();

                                                                        }
                                                                    }
                                                                }









                                                                messages.CurrentPage[i].IsRead = true;
                                                                
                                                                //messages.RemoveAt(Convert.ToInt32( messages.CurrentPage[i].Id));





                                                                string strTestSybject = messages.CurrentPage[i].Subject.ToString();

                                                                await graphServiceClient.Users[strEmail].Messages[messages.CurrentPage[i].Id.ToString()].Request().DeleteAsync();


                                                                System.Threading.Thread.Sleep(1000);




                                                                //Below lines commented on 13th SEP 2022
                                                                //GraphServiceClient myClient = GetAuthenticatedClient(clientID, clientPassword, clientTenant);


                                                                //await client.Users[strEmail].Messages[messages.CurrentPage[i].Id.ToString()].Request().DeleteAsync();

                                                                //Above lines commented on 13th Sep 2022



                                                                //  await graphServiceClient.Me.MailFolders.Inbox.Messages.Request().Expand("attachments").GetAsync();

                                                                // await client.Users[strEmail].Messages[messages.CurrentPage[i].InternetMessageId.ToString()].Request().DeleteAsync();












                                                                //  graphClient.Users[clientID]


                                                            }
                                                            catch (Exception eee)
                                                            {

                                                                string strerr = eee.Message.ToString();
                                                            }
                                                        }
                                                    }

                                                }

                                                _ = messages[i].IsRead == true;
                                                strFrom = "";
                                                strCC = "";
                                                strSubject = "";
                                                strHTMLBody = "";
                                                strBody = "";
                                                strMessageID = "";
                                                strSentDate = "";
                                                strReceivedDate = "";
                                                strDisplayName = "";

                                            }









                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine(ex.Message + " in the email :" + strEmail);
                                        }

                                    }

                                    clientDBreader.Close();
                                    sqlClientDBConnection.Close();


                                }
                                catch (Exception ex)
                                {
                                    strError = ex.Message.ToString();




                                }


                            }


                        }

                    }



                    strClientID = "";
                    strClientDatabase = "";
                    strClientConnectionString = "";



                }

                reader.Close();
                sqlConnection.Close();


            }
            catch (Exception ex)
            {
                strError = ex.Message.ToString();



            }
            // THE ABOVE PART OF THE CODE IS USED TO READ EMAILS





            //BELOW PART OF THE CODE IS USED TO SEND EMAILS 
            try
            {
                string emailID, bodyIsHTML;



                sqlConnection.Open();

                strQuery = "select  acctmID , acctmAccountName , acctmConnectionString  from tblmAccount  where acctmAccountName='LiveJepson' order by acctmAccountName ";
                sqlCommand = new SqlCommand(strQuery, sqlConnection);

                SqlDataReader reader = sqlCommand.ExecuteReader();

              

                while (reader.Read())
                {

                    strClientID = reader["acctmID"].ToString();
                    strClientDatabase = reader["acctmAccountName"].ToString();
                    strClientConnectionString = reader["acctmConnectionString"].ToString();
                    SqlCommand sqlClientDBCommand = new SqlCommand();

                  

                    if (strClientID.Length > 0)
                    {


                        if (strClientDatabase.Length > 0)
                        {
                            if (strClientConnectionString.Length > 0)
                            {
                                try
                                {

                                    SqlConnection sqlClientDBConnection = new SqlConnection(strClientConnectionString);

                                    //READ THE RECORDS FROM TBLEMAIL
                                    sqlClientDBConnection.Open();
                                    strQuery = " select top 1 emaid, emasubject , emafromMailbox, emato, emacc, emabcc, emabody,  ";
                                    strQuery = strQuery + "  emaisbodyhtml, emaemaildate, mbname, mbemailaddress , ";
                                    strQuery = strQuery + " mboutgoingprotocol, mboutgoingserveraddress, mbOutgoingPort, ";
                                    strQuery = strQuery + "  mboutgoingpassword, mbOAuth2Name    ,mbOAuth2EmailAddress, ";
                                    strQuery = strQuery + "  mbOAuth2MboxID ,mbOAuth2ClientID      ,mbOAuth2TenantID      ,mbOAuth2Password";
                                    strQuery = strQuery + "  ,(select  mbIncomingPassword  from tblMailbox where mbid=[mbOAuth2MboxID])  as mbPassword";
                                    strQuery = strQuery + "   from vwEmailsNotProcessed";
                                    strQuery = strQuery + " where mbOAuth2ClientID   <>'' and  mbOAuth2TenantID <>''      and mbOAuth2Password <>'' and emato<>''  and emaEmailDate > getdate()-1  order by emaemaildate ";


                                    sqlClientDBCommand.CommandText = strQuery;


                                    sqlClientDBCommand.Connection = sqlClientDBConnection;
                                    SqlDataReader clientDBreader = sqlClientDBCommand.ExecuteReader();

                                   
                                    if(clientDBreader.Read()==true) 
                                    //while (clientDBreader.Read())    
                                    {
                                        emailID = "";
                                        clientID = "";
                                        clientPassword = "";
                                        clientTenant = "";
                                        mailBoxID = "";
                                        strEmail = "";

                                        try
                                        {

                                            emailID = clientDBreader["emaid"].ToString();
                                            clientID = clientDBreader["mbOAuth2ClientID"].ToString();
                                            clientPassword = clientDBreader["mbOAuth2Password"].ToString();
                                            clientTenant = clientDBreader["mbOAuth2TenantID"].ToString();
                                            mailBoxID = clientDBreader["mbOAuth2MboxID"].ToString();
                                            strEmail = clientDBreader["mbOAuth2EmailAddress"].ToString();
                                            strEmailPassword = clientDBreader["mbPassword"].ToString();
                                            //Test -  https://graph.microsoft.com/v1.0/users/research@simplisysltd.onmicrosoft.com/mailFolders/Inbox/messages
                                            //WriteToFile("0 Started " + DateTime.Now.ToString());
                                            //WriteToFile("1  Database :  " + strClientDatabase + "  " + DateTime.Now.ToString());
                                            //WriteToFile("2  EmailID :  " + emailID + "  " + DateTime.Now.ToString());

                                            //UPDATE AS READ IN TBLEMAIL
                                            try
                                            {
                                                using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                                {
                                                    using (SqlCommand cmd = new SqlCommand("spSetEmailAsProcessed", con))
                                                    {
                                                        cmd.CommandType = CommandType.StoredProcedure;

                                                        cmd.Parameters.AddWithValue("@emaID", emailID);

                                                        con.Open();
                                                        cmd.ExecuteNonQuery();
                                                       // WriteToFile("1 Ran spSetEmailAsProcessed :    " + DateTime.Now.ToString());

                                                    }
                                                }

                                            }
                                            catch (Exception ex)
                                            {
                                                strError = ex.Message.ToString();
                                              //  WriteToFile("Error :  " + strError + "  " + DateTime.Now.ToString());

                                            }



                                            //Below lines commented on 13th SEP 2022

                                            //GraphServiceClient client = GetAuthenticatedClient(clientID, clientPassword, clientTenant);

                                            //var users = client.Users.Request().GetAsync().Result;
                                            //Above lines commented on 13th SEP 2022



                                            MGraphServices graphServices = new MGraphServices();

                                            GraphServiceClient client = await graphServices.connectToMailBox(clientID,
                                                 clientTenant,
                                                 strEmail,
                                                 strEmailPassword);













                                            var message = new Message();
                                            MessageAttachmentsCollectionPage MessageAttachmentsCollectionPage = new MessageAttachmentsCollectionPage();

                                            string filePath = "";



                                            //READ ATTACHMENTS

                                            //READ THE RECORDS FROM TBLEMAIL


                                            SqlConnection sqlClienAtttDBConnection = new SqlConnection(strClientConnectionString);
                                            SqlCommand sqlClientAttDBCommand = new SqlCommand();
                                            sqlClienAtttDBConnection.Open();
                                            MessageAttachmentsCollectionPage attachments = new MessageAttachmentsCollectionPage();
                                            try
                                            {

                                                strQuery = "  select   eml.emaid  , emlatt.ema2attEmailID,emlatt.ema2attAttachmentID ,  ";
                                                strQuery = strQuery + "   att.attid, att.attDiskPath , att.attFileName  ";
                                                strQuery = strQuery + "  from tblEmail  eml, tblEmailToAttachment emlAtt , tblAttachment att ";
                                                strQuery = strQuery + "  where  emlAtt.ema2attEmailID = eml.emaid and ";
                                                strQuery = strQuery + "   att.attID = emlAtt.ema2attAttachmentID and  ";
                                                strQuery = strQuery + "  att.attIsDeleted =0 and  eml.emaid= '" + emailID + "'";



                                                sqlClientAttDBCommand.CommandText = strQuery;


                                                sqlClientAttDBCommand.Connection = sqlClienAtttDBConnection;
                                                SqlDataReader clientAttDBreader = sqlClientAttDBCommand.ExecuteReader();

                                                while (clientAttDBreader.Read())
                                                {

                                                    byte[] contentBytes = System.IO.File.ReadAllBytes(@clientAttDBreader["attDiskPath"].ToString());
                                                    //string contentType = "image/png";
                                                    //MessageAttachmentsCollectionPage attachments = new MessageAttachmentsCollectionPage();
                                                    attachments.Add(new FileAttachment
                                                    {
                                                        ODataType = "#microsoft.graph.fileAttachment",
                                                        ContentBytes = contentBytes,
                                                        // ContentType = contentType,
                                                        ContentId = clientAttDBreader["attFileName"].ToString(),
                                                        Name = clientAttDBreader["attFileName"].ToString()
                                                    });




                                                }
                                                clientAttDBreader.Close();
                                                sqlClienAtttDBConnection.Close();

                                            }
                                            catch (Exception ex)
                                            {

                                                strError = ex.Message.ToString();
                                                //WriteToFile("2 Error in attachment read :    " +  strError + "  "+ DateTime.Now.ToString());
                                            }





                                            if (clientDBreader["emaid"].ToString() == "1")
                                            {
                                                bodyIsHTML = "BodyType.Html";
                                            }
                                            else
                                            {
                                                bodyIsHTML = "BodyType.Text";

                                            }

                                            string strToAddress = clientDBreader["emato"].ToString();

                                            var varToAddress = strToAddress.Split(';');
                                            strToAddress = strToAddress.Replace(',', ';');



                                            int countToAddress = varToAddress.Count();


                                            string[] toMail = clientDBreader["emato"].ToString().Split(',');
                                            List<Recipient> toRecipients = new List<Recipient>();
                                            int i = 0;
                                            for (i = 0; i < toMail.Count(); i++)
                                            {
                                                Recipient toRecipient = new Recipient();
                                                EmailAddress toEmailAddress = new EmailAddress();

                                                toEmailAddress.Address = toMail[i];
                                                toRecipient.EmailAddress = toEmailAddress;
                                                toRecipients.Add(toRecipient);
                                            }

                                            List<Recipient> ccRecipients = new List<Recipient>();
                                            if (!string.IsNullOrEmpty(clientDBreader["emacc"].ToString()))
                                            {
                                                string[] ccMail = clientDBreader["emacc"].ToString().Split(',');
                                                int j = 0;
                                                for (j = 0; j < ccMail.Count(); j++)
                                                {
                                                    Recipient ccRecipient = new Recipient();
                                                    EmailAddress ccEmailAddress = new EmailAddress();

                                                    ccEmailAddress.Address = ccMail[j];
                                                    ccRecipient.EmailAddress = ccEmailAddress;
                                                    ccRecipients.Add(ccRecipient);
                                                }
                                            }


                                            List<Recipient> bccRecipients = new List<Recipient>();
                                            if (!string.IsNullOrEmpty(clientDBreader["emabcc"].ToString()))
                                            {
                                                string[] bccMail = clientDBreader["emabcc"].ToString().Split(',');
                                                int j = 0;
                                                for (j = 0; j < bccMail.Count(); j++)
                                                {
                                                    Recipient bccRecipient = new Recipient();
                                                    EmailAddress bccEmailAddress = new EmailAddress();

                                                    bccEmailAddress.Address = bccMail[j];
                                                    bccRecipient.EmailAddress = bccEmailAddress;
                                                    bccRecipients.Add(bccRecipient);
                                                }
                                            }

                                            message = new Message
                                            {
                                                Subject = clientDBreader["emasubject"].ToString(),

                                                Body = new ItemBody
                                                {
                                                    ContentType = BodyType.Html,
                                                    Content = clientDBreader["emabody"].ToString()
                                                },
                                                ToRecipients = toRecipients,
                                                CcRecipients = ccRecipients,
                                                BccRecipients = bccRecipients,
                                                Attachments = attachments

                                            };






                                            //UPDATE AS READ IN TBLEMAIL
                                            try
                                            {
                                                using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                                {
                                                    using (SqlCommand cmd = new SqlCommand("spSetEmailAsProcessed", con))
                                                    {
                                                        cmd.CommandType = CommandType.StoredProcedure;

                                                        cmd.Parameters.AddWithValue("@emaID", emailID);

                                                        con.Open();
                                                        cmd.ExecuteNonQuery();

                                                       // WriteToFile("2 Ran spSetEmailAsProcessed " + DateTime.Now.ToString());

                                                    }
                                                }

                                            }
                                            catch (Exception ex)
                                            {
                                                strError = ex.Message.ToString();

                                            }


                                          

                                            await client.Users[strEmail]
                                            .SendMail(message, false)
                                            .Request().PostAsync();



                                            //BELOW LINE COMMENTED ON 20TH APRIL 2023
                                           // System.Threading.Thread.Sleep(2000);




                                          //  WriteToFile("3  EmailID :  " + emailID + "  - Email sent  " + DateTime.Now.ToString());

                                            
                                            //UPDATE AS READ IN TBLEMAIL
                                            //BELOW LINES COMMENTED ON 20TH APRIL 2023
                                            //try
                                            //{
                                            //    using (SqlConnection con = new SqlConnection(strClientConnectionString))
                                            //    {
                                            //        using (SqlCommand cmd = new SqlCommand("spSetEmailAsProcessed", con))
                                            //        {
                                            //            cmd.CommandType = CommandType.StoredProcedure;

                                            //            cmd.Parameters.AddWithValue("@emaID", emailID);

                                            //            con.Open();
                                            //            cmd.ExecuteNonQuery();

                                            //           // WriteToFile("3 Ran spSetEmailAsProcessed " + DateTime.Now.ToString());

                                            //        }
                                            //    }

                                            //}
                                            //catch (Exception ex)
                                            //{
                                            //    strError = ex.Message.ToString();
                                            // //   WriteToFile(" 3 error  line 1430" + strError+ "   "+ DateTime.Now.ToString());

                                            //}
                                            //ABOVE LINES COMMENTED ON 20TH APRIL 2023



                                           // WriteToFile("4  EmailID :  " + emailID + "  set as processed " + DateTime.Now.ToString());
                                            //WriteToFile("5  EmailID :  " + emailID + "  completed " + DateTime.Now.ToString());
                                           // WriteToFile("------------------------------------------------- " );
                                        }
                                        catch (Exception ex)
                                        {
                                            strError = ex.Message.ToString();
                                            // WriteToFile(" 4 error line 1440" + strError + "   "+ DateTime.Now.ToString());
                                            if (emailID.Length > 0)
                                            {
                                                captureErrorMessage(strClientConnectionString, "Error where sending email", "error when executing send email", strError, strEmail, Guid.Parse(emailID), 1);

                                            }
                                        }
                                    }





                                }
                                catch (Exception ex)
                                {

                                    strError = ex.Message.ToString();
                                     
                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                strError = ex.Message.ToString();
            }
            // THE ABOVE PART OF THE CODE IS USED TO SEND EMAILS




        }

         //Capture error message
        public static void captureErrorMessage( string errorCnnectionString, string customMessage, string internalMessage, string exceptionMessage, string fromData, Guid objectID, int isOutgoingEmail )
        {
                    //@customMessage nvarchar(max),
                    //@internalMessage nvarchar(max),
                    //@exceptionMessage nvarchar(max),
                    //@fromData nvarchar(max),
                    //@objecID uniqueidentifier,
                    //@objectClassName nvarchar(200),
                    //@isOutgoingEmail int

            try
            {

                using (SqlConnection con = new SqlConnection(errorCnnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spInsertEmailError", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;




                        cmd.Parameters.AddWithValue("@customMessage", customMessage);  //@customMessage nvarchar(max),
                        cmd.Parameters.AddWithValue("@internalMessage", internalMessage);//@internalMessage nvarchar(max),
                        cmd.Parameters.AddWithValue("@exceptionMessage", exceptionMessage);//@exceptionMessage nvarchar(max),
                        cmd.Parameters.AddWithValue("@fromData", fromData);//@fromData nvarchar(max),
                        cmd.Parameters.AddWithValue("@objecID", objectID);//@objecID uniqueidentifier,
                        cmd.Parameters.AddWithValue("@objectClassName", "Email");//@objectClassName nvarchar(200),
                        cmd.Parameters.AddWithValue("@isOutgoingEmail", isOutgoingEmail);//@isOutgoingEmail int


                        con.Open();
                        cmd.ExecuteNonQuery();

                    }
                }


            }
            catch (Exception ex)
            {

                string captureErr = ex.Message.ToString();


            }
        }


        public static string HTMLToText(string HTMLCode)
        {
            HTMLCode = HTMLCode.Replace("&nbsp;", " ");
            // Remove new lines since they are not visible in HTML
            HTMLCode = HTMLCode.Replace("\n", " ");
            // Remove tab spaces
            HTMLCode = HTMLCode.Replace("\t", " ");
            // Remove multiple white spaces from HTML
            HTMLCode = Regex.Replace(HTMLCode, "\\s+", " ");
            // Remove HEAD tag
            HTMLCode = Regex.Replace(HTMLCode, "<head.*?</head>", ""
                                , RegexOptions.IgnoreCase | RegexOptions.Singleline);
            // Remove any JavaScript
            HTMLCode = Regex.Replace(HTMLCode, "<script.*?</script>", ""
              , RegexOptions.IgnoreCase | RegexOptions.Singleline);
            // Replace special characters like &, <, >, " etc.
            StringBuilder sbHTML = new StringBuilder(HTMLCode);
            // Note: There are many more special characters, these are just
            // most common. You can add new characters in this arrays if needed
            string[] OldWords = {"&nbsp;", "&amp;", "&quot;", "&lt;",
    "&gt;", "&reg;", "&copy;", "&bull;", "&trade;","&#39;"};
            string[] NewWords = { " ", "&", "\"", "<", ">", "Â®", "Â©", "â€¢", "â„¢", "\'" };
            for (int i = 0; i < OldWords.Length; i++)
            {
                sbHTML.Replace(OldWords[i], NewWords[i]);
            }
            // Check if there are line breaks (<br>) or paragraph (<p>)
            sbHTML.Replace("<br>", "\n<br>");
            sbHTML.Replace("<br ", "\n<br ");
            sbHTML.Replace("<p ", "\n<p ");
            // Finally, remove all HTML tags and return plain text
            return System.Text.RegularExpressions.Regex.Replace(
              sbHTML.ToString(), "<[^>]*>", "");
        }


        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!System.IO.File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = System.IO.File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = System.IO.File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }














        public static SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }


        static public byte[] EncodeTobase64Bytes(byte[] rawData)
        {
            string base64String = System.Convert.ToBase64String(rawData);
            var returnValue = Convert.FromBase64String(base64String);
            return returnValue;
        }


        public static string userToken = null;

        private static GraphServiceClient graphClient = null;







        public static GraphServiceClient GetAuthenticatedClient(string clinetID, string clientPassword, string clientTenant)
        {




            // From app registration registration.
            //const string clientId = "7e948f19-0512-4bcc-867f-7f35e477e7e7";
            //const string password = "L6P8Q~vPEYnGfeNaxrStm_E9fb0biXa2b9NgzcZW";

            //// Form url
            //const string tenantId = "5cd3d2e2-ef76-46ba-9970-775b4fd35977";

            string clientId = clinetID;
            string password = clientPassword;

            // Form url
            string tenantId = clientTenant;


            string getTokenUrl = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            // Form the POST body.
            const string grantType = "client_credentials";
            const string myScopes = "https://graph.microsoft.com/.default"; // Indicates that it should use scopes in the registration.
                                                                            // const string myScopes = " https://graph.microsoft.com/v1.0/users/research@simplisysltd.onmicrosoft.com/mailFolders/Inbox/messages"; // Indicates that it should use scopes in the registration.
            string postBody = $"client_id={clientId}&scope={myScopes}&client_secret={password}&grant_type={grantType}";

            // Create Microsoft Graph client.
            try
            {
                graphClient = new GraphServiceClient(
                    "https://graph.microsoft.com/v1.0",
                    new DelegateAuthenticationProvider(
                        async (requestMessage) =>
                        {
                            // TODO: Create the HttpRequestMessage to request a token for our app.
                            HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, getTokenUrl);
                            httpRequestMessage.Content = new StringContent(postBody, Encoding.UTF8, "application/x-www-form-urlencoded");

                            // TODO: Create the HttpClient, send the request, and get the HttpResponseMessage.
                            HttpClient client = new HttpClient();
                            HttpResponseMessage httpResponseMessage = await client.SendAsync(httpRequestMessage);

                            // TODO: Get the access token from the response and inject the access token into the GraphServiceClient object.
                            string responseBody = await httpResponseMessage.Content.ReadAsStringAsync();
                            userToken = JObject.Parse(responseBody).GetValue("access_token").ToString();
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", userToken);
                        }));

                return graphClient;
            }

            catch (Exception ex)
            {
                Debug.WriteLine("Could not create a graph client: " + ex.Message);
            }

            return graphClient;
        }

        [Obsolete]
        public async Task<GraphServiceClient> CreateGraphClientService(PublicClientApplicationOptions _PublicClientApplicationOptions, string _EmailId, string _Password)
        {
            try
            {
                var pca = PublicClientApplicationBuilder.CreateWithApplicationOptions(_PublicClientApplicationOptions).WithAuthority(AzureCloudInstance.AzurePublic, _PublicClientApplicationOptions.TenantId).Build();



                var authResult = await pca.AcquireTokenByUsernamePassword(new string[] { "https://graph.microsoft.com/.default" }, _EmailId, new NetworkCredential("", _Password).SecurePassword).ExecuteAsync();



                return new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) => { requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken); }));
            }
            catch (Exception ex)
            {
                string exc = ex.Message.ToString();
                throw;
            }
        }












        public static string Between(string STR, string FirstString, string LastString)
        {
            string FinalString;

            STR = STR.Trim();
            FirstString = FirstString.Trim();
            LastString = LastString.Trim();



            try
            {
                int Pos1 = STR.IndexOf(FirstString) + FirstString.Length;
                int Pos2 = STR.IndexOf(LastString);
                FinalString = STR.Substring(Pos1, Pos2 - Pos1);

            }
            catch (Exception ee)
            {

                string strerr = ee.Message.ToString();
                FinalString = "NoMatchFound";
            }

            return FinalString.Trim();




        }

    }
}
