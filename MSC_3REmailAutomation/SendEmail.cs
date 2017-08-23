using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;


namespace MSC_3REmailAutomation
{
    class SendEmail
    {
    //   
        public static void sendExchangeEmail(string strMsg, string strMR_ID, string strMR_Title, string strMR_To, string strMR_CC, int iReminder)
        {
            //String userName = "v-sumgeo@microsoft.com";
            //String password = "United@123";
            //MailMessage msg = new MailMessage();
            //msg.To.Add(new MailAddress("v-sumgeo@microsoft.com"));
            //msg.From = new MailAddress(userName);
            //msg.Subject = "REMINDER: Approval is needed for work completed on Marketing Request " + strMR_ID + ": " + strMR_Title;
            //msg.Body = strPassword;
            //msg.IsBodyHtml = true;
            //SmtpClient client = new SmtpClient();
            //client.Host = "smtp.office365.com";
            //client.Credentials = new System.Net.NetworkCredential(userName, password);
            //client.Port = 587;
            //client.EnableSsl = true;
            //client.Send(msg);
          
            OutlookApp outlookApp = new OutlookApp();
            MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);

            int i = 0;
            switch (iReminder)
            {
                case 2:
                i = 1;   
                break;

                case 5:
                i = 2;
                break;

                case 8:
                i = 3;   
                break;

                default:
                i  = 0;
                break;


            }
            mailItem.Subject = "REMINDER-" + i.ToString() + ": Approval is needed for work completed on Marketing Request " + strMR_ID + ": " + strMR_Title;
            mailItem.HTMLBody = strMsg;

            if (strMR_To.EndsWith(";"))
            {
                strMR_To = strMR_To.Substring(0, strMR_To.Length - 1);
            }

            if (strMR_CC.EndsWith(";"))
            {
                strMR_CC = strMR_CC.Substring(0, strMR_CC.Length - 1);
            }
            //Account acc;
            string[] arrTo = strMR_To.Split(';');
            string[] arrCC = strMR_CC.Split(';');

           string toMsg = "To recipients: " + "\n";
           ExchangeUser currentUser  = mailItem.Application.Session.CurrentUser.AddressEntry.GetExchangeUser();
           //MessageBox.Show(currentUser.PrimarySmtpAddress);
            
            foreach (string word in arrTo)
            {
               Recipient recipTo = mailItem.Recipients.Add(currentUser.PrimarySmtpAddress);
               recipTo.Type = (int)OlMailRecipientType.olTo;
               toMsg += word + "\n";
            }

            MessageBox.Show(toMsg);

            string toCC = "CC recipients: " + "\n";
            
            foreach (string word in arrCC)
            {
                Recipient recipCC = mailItem.Recipients.Add(currentUser.PrimarySmtpAddress);
                recipCC.Type = (int)OlMailRecipientType.olCC;
                toCC += word + "\n";
              
            }

            MessageBox.Show(toCC);

            mailItem.Recipients.ResolveAll();
            mailItem.Send();


        }
        public static DataSet GetDataSet(string ConnectionString, string SQL)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommand cmd = conn.CreateCommand();
            cmd.CommandText = SQL;
            cmd.CommandTimeout = 100;
            da.SelectCommand = cmd;

            DataSet ds = new DataSet();

            conn.Open();
            da.Fill(ds);
            conn.Close();

            return ds;
        }
        public static void AddOutOfOfficeColumn(DataGridView DataGridView1)
        {
            DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
            checkBoxColumn.HeaderText = "Select Email";
            checkBoxColumn.Width = 70;
            checkBoxColumn.Name = "checkBoxColumn";
            
            DataGridView1.Columns.Insert(0, checkBoxColumn);
        }
        public static string getReminderDraft(int i)
        {
            string str = "";
            switch (i)
            {
                case 2:
                str = @"It appears that you haven’t approved the below work we did on your service order(s).  You can click on the “Approve” link(s) below to approve our work, or if this work is no longer needed we will cancel the service order(s) in 6 business days.  Thanks!
            ";
                break;

                case 5:
                str = @"It appears that you haven’t approved the below work we did on your service order(s).  You can click on the “Approve” link(s) below to approve our work, or if this work is no longer needed we will cancel the service order(s) in 3 business days.  Thanks!";
                break;

                case 8:
                str = @"We haven’t heard from you regarding the status of this request so we’ll cancel it shortly.  Thanks!";
                break;
               
                default:
                str = "Test";
                break;
            }
            return str; 

        }

        public static string getEmail(int i, string strMarketingRequestName, string strMR_ID, string strMST_ID, string strMS_ID, string strMR_Title,string strTitle, string strEmailAddress, string strSubmission,string strDueDate, string strCountry, string strArea,string strProgram,string strServiceType , string strCampaign, string strLink )
         {
             string str = @"
             <!DOCTYPE html><html xmlns='http://www.w3.org/1999/xhtml'><head>
            <meta http-equiv='Content-Type' content='text/html; charset=us-ascii'><title></title>
    
                <meta name='viewport' content='width=device-width; initial-scale=1.0;'>
                <meta http-equiv='X-UA-Compatible' content='IE=edge'>


                <style type='text/css'>
                    a {
                        color: #1a86cd;
                    }

                    body {
                        margin: 0;
                        padding: 0;
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                    }

                    img {
                        border: 0;
                        height: auto;
                        line-height: 100%;
                        outline: none;
                        text-decoration: none;
                    }

                    table {
                        border-collapse: collapse !important;
                    }

                    body {
                        font-size: 14px;
                        color: #2a2a2a;
                    }

                    th {
                        text-align: left;
                    }

                    #outlook a {
                        padding: 0;
                    }

                    .ExternalClass {
                        width: 100%;
                    }
                        /* Force Hotmail to display emails at full width */
                        .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {
                            line-height: 100%;
                        }
                    /* Force Hotmail to display normal line spacing */

                    .emailTitleStyle {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 23px;
                        width: 700px;
                        color: #fff;
                        background-color: #1a86cd;
                        padding-left: 10px;
                        height: 40px;
                    }

                    .emailTitleStyle2 {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 18px;
                        font-weight: bold;
                        vertical-align: bottom;
                        color: #2a2a2a;
                        padding-left: 10px;
                        padding: 12px 0 5px 0;
                        background-color: #fff;
                    }
	            /*New css class added for the logo changes by Pankaj START*/
                    .emailTitleRightAlignBoldStyle {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        width: 700px;
                        color: #fff;
                        background-color: #1a86cd;
                        height: 40px;
                        font-size: 20px;
                        font-weight: bold;
                        padding-right: 14px;
                        text-align:right;
                    }
                    .emailTitleRightAlignNormalStyle {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        width: 700px;
                        color: #fff;
                        background-color: #1a86cd;
                        height: 40px;
                        font-size: 14px;
                        padding-right: 13px;
                        text-align:right;
                    }
                    /*New css class added for the logo changes by Pankaj ENDS*/
                    .copy1 {
                        font-family: 'Segoe UI', Verdana, Arial, Helvetica, sans-serif;
                        font-size: 14px;
                        color: #2a2a2a;
                        padding: 20px 0 10px 10px;
                    }

                    .linkStyle {
                        font-size: 14px;
                        text-decoration: none;
                        color: #1a86cd;
                    }

                    .dataStyle {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 14px;
                        color: #2a2a2a;
                    }

                    .serviceHeader {
                        color: #555555;
                        font-size: 12px;
                    }

                    .percentComplete {
                        width: 37%;
                        font-size: 30px;
                        /*float: left;*/
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        color: #555555;
                        vertical-align: bottom;
                    }

                    .serviceHeader {
                        color: #555555;
                        font-size: 12px;
                    }

                    .title-width {
                        width: 30%;
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 11px;
                        color: #555555;
                        float: left;
                    }

                    .title-width2 {
                        width: 65%;
                        float: left;
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 11px;
                    }

                    .notes {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 25px;
                        font-weight: bold;
                        color: #2a2a2a;
                        padding: 10px 0 0 10px;
                    }

                    .notes2 {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 14px;
                        color: #2a2a2a;
                        word-wrap: break-word;
                        padding: 0 0 10px 10px;
                    }

                    .footer_txt {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 15px;
                        color: #2a2a2a;
                    }

                    .ratings {
                        font-family: Segoe UI;
                        font-size: 10px;
                        word-wrap: break-word;
                        color: #5075B6;
                    }

                    .ratings2 {
                        font-family: Segoe UI;
                        font-size: 10px;
                        word-wrap: break-word;
                        color: #5075B6;
                        height: 40px;
                        background-color: #D9D9D9;
                    }

                    .header_txt {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 30px;
                        color: #000;
                    }

                    .header_txt2 {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 22px;
                        color: #000;
                    }

                    .so_txt {
                        vertical-align: top;
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 14px;
                    }

                    .so_margin {
                        float: left;
                        height: 50px;
                    }

                    .so_margin2 {
                        float: left;
                        height: 50px;
                    }

                    .so_margin3 {
                        height: 50px;
                    }

                    .bar_color {
                        width: 14px;
                        height: 14px;
                        background-color: rgb(0, 171, 70);
                        float: left;
                        font-size: 1px;
                        line-height: 1px;
                        margin-left: 1px;
                        padding-left: 1px;
                        vertical-align: bottom;
                    }

                    .emailFooterStyle {
                        font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                        font-size: 10px;
                        width: 700px;
                        color: #2a2a2a;
                    }

                    .Hide {
                        display:none;
                    }

                    .Disabled {
                        pointer-events: none;
                        cursor: default;
                    }

                    .bannerImage img {
			            display: block;
			            max-width:290px;
			            max-height:190px;
			            width: auto;
			            height: auto;
		            }

                    /* MOBILE STYLES */
                    @media (max-width: 767px) {
                        body {
                            -webkit-text-size-adjust: 100%;
                        }
                        /* FULL-WIDTH TABLES */
                        table[class='responsive-table'] {
                            width: 100% !important;
                        }

                        td[class='responsive-td'] {
                            width: 100% !important;
                            float:left;
                        }

                        td[class='responsive-td-paragraph'] {
                            width: 99% !important;
                        }

                        td[class='responsive-td-paragraph2'] {
                            width: 100% !important;
                            padding-left: 0px !important;
                            float:left;
                        }

                        td[class='responsive-td-paragraph3'] {
                            padding-bottom: 8px !important;
                        }

                        td[class='extr_td'] {
                            width: 0px;
                        }

                        td[class='title-width'] {
                            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                            font-size: 11px;
                            color: #555555;
                            float: left;
                        }

                        td[class='device_hide'] {
                            display: block;
                        }

                        .percentComplete {
                            width:40%;
                            padding-left:10px; 
                        }

                        table[class=responsive-table-header] {
                            width: 100% !important;
                        }

                        table[class='percent-margin'] {
                            margin-left: 10px !important;
                        }

                        .so_margin {
                            float: left;
                            height: 30px;
                        }

                        .so_margin2 {
                            float: left;
                            height: 40px;
                        }

                        .so_margin3 {
                            height: 30px;
                            vertical-align: top;
                        }

                        .banner {
                            padding: 10px 0 0 0px;
                        }
                    }

                    @media (min-width: 768px) and (max-width: 1280px) {
                        td[class='responsive-td-paragraph2'] {
                            width: 960px !important;
                            padding-left: 0px !important;
                            float:left;
                        }

                        td[class='responsive-td'] {
                            float:left;
                        }

                        td[class='responsive-tablet'] {
                            width: 960px !important;
                        }

                        table[class='responsive-table'] {
                            width: 960px !important;
                        }

                        table[class=responsive-table-header] {
                            width: 962px !important;
                        }

                        table[class='tablet-left-content'] {
                            width: 20% !important;
                        }

                        table[class='tablet-right-content'] {
                            width: 65% !important;
                            padding-top: 10px;
                        }
                    }
                </style>

            </head>

            <body>
                <table border='0' cellpadding='0' cellspacing='0' width='100%' bgcolor='#d7dfc9' style='background-color: #F0F0F0'>
                    <tr>
                        <td align='center'>
                            <table width='1280' border='0' cellpadding='0' cellspacing='0' class='responsive-table'>
                                <tr>
                                    <td style='padding: 0px 0px 20px;'>
                                        <table width='100%' border='0' cellpadding='0' cellspacing='0' style='border-left: 1px #ffffff solid; border-right: 1px #ffffff solid; border-top: 1px #ffffff solid' class='responsive-table-header'>
                                            <tr style='vertical-align: bottom; background-color: rgb(26, 134, 205);'>
                                                <td class='emailTitleRightAlignBoldStyle'>Marketing Services</td>
                                            </tr>
                                            <tr style='vertical-align: top; background-color: rgb(26, 134, 205);'>
                                                <td class='emailTitleRightAlignNormalStyle'>Marketing IT / GMO</td>
                                            </tr>
                                             <tr style='vertical-align: top; background-color: rgb(26, 134, 205);'>
                                                <td class='emailTitleStyle'>REMINDER: Approval Requested</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                    
                                <tr>
                                    <td>
                                        <table width='100%' border='0' cellpadding='0' cellspacing='0' class='responsive-table' style='border-left: 1px #ffffff solid; border-right: 1px #ffffff solid'>
                                            <tr>
                                                <td>
                                                    <table width='100%' cellspacing='0' cellpadding='0' bgcolor='#ffffff' style='border-collapse: collapse;'>
                                                        <tr>
                                                            <td valign='top' width='958' class='responsive-td' style='border-right:1px solid #d9d9d9;'>
                                                                <table width='100%' bgcolor='#ffffff' cellspacing='0' cellpadding='0' border='0' class='responsive-table'>
                                                                    <tr>
                                                                        <td align='center'>
                                                                            <table bgcolor='#ffffff' cellspacing='0' cellpadding='0' width='98%'>
                                                                                <tr>
                                                                                    <td colspan='2' style='padding-bottom: 10px;' class='emailTitleStyle3'>Hello there,<br><br>" + SendEmail.getReminderDraft(i) + @"</td>
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>

                                                                    <tr>
                                                                        <td bgcolor='#fafafa' style='padding: 10px'>
                                                                            <table width='100%' cellpadding='0' cellspacing='0'>
                                                                                <tr>
                                                                                    <td width='52%' align='left' style='float: left; padding-top: 10px; padding-bottom: 10px' class='responsive-td'>
                                                                                        <table style='width: 100%; border-collapse: collapse;' cellpadding='0' cellspacing='0'>
                                                                                            <tr>
                                                                                                <td colspan='2'><a class='linkStyle' href='http://execservicessf/Runtime/Runtime/Form/MarketingRequestDetailsEditForm/?MarketingRequestId=" + strMR_ID + @"'>" + strMR_Title + @"</a></td>
                                                                                            </tr>
                                                                                            <tr>
                                                                                                <td class='title-width'>Request Id</td>
                                                                                                <td style='width: 50%; float: left; font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif; font-size: 11px; color: #1a86cd; text-decoration: none;'>
                                                                                                    <a style='text-decoration: none;' href='http://execservicessf/Runtime/Runtime/Form/MarketingRequestDetailsEditForm/?MarketingRequestId=" + strMR_ID + @"'>" + strMR_ID + @"</a>
                                                                                                </td>
                                                                                            </tr>
                                                                                            <tr>
                                                                                                <td class='title-width'>MSC Contact</td>

                                                                                                <td style='width: 50%; float: left; font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif; font-size: 11px; color: #1a86cd; text-decoration: none;'>
                                                                                                    <a style='text-decoration: none;' href=" + "mailto:" + strEmailAddress +  @">" + strEmailAddress + @"</a>
                                                                                                </td>
                                                                                            </tr>
                                                                                            <tr>
                                                                                                <td class='title-width'>Campaign</td>
                                                                                                <td class='title-width2'>" + strCampaign +@" </td>
                                                                                            </tr>
                                                                                            <tr>
                                                                                                <td class='title-width'>Area</td>
                                                                                                <td class='title-width2'>" + strArea + @" </td>
                                                                                            </tr>
                                                                                            <tr>
                                                                                                <td class='title-width'>Country</td>
                                                                                                <td class='title-width2'>" + strCountry +@" </td>
                                                                                            </tr>
                                                                                            <tr>
                                                                                                <td class='title-width'>Program</td>
                                                                                                <td class='title-width2'>" + strProgram + @"</td>
                                                                                            </tr>
                                                                                        </table>
                                                                                    </td>

                                                                                    <td width='30%' style='background-color: #fafafa; float: right; padding-top: 30px' class='responsive-td'>
                                                                                        <table border='0' width='180px' cellpadding='0' cellspacing='0'>
                                                                                            <tr>
                                                                                                <td class='percentComplete'>67%</td>
                                                                                                <td width='50%' class='dataStyle' style='margin-top: 5px; padding-left: 32px;font-size: 11px; color: #555555;'>request
                                                                                                    <br>
                                                                                                    completed
                                                                                                </td>
                                                                                            </tr>
                                                                                            <tr>
                                                                                                <td colspan='2'>
                                                                                                    <table width='100%' class='percent-margin'>
                                                                                                        <tr style='width: 100%; background-color: rgb(232, 232, 232); height: 20px;'>
                                                                                                            <td>
                                                                                                                <table style='width: 83%; background-color: rgb(219, 219, 219); height: 18px;'></table>
                                                                                                            </td>
                                                                                                        </tr>
                                                                                                    </table>
                                                                                                </td>
                                                                                            </tr>
                                                                                        </table>
                                                                                    </td>

                                                                                    <!--<td width='30%' style=' background-color: #fafafa;float:right;padding-top:30px' class='responsive-td'>
                                                                                        <table border='0' width='80%' cellpadding='0' cellspacing='0' >
                                                                                            <tr>
                                                                                                <td align='left' class='percentComplete' style='padding-left: 10px;'>83%</td>
                                                                                                <td align='left' width='50%' class='dataStyle' style='margin-top: 5px; float: right; font-size: 11px; color: #555555;'>request
                                                                                                    <br />
                                                                                                    completed
                                                                                                </td>
                                                                                            </tr>
                                                                                            <tr>
                                                                                                <td colspan='2'>
                                                                                                    <table width='80%' class='percent-margin'>
                                                                                                        <tr style='width: 100%; background-color: rgb(232, 232, 232); height: 20px;'>
                                                                                                            <td>
                                                                                                                <table style='width: 83%; background-color: rgb(219, 219, 219); height: 18px;'></table>
                                                                                                            </td>
                                                                                                        </tr>
                                                                                                    </table>
                                                                                                </td>
                                                                                            </tr>
                                                                                        </table>
                                                                                    </td>-->

                                                                        
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>

                                                                    <tr>
                                                                        <td>
                                                                            <table bgcolor='#ffffff' width='100%' border='0' cellpadding='0' cellspacing='0' class='responsive-table'>
                                                                                <tr>
                                                                                    <td style='font-weight: bold; padding: 10px 0 0 10px'>My Service Orders</td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td style='padding-left: 10px'>
                                                                                        <table width='100%' border='0' cellpadding='0' cellspacing='0'>
                                                                                            <tr>
                                                                                        <td width='25%' class='responsive-td' style='float: left;vertical-align: top;'>
                                                                                            <table width='100%' border='0' cellpadding='0' cellspacing='0'>
                                                                                                <tr>
                                                                                                    <td width='100%' class='responsive-td so_margin2' style='padding-bottom:20px'>
                                                                                                        <table border='0' cellpadding='0' cellspacing='0' width='100%'>
                                                                                                            <tr>
                                                                                                                <td style='vertical-align: top; color: #1a86cd; font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif; font-size: 14px;'>
                                                                                                                    <span class='serviceHeader'>Service Order Name: <br></span><span><a style='text-decoration: none;' href='http://execservicessf/Runtime/Runtime/Form/MarketingRequestDetailsEditForm/?MarketingRequestId=" + strMR_ID + @"&amp;MarketingServiceId=" + strMS_ID + @"'>" + strServiceType + "</br> (" + strTitle + ")" + @"</a></span>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </td>
                                                                                        <td width='15%' class='responsive-td' style='float: left;vertical-align: top;'>
                                                                                            <table width='100%' border='0' cellpadding='0' cellspacing='0'>
                                                                                                <tr>
                                                                                                    <td width='100%' class='responsive-td so_margin' style='padding-bottom:20px'>
                                                                                                        <table border='0' cellpadding='0' cellspacing='0' width='100%'>
                                                                                                            <tr>
                                                                                                                <td> <span class='serviceHeader'>Service Order ID: <br></span><span class='so_txt' style=' display:@@show;'>" + strMS_ID + "/" + strMST_ID + @"</span>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </td>
                                                                                        <td width='12%' class='responsive-td' style='float: left;vertical-align: top;'>
                                                                                            <table width='100%' border='0' cellpadding='0' cellspacing='0'>
                                                                                                <tr>
                                                                                                    <td width='100%' class='responsive-td so_margin' style='padding-bottom:20px'>
                                                                                                        <table border='0' cellpadding='0' cellspacing='0' width='100%'>
                                                                                                            <tr>
                                                                                                                <td><span class='serviceHeader'>Submitted: <br></span><span class='so_txt' title='" + strSubmission + @"'>" + DateTime.Parse(strSubmission).ToShortDateString() + @"</span>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </td>
                                                                                        <td width='12%' class='responsive-td' style='float: left;vertical-align: top;'>
                                                                                            <table width='100%' border='0' cellpadding='0' cellspacing='0'>
                                                                                                <tr>
                                                                                                    <td width='100%' class='responsive-td so_margin' style='padding-bottom:20px'>
                                                                                                        <table border='0' cellpadding='0' cellspacing='0' width='100%'>
                                                                                                            <tr>
                                                                                                                <td><span class='serviceHeader'>Est. Delivery: <br></span><span class='so_txt' title='" + strDueDate + @"'>" + DateTime.Parse(strDueDate).ToShortDateString() + @"</span>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>

                                                                                        </td>
                                                                                        <td width='15%' class='responsive-td' style='float: left;vertical-align: top;'>
                                                                                            <table width='100%' border='0' cellpadding='0' cellspacing='0'>
                                                                                                <tr>
                                                                                                    <td width='100%' class='responsive-td so_margin' style='padding-bottom:20px'>
                                                                                                        <table border='0' cellpadding='0' cellspacing='0' width='100%' ?=''>
                                                                                                            <tr>
                                                                                                                <td><span class='serviceHeader'>Status: <br></span><a style='text-decoration: none;' href='" + strLink + @"'>Approve</a></span>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </td>
                                                                                        <td width='12%' class='responsive-td' style='float: left;vertical-align: top;'>
                                                                                            <table width='100%' border='0' cellpadding='0' cellspacing='0'>
                                                                                                <tr>
                                                                                                    <td width='100%' class='responsive-td so_margin' style='padding-bottom:20px'>
                                                                                                        <table border='0' cellpadding='0' cellspacing='0' width='100%'>
                                                                                                            <tr>
                                                                                                                <td class='so_txt' style='padding-top: 5px'>
                                                                                                                    <table cellspacing='0' cellpadding='0' border='0'>
                                                                                                                                <tr>
                                                                                                                                    <td style='width: 14px; height: 12px; background-color: rgb(219, 219, 219); float: left; margin-top: 2px; margin-right: 4px; font-size: 1px; line-height: 1px;'>&nbsp;</td><td width='1px'>&nbsp;</td>
                                                                                                                                    <td style='width: 14px; height: 12px; background-color: rgb(219, 219, 219); float: left; margin-top: 2px; margin-right: 4px; font-size: 1px; line-height: 1px;'>&nbsp;</td><td width='1px'>&nbsp;</td>
                                                                                                                                    <td style='width: 14px; height: 12px;background-color: rgb(219, 219, 219);float: left;margin-top: 2px;margin-right: 4px; font-size:1px;line-height:1px;'>&nbsp;</td><td width='1px'>&nbsp;</td>
                                                                                                                                    <td style='width: 14px; height: 12px; background-color: rgb(219, 219, 219); float: left; margin-top: 2px; margin-right: 4px; font-size: 1px; line-height: 1px;'>&nbsp;</td><td width='1px'>&nbsp;</td>
                                                                                                                                    <td style='width: 14px; height: 12px; background-color: rgb(128,0,128); float: left; margin-top: 2px; margin-right: 4px; font-size: 1px; line-height: 1px;'>&nbsp;</td><td width='1px'>&nbsp;</td>
                                                                                                                                    <td style='width: 14px; height: 12px;background-color: rgb(219, 219, 219);float: left;margin-top: 2px;margin-right: 4px; font-size:1px;line-height:1px;'></td>
                                                                                                                                </tr>
                                                                                                                            </table>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </td>
                                                                                        </tr>
                                                                                            <tr>
                                                                                                <td height='50' style='font-style: italic; font-weight: lighter; font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif; font-size: 14px; color: #2a2a2a; padding-left: 10px'>&nbsp;</td>
                                                                                            </tr>
                                                                                        </table>
                                                                                    </td>
                                                                                </tr>
                                                                                <tr class='dataStyle'>
                                                                                    <td style='padding: 0 010px 10px; margin-right: 30px;' class='serviceHeader'>
                                                                                        <br>
                                                                                       <span class='@@ShowOrNot'>Approve Comments:&nbsp;&nbsp;&nbsp; </span> 
                                                                                    </td>                                                                        
                                                                                </tr>
                                                                                <tr>
                                                                                    <td height='50' style='font-style: italic; font-weight: lighter; font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif; font-size: 14px; color: #2a2a2a; padding-left: 10px'>&nbsp;</td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td height='50' style='font-style: italic; font-weight: lighter; font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif; font-size: 14px; color: #2a2a2a; padding-left: 10px'>&nbsp;</td>
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>

                                                                    <!--<tr>
                                                            <td>
                                                                <table border='0' cellpadding='0' cellspacing='0' width='100%' bgcolor='#d7dfc9' style='background-color: #F0F0F0'>
                                                                    <tr>
                                                                        <td valign='top' align='center'>
                                                                            <table width='100%' border='0' cellpadding='0' cellspacing='0' class='responsive-table'>
                                                                                <tr>
                                                                                    <td class='notes'>Notes</td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td class='notes2'><span style='font-weight: bold; font-family: segoe ui, verdana, arial, helvetica, sans-serif;'>@@LatestNotesCreator</span>
                                                                                        @@Notes</td>
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </td>
                                                        </tr>-->

                                                                </table>
                                                            </td>
                                                            <td width='310' valign='top' bgcolor='#ffffff' class='responsive-td-paragraph2' style='border-left:1px solid #d9d9d9; padding-left: 5px'>
                                                                <table align='left' border='0' cellpadding='0' cellspacing='0' width='100%' bgcolor='#d7dfc9' style='background-color: #ffffff'>

                                                                    <tr>
                                                                        <td bgcolor='#ffffff'>
                                                                            <table class='tablet-left-content' style='float: left' width='100%' cellpadding='0' cellspacing='0' border='0'>
                                                                                <tr>
                                                                                    <td align='center' style='padding: 14px 10px' class='bannerImage'>
                                                                                    <!--    <span style='@@BannerWithLink'><a href='@@BannerLink'><img src='@@BannerUrl' class='img-max' width='290' height='190'/></a></span>
                                                                                        <span style='@@BannerWithoutLink'> <img src='@@BannerUrl' class='img-max' width='290' height='190'/></span> -->
																			            <a href='http://aka.ms/msc-getting-started'><img src='http://execservicessf/Runtime/Styles/Themes/Marketing%20MTTA/Images/Tips-n-Tricks.jpg' width='290' height='190'></a>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                            <table class='tablet-right-content' style='float: left' width='100%'>
                                                                                <tr>
                                                                                    <td align='left' style='padding: 10px' class='header_txt2'></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td style='padding: 0 10px 0 10px'><p><span style='font-size: 18px;'>Did you know?</span></p><p>Getting started with the MSC services? Please visit <a href='http://aka.ms/msc-getting-started'>http://aka.ms/msc-getting-started</a>.</p>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td width='1280' class='responsive-tablet' style='background-color: #f0f0f0'>&nbsp;</td>
                    </tr>
                    <tr>
                        <td>
                            <table border='0' cellpadding='0' cellspacing='0' width='100%' style='background-color: #ffffff' bgcolor='#ffffff'>
                                <tr>
                                    <td>
                                        <table width='1280' align='center' border='0' cellpadding='0' cellspacing='0' class='responsive-table'>
                                            <tr>
                                                <td align='right' style='padding: 10px'>
                                                    <img style='margin-top: 15px;' src='http://execservicessf/Runtime/Styles/Themes/Marketing MTTA/images/MSLogo.jpg' alt='MS'>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </body>
            </html>


            ";
             return str; 
         }
  
     
        public static string getCSS()
        {
            string str = @"
        <head>    
        <style type='text/css'>
        a {
            color: #1a86cd;
        }

        body {
            margin: 0;
            padding: 0;
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
        }

        img {
            border: 0;
            height: auto;
            line-height: 100%;
            outline: none;
            text-decoration: none;
        }

        table {
            border-collapse: collapse !important;
        }

        body {
            font-size: 14px;
            color: #2a2a2a;
        }

        th {
            text-align: left;
        }

        #outlook a {
            padding: 0;
        }

        .ExternalClass {
            width: 100%;
        }
            /* Force Hotmail to display emails at full width */
            .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {
                line-height: 100%;
            }
        /* Force Hotmail to display normal line spacing */

        .emailTitleStyle {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 23px;
            width: 700px;
            color: #fff;
            background-color: #1a86cd;
            padding-left: 10px;
            height: 40px;
        }

        .emailTitleStyle2 {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 18px;
            font-weight: bold;
            vertical-align: bottom;
            color: #2a2a2a;
            padding-left: 10px;
            padding: 12px 0 5px 0;
            background-color: #fff;
        }
	/*New css class added for the logo changes by Pankaj START*/
        .emailTitleRightAlignBoldStyle {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            width: 700px;
            color: #fff;
            background-color: #1a86cd;
            height: 40px;
            font-size: 20px;
            font-weight: bold;
            padding-right: 14px;
            text-align:right;
        }
        .emailTitleRightAlignNormalStyle {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            width: 700px;
            color: #fff;
            background-color: #1a86cd;
            height: 40px;
            font-size: 14px;
            padding-right: 13px;
            text-align:right;
        }
        /*New css class added for the logo changes by Pankaj ENDS*/
        .copy1 {
            font-family: 'Segoe UI', Verdana, Arial, Helvetica, sans-serif;
            font-size: 14px;
            color: #2a2a2a;
            padding: 20px 0 10px 10px;
        }

        .linkStyle {
            font-size: 22px;
            text-decoration: none;
            color: #1a86cd;
        }

        .dataStyle {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 14px;
            color: #2a2a2a;
        }

        .serviceHeader {
            color: #555555;
            font-size: 12px;
        }

        .percentComplete {
            width: 37%;
            font-size: 30px;
            /*float: left;*/
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            color: #555555;
            vertical-align: bottom;
        }

        .serviceHeader {
            color: #555555;
            font-size: 12px;
        }

        .title-width {
            width: 30%;
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 11px;
            color: #555555;
            float: left;
        }

        .title-width2 {
            width: 65%;
            float: left;
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 11px;
        }

        .notes {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 25px;
            font-weight: bold;
            color: #2a2a2a;
            padding: 10px 0 0 10px;
        }

        .notes2 {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 14px;
            color: #2a2a2a;
            word-wrap: break-word;
            padding: 0 0 10px 10px;
        }

        .footer_txt {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 15px;
            color: #2a2a2a;
        }

        .ratings {
            font-family: Segoe UI;
            font-size: 10px;
            word-wrap: break-word;
            color: #5075B6;
        }

        .ratings2 {
            font-family: Segoe UI;
            font-size: 10px;
            word-wrap: break-word;
            color: #5075B6;
            height: 40px;
            background-color: #D9D9D9;
        }

        .header_txt {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 30px;
            color: #000;
        }

        .header_txt2 {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 22px;
            color: #000;
        }

        .so_txt {
            vertical-align: top;
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 14px;
        }

        .so_margin {
            float: left;
            height: 50px;
        }

        .so_margin2 {
            float: left;
            /*height: 50px;*/
        }

        .so_margin3 {
            height: 50px;
        }

        .bar_color {
            width: 14px;
            height: 14px;
            background-color: rgb(0, 171, 70);
            float: left;
            font-size: 1px;
            line-height: 1px;
            margin-left: 1px;
            padding-left: 1px;
            vertical-align: bottom;
        }

        .emailFooterStyle {
            font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
            font-size: 10px;
            width: 700px;
            color: #2a2a2a;
        }

        .Hide {
            display: none;
        }

        .Disabled {
            pointer-events: none;
            cursor: default;
        }

        .bannerImage img {
            display: block;
            max-width: 290px;
            max-height: 190px;
            width: auto;
            height: auto;
        }

        #LocalizationShow {
            display: none;
        }

        #LocalizationShow {
            display: block;
        }

        /* MOBILE STYLES */
        @media (max-width: 767px) {
            body {
                -webkit-text-size-adjust: 100%;
            }
            /* FULL-WIDTH TABLES */
            table[class='responsive-table'] {
                width: 100% !important;
            }

            td[class='responsive-td'] {
                width: 100% !important;
                float: left;
            }

            td[class='responsive-td-paragraph'] {
                width: 99% !important;
            }

            td[class='responsive-td-paragraph2'] {
                width: 100% !important;
                padding-left: 0px !important;
                float: left;
            }

            td[class='responsive-td-paragraph3'] {
                padding-bottom: 8px !important;
            }

            td[class='extr_td'] {
                width: 0px;
            }

            td[class='title-width'] {
                font-family: Segoe UI, Verdana, Arial, Helvetica, sans-serif;
                font-size: 11px;
                color: #555555;
                float: left;
            }

            td[class='device_hide'] {
                display: block;
            }

            .percentComplete {
                width: 40%;
                padding-left: 10px;
            }

            table[class=responsive-table-header] {
                width: 100% !important;
            }

            table[class='percent-margin'] {
                margin-left: 10px !important;
            }

            .so_margin {
                float: left;
                height: 30px;
            }

            .so_margin2 {
                float: left;
                /*height: 40px;*/
            }

            .so_margin3 {
                height: 30px;
                vertical-align: top;
            }

            .banner {
                padding: 10px 0 0 0px;
            }
        }

        @media (min-width: 768px) and (max-width: 1280px) {
            td[class='responsive-td-paragraph2'] {
                width: 960px !important;
                padding-left: 0px !important;
                float: left;
            }

            td[class='responsive-td'] {
                float: left;
            }

            td[class='responsive-tablet'] {
                width: 960px !important;
            }

            table[class='responsive-table'] {
                width: 960px !important;
            }

            table[class=responsive-table-header] {
                width: 962px !important;
            }

            table[class='tablet-left-content'] {
                width: 20% !important;
            }

            table[class='tablet-right-content'] {
                width: 65% !important;
                padding-top: 10px;
            }
                }
            </style>
            </head>";
            return str;
        }
        public static string getHeader()
        {
            string str = @"
            <body>
            <table border='0' cellpadding='0' cellspacing='0' width='100%' bgcolor='#d7dfc9' style='background-color: #F0F0F0'>
                <tr>
                    <td align='center'>
                        <table width='1280' border='0' cellpadding='0' cellspacing='0' class='responsive-table'>
                            <tr>
                                <td style='padding: 0px 0px 20px;'>
                                    <table width='100%' border='0' cellpadding='0' cellspacing='0' style='border-left: 1px #ffffff solid; border-right: 1px #ffffff solid; border-top: 1px #ffffff solid' class='responsive-table-header'>
                                        <tr style='vertical-align: bottom; background-color: rgb(26, 134, 205);'>
                                            <td class='emailTitleRightAlignBoldStyle'>Marketing Services</td>
                                        </tr>
                                        <tr style='vertical-align: top; background-color: rgb(26, 134, 205);'>
                                            <td class='emailTitleRightAlignNormalStyle'>Marketing IT / GMO</td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td width='100%' style='width: 100%'>
                                    <table width='100%' border='0' cellpadding='0' cellspacing='0' style='border-left: 1px #ffffff solid; border-right: 1px #ffffff solid; border-top: 1px #ffffff solid' class='responsive-table-header'>
                                        <tr>
                                            <td class='emailTitleStyle'>Reminder Email</td>
                                        </tr>
                                        <tr>
                                            <td bgcolor='#ffffff' style='background-color: #ffffff; height: 5px;'>&nbsp;</td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
            ";
            return str;
        }
        public static string getName(string strName)
        {
            string str = @"
              <tr>
                <td>
                    <table width='100%' border='0' cellpadding='0' cellspacing='0' class='responsive-table' style='border-left: 1px #ffffff solid; border-right: 1px #ffffff solid'>
                        <tr>
                            <td>
                                <table width='100%' cellspacing='0' cellpadding='0' bgcolor='#ffffff' style='border-collapse: collapse;'>
                                    <tr>
                                        <td valign='top' width='958' class='responsive-td' style='border-right:1px solid #d9d9d9;'>
                                            <table width='100%' bgcolor='#ffffff' cellspacing='0' cellpadding='0' border='0' class='responsive-table'>
                                                <tr>
                                                    <td align='center'>
                                                        <table bgcolor='#ffffff' cellspacing='0' cellpadding='0' width='98%'>
                                                            <tr>
                                                                <td colspan='2' class='emailTitleStyle2'>Hello " + strName  + @",</td>
                                                            </tr>
                                                            <tr>
                                                                <td colspan='2' style='padding-bottom: 10px;' class='emailTitleStyle3'>Your Service Order has been completed.  Please click <a style='text-decoration: none;' href='http://execservicessf/Runtime/Runtime/Form/MareketingRequestTicketPublishedForm/?MarketingServiceTicketId=575610&amp;FarmId=1'>here</a> to view the completed deliverable(s).</td>
                                                            </tr>                                                                   
                                                        </table>
                                                    </td>
                                                </tr>
            ";
            return str;
        }
         public static string getFooter()
        {
            string str = @"
             </table>
             </td>
             </tr>
             </table>
             </td>
             </tr>
             </table>
             </td>
             </tr>
             </table>
            </body>
             ";
            return str;
        }
        public static string getQuery()
        {
            string str = @"select * from
            (select distinct
            abc.[MST_ID],
            abc.[MS_ID],
            abc.[MR_ID],
            abc.[Title],
            abc.[MR_Title],
            abc.[EmailAddress],
            abc.[Program],
            ISNULL(abc.[EPCampaignName],'-') 'EPCampaignName',
            abc.[BatchCount_recalc],
            abc.[SO_SubmittedDate_SubTimezone],
            abc.[DueDate_PST],
            abc.[DueDate_SubTimezone],
            abc.[ServiceTypeName],
            abc.[MarketingRequestName],
            abc.[SubsidiaryName],
            abc.areaname,
            abc.[TicketStepName],
            abc.TimeOfLastQualityCheck_IST,
             AttUrl.[FileUrl],
            AttUrl.[FileName],
            --case when abc.[MST_ID] in ('559568') then 3
            -- when abc.[MST_ID] in ('559570') then 6
            -- when abc.[MST_ID] in ('559572') then 9
            -- --when abc.[MST_ID] in ('559573') then 8
            --else 1 end as [Gap of Days],
            datediff(day,abc.TimeOfLastQualityCheck_IST,getdate()) as [Gap of Days],
            case when datediff(day,abc.TimeOfLastQualityCheck_IST,getdate()) in (2,5,8) then 'Yes'
            else 'No' end as [Email to be sent]
            ,AssignedUser.AssignedTo
            ,case 
						            when charindex('\',[MarketingRequestOwner]) > 0 
						            then Right([MarketingRequestOwner],len([MarketingRequestOwner])-(charindex('\',[MarketingRequestOwner])))+'@microsoft.com'
						            else ''
					            end as [MarketingRequestOwner]

            ,replace(abc.IRRN_Aliases, ';', '@microsoft.com;') as [IRRN_Aliases]
            ,(BuildUser.[BuilderEmail] +';'+PeerReviewer.peeremail) as [BuilderEmail]
            ,abc.TicketTriagedBy+'@microsoft.com' as SOP_Email
            ,('v-savyas@microsoft.com;smscodc@microsoft.com;' + FactoryLead.FactoryLeadEmail) as FactoryLeadEmail
           --    ,'v-himana@microsoft.com' as FL_Email
            --,case when datediff(day,abc.TimeOfLastQualityCheck_IST,getdate()) <=6  then 'Shaina.Mackie@dentsuaegis.com;Sudeep.Mishra@us.gt.com'
            --when datediff(day,abc.TimeOfLastQualityCheck_IST,getdate()) > 6 then 'Shaina.Mackie@dentsuaegis.com;erinwa@microsoft.com;shudson@microsoft.com;nclay@microsoft.com;Sudeep.Mishra@us.gt.com'
            --else '' end as [To be CCed]

            --,'v-sanaro@microsoft.com' as [MarketingRequestOwner],
            --'v-shecha@microsoft.com;v-deda@microsoft.com;' as [IRRN_Aliases],
            --'v-ankurb@microsoft.com' as BuilderEmail,
            --'v-ankurb@microsoft.com' as SOP_Email,
            --'v-ankurb@microsoft.com' as FL_Email
            --,'v-shecha@microsoft.com' as [To be CCed]



            ,cast(('http:'+'//'+'execservicessf/Runtime/Runtime/Form/MarketingRequestTktApprovalForm/?MarketingServiceTicketId='+cast(abc.[MST_ID] as nvarchar(max))+'&FarmId=1&SN='+cast(AssignedUser.K2SerialNum as nvarchar(max))) as  nvarchar(max)) as [Link]

             from 
            (
            SELECT 
                MST.Title,
	            MST.MarketingServiceTicketID as MST_ID
	            ,MS.MarketingServiceID as MS_ID
	            ,MR.MarketingRequestId as MR_ID
                ,MR.Title as MR_Title
	            ,'IRRN_Aliases' = IRRN.NotifiedUserAlias
            --=================================================	
            -- Modern VS Legacy requests
            --=================================================		
	
	            ,'ModernVSLegacyMarketing'=
							
				            case
					            --=========================================================
					            --Modern: Events (Certain, On24)- (this is old definition before 2016-07-01) 
					            --=========================================================
							            when (MST.TicketTag like N'%[#]Certain%' or EventTypeE.EventTypeName in ('Certain') or EventTypeA.EventTypeName in ('Certain') or EventTypeM.EventTypeName in ('Certain') or MST.TicketTag like N'%[#]On24%' or EventTypeE.EventTypeName in ('On24/Marketo') or EventTypeA.EventTypeName in ('On24/Marketo') or EventTypeM.EventTypeName in ('On24/Marketo'))
					
					            --=========================================================
					            --Modern: Events (Certain, On24)- (this is new definition after 2016-07-01) 
					            --=========================================================
							            OR (MSType.ServiceTypeName in ('Event Creation/Management','Update Existing Event','Event Reports') and MSTLatestUpdate.TicketLastUpdateDate>'2016-07-01 00:00:00.000')
							            OR MSType.ServiceTypeName in ('Event Creation and Demand Generation')

					            --=========================================================
					            --Modern: Smartlist
					            --=========================================================
							            OR MST.TicketTag like N'%[#]smartlist%' --old definiation before 2016-11-17
							            --select MarketingServiceID, UseMarketo from SuperSOEventCommunication where UseMarketo=1 --this defines that smart list was used within the Event Creation and Demand Generation SO but it is not needed because as it is under that SO, it is automatically counted as modern marketing

					            --=========================================================
					            --Modern:  1st party Gated Experience
					            --=========================================================
							            OR (((MST.TicketTag like N'%[#]CP%' OR MST.TicketTag like N'%[#]CLE%') AND MSType.ServiceTypeName in ('Website/Page Maintenance','Website/Page Production')
								            AND MP.ProgramName not in ('Dynamics Portal','Enterprise Portal','GigJam','MPN Portal','Public Sector','SMB Portal','Windows Commercial') AND Site.SiteName not in ('Licensing')) AND Site.SiteName not in ('Student'))
										
							            OR ContentMarketingTypes.ContentReasonName='Marketo form (MSC to build the form)'

					            --=========================================================
					            --Modern: 3rd party gated Experience
					            --=========================================================
							            OR MSType.ServiceTypeName in ('Syndicated Content') --old definition for Rapid Request
							            OR ContentMarketingTypes.ContentReasonName='Content Syndication in Marketo form, built by approved vendor'

					            --=========================================================
					            --Modern: Activate Marketo Campaign
					            --=========================================================
							            OR MSType.ServiceTypeName in ('Activate Marketo Campaign')

					            --=========================================================
					            --Modern: Lead Match/Upload - List upload for Demand Centre (consists of subprograms: Upload to Marketo, Upload from Marketo to CRMs)
					            --=========================================================
							            OR MST.TicketTag like N'%[#]MKTO%' OR MST.TicketTag like N'%[#]marketo2crm%'
							
							            ---------------LeadManagementSO---------------(question in the form: all leads are matched againts MSX-if any of the radio button is selected it means it is modern marketing)
							            OR LeadManagement.IsLeadMatchHandled is not null


							
					            then 'Modern Marketing'
					            else 'Legacy Marketing'
				            end


            --=================================================	
            -- Ticket level columns
            --=================================================	

	            ,MST.TicketTag
	            ,'Program'=	 case 
							            --====================================================================================
							            --====================================================================================
							            --Programs on SO level (all of those programs are also showed in Online Dashboard)
							            --====================================================================================
							            --====================================================================================
								            --=========================================================
								            --Partner Learning Program
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND (((((MST.Title like N'%[_]PLC[_]%' or MST.Title like N'%PLC[_]%' or MST.Title like N'%[_]PLC%' or MST.Title like N'%[#]PLC%' or MST.Title like N'%[ ]PLC%' ) AND MS.ServiceTypeID IN (10,15,22)) OR ((MST.Title like N'%[#]PLC%') AND MS.ServiceTypeID IN (31,32))) OR (MRSub.SubsidiaryName = 'GBS-NA' OR MRSub.SubsidiaryName = 'GBS-ASIA' OR MRSub.SubsidiaryName = 'GBS-JAPAN' OR MRSub.SubsidiaryName = 'GBS-LATAM' OR MRSub.SubsidiaryName = 'GBS-EMEA')) 
													            OR EventTypeE.EventTypeName in ('PLC','PLC-PA') OR EventTypeA.EventTypeName in ('PLC','PLC-PA') OR EventTypeM.EventTypeName in ('PLC','PLC-PA')))
											
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------								
													            OR (MP.ProgramName in ('Partner Learning Program'))
											            then 'Partner Learning Program'
																												
								            --=========================================================
								            --Microsoft Events
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.Title like N'%[#]Ignite%') 
													            OR MST.TicketTag like N'%[#]Ignite%'))													
											
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------								
													            OR (MP.ProgramName in ('Microsoft Events'))
											            then 'Microsoft Events'

								            --=========================================================
								            --MPN Portal
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((((MST.title like N'%[_]MPN[_]%' OR MST.title like N'%[#]MPN%' OR MST.Title like N'%[#]MPN'OR MST.title like N'%[_]MPN[ ]%') AND MS.ServiceTypeID IN (4,6,16,24,26,27,28,29)) OR (MRSub.SubsidiaryName = 'SMSG-MPN' OR MRSub.SubsidiaryName = 'MPN-portal')) 
													            OR Site.SiteName in ('Partner')))
																							
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------								
													            OR (MP.ProgramName in ('MPN Portal'))
											            then 'MPN Portal'
																																						
								            --=========================================================
								            --PARTNER RM
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015--exceptionally, Partner RM old taxonomy was extend till 2015-10-01 because Delhi team did not track it correctly via MR program level.
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.title like N'PRM[_]%' and  MRArea.AreaName = 'CorpSMSG' and MS.ServiceTypeID not IN (4,6,16,24,26,27,28,29)) OR (MRSub.SubsidiaryName = 'Partner-RM')))
																																			
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------								
													            OR (MP.ProgramName in ('PARTNER RM') OR MP.ProgramName in ('PARTNER AdHoc'))
											            then 'PARTNER RM'							

								            --=========================================================
								            --SMB ATM
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((((MST.Title like N'SMB[_]%') and MS.ServiceTypeID NOT IN (4,6,26,27,28)) OR (MRSub.SubsidiaryName = 'SMB-ATM'))
													            OR (MST.TicketTag like N'%[#]SMB%' and MS.ServiceTypeID NOT IN (4,6,26,27,28))))
																																		
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------								
													            OR (MP.ProgramName in ('SMB ATM'))
											            then 'SMB ATM'	

								            --=========================================================
								            --SMB Portal
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((((MST.Title like N'%[#]SMBver1.0%') and MS.ServiceTypeID IN (4,6,26,27,28)) OR (MRSub.SubsidiaryName = 'SMB-portal')) 
													            OR (Site.SiteName in ('SMB') AND RTHd.TemplateVersion like N'1.0%')))
																																													
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------								
													            OR (MP.ProgramName in ('SMB Portal'))
											            then 'SMB Portal'																			

								            --=========================================================
								            --DX RM
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015--exceptionally, DX RM old taxonomy was extend till 2015-10-01 because Delhi team did not track it correctly via MR program level.
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.title like N'%[#]DXRM%') 
													            OR MST.TicketTag like N'%[#]DXRM%' OR MRSub.SubsidiaryName= 'DX'))
											
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------							
													            OR (MP.ProgramName in ('DX RM') OR MST.TicketTag like N'%[#]DXRM%')
											            then 'DX RM'	

								            --=========================================================
								            --DX Learning Experience
								            --=========================================================
											            when MP.ProgramName in ('DX Learning Experience')
											            then 'DX Learning Experience'	

								            --=========================================================
								            --Cloud and Enterprise Demand Center
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.Title like N'%[#]ADC%') 
													            OR MST.TicketTag like N'%[#]ADC%' OR MRSub.SubsidiaryName = 'CnE-AzureDemandCenter'))
																							
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------							
													            OR (MP.ProgramName in ('Cloud and Enterprise Demand Center') OR MP.ProgramName in ('C+E - Advanced Analytics & IOT','C+E - Azure Platform','C+E - Business Applications (Dynamics)','C+E - Business Intelligence','C+E - Enterprise Mobility','C+E - Hybrid Cloud','C+E - Mission Critical Intelligence','C+E - Mobile Application Development','C+E - OneCommercial - Security'))
											            then 'Cloud and Enterprise Demand Center'	

								            --=========================================================
								            --MSDN/VSO
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.title like N'CVS[_]%' OR MRSub.SubsidiaryName = 'CnE') and MSType.ServiceTypeName<>'Content Localization'))
																							
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------							
													            OR (MP.ProgramName in ('MSDN/VSO'))
											            then 'MSDN/VSO'	
	
								            --=========================================================
								            --CRMOL
								            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.title like N'%[#]CRMOL%' OR MRSub.SubsidiaryName = 'MBS' OR MST.TicketTag like N'%[#]CRMOL%') AND MS.ServiceTypeID not IN (4,6,16,24,26,27,28,29,30)))
																																						
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------							
													            OR (MP.ProgramName in ('CRMOL'))
											            then 'CRMOL'

								            --=========================================================
								            --Office Demand Center
								            --=========================================================
											            when MP.ProgramName in ('Office Demand Center') OR MP.ProgramName in ('Office - Advanced Enterprise','Office - Collaboration','Office - Email Upgrade','Office - Mobile Productivity','Office - Modern Meetings')
											            then 'Office Demand Center'

								            --=========================================================
								            --Modern Workplace
								            --=========================================================
											            when MP.ProgramName in ('Modern Workplace')
											            then 'Modern Workplace'	

								            --=========================================================
								            --OEM AYB
								            --=========================================================
											            when MP.ProgramName in ('OEM AYB') OR MRSub.SubsidiaryName = 'OEM-AYB'
											            then 'OEM AYB'
											 
							            --====================================================================================
							            --====================================================================================
							            --NON SO Level Programs NOT showed in Online Dashboard
							            --====================================================================================
							            --====================================================================================
											
								            --=========================================================
								            --Modern Events
								            --=========================================================		
										            when MST.TicketTag like N'%[#]Certain%' or MST.TicketTag like N'%[#]On24%'
												            or EventTypeE.EventTypeName in ('Certain','On24/Marketo') or EventTypeA.EventTypeName in ('Certain','On24/Marketo') or EventTypeM.EventTypeName in ('Certain','On24/Marketo')
												            OR (MSType.ServiceTypeName in ('Event Creation/Management','Update Existing Event','Event Reports') and MSTLatestUpdate.TicketLastUpdateDate>'2016-07-01 00:00:00.000')
												            OR MSType.ServiceTypeName in ('Event Creation and Demand Generation')
										            then 'Modern Events'

						            --=====================================================================
						            --ALL else is considered as N/A (no program associated to that ticket)
						            --=====================================================================							
							            else 'N/A'
					             end		
					 
		            --===================
		            --SubProgram
		            --===================	
			            ,'SubProgram'=	 case	
				            --====================================================================================
				            --====================================================================================
				            --GEPs
				            --====================================================================================
				            --====================================================================================
					            --=========================================================
					            --Cloud and Enterprise Demand Center
					            --=========================================================
						            ----------------------------------
						            --Old definition before 10/01/2015
						            ----------------------------------
								            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.Title like N'%[#]ADC%') 
										            OR MST.TicketTag like N'%[#]ADC%' OR MRSub.SubsidiaryName = 'CnE-AzureDemandCenter'))
																							
						            ----------------------------------
						            --New definition after 10/01/2015
						            ----------------------------------							
										            OR (MP.ProgramName in ('Cloud and Enterprise Demand Center'))
								            then 'Cloud and Enterprise Demand Center'	
							
							            --=========================================================
							            --C+E - Advanced Analytics & IOT
							            --=========================================================
										            when MP.ProgramName in ('C+E - Advanced Analytics & IOT')
										            then 'GEP: C+E - Advanced Analytics & IOT'

							            --=========================================================
							            --C+E - Azure Platform
							            --=========================================================
										            when MP.ProgramName in ('C+E - Azure Platform')
										            then 'GEP: C+E - Azure Platform'
				
							            --=========================================================
							            --C+E - Business Applications (Dynamics)
							            --=========================================================
										            when MP.ProgramName in ('C+E - Business Applications (Dynamics)')
										            then 'GEP: C+E - Business Applications (Dynamics)'

							            --=========================================================
							            --C+E - Business Intelligence
							            --=========================================================
										            when MP.ProgramName in ('C+E - Business Intelligence')
										            then 'GEP: C+E - Business Intelligence'

							            --=========================================================
							            --C+E - Enterprise Mobility
							            --=========================================================
										            when MP.ProgramName in ('C+E - Enterprise Mobility')
										            then 'GEP: C+E - Enterprise Mobility'

							            --=========================================================
							            --C+E - Hybrid Cloud
							            --=========================================================
										            when MP.ProgramName in ('C+E - Hybrid Cloud')
										            then 'GEP: C+E - Hybrid Cloud'

							            --=========================================================
							            --C+E - Mission Critical Intelligence
							            --=========================================================
										            when MP.ProgramName in ('C+E - Mission Critical Intelligence')
										            then 'GEP: C+E - Mission Critical Intelligence'

							            --=========================================================
							            --C+E - Mobile Application Development
							            --=========================================================
										            when MP.ProgramName in ('C+E - Mobile Application Development')
										            then 'GEP: C+E - Mobile Application Development'
							
							            --=========================================================
							            --C+E - OneCommercial - Security
							            --=========================================================
										            when MP.ProgramName in ('C+E - OneCommercial - Security')
										            then 'GEP: C+E - OneCommercial - Security'

										
					            --=========================================================
					            --Office Demand Center
					            --=========================================================
								            when MP.ProgramName in ('Office Demand Center')
								            then 'Office Demand Center'

						            --=========================================================
						            --Office - Advanced Enterprise
						            --=========================================================
									            when MP.ProgramName in ('Office - Advanced Enterprise')
									            then 'GEP: Office - Advanced Enterprise'

						            --=========================================================
						            --Office - Collaboration
						            --=========================================================
									            when MP.ProgramName in ('Office - Collaboration')
									            then 'GEP: Office - Collaboration'

						            --=========================================================
						            --Office - Email Upgrade
						            --=========================================================
									            when MP.ProgramName in ('Office - Email Upgrade')
									            then 'GEP: Office - Email Upgrade'

						            --=========================================================
						            --Office - Mobile Productivity
						            --=========================================================
									            when MP.ProgramName in ('Office - Mobile Productivity')
									            then 'GEP: Office - Mobile Productivity'

						            --=========================================================
						            --Office - Modern Meetings
						            --=========================================================
									            when MP.ProgramName in ('Office - Modern Meetings')
									            then 'GEP: Office - Modern Meetings'

					            --=========================================================
					            --Windows
					            --=========================================================
								            when MP.ProgramName in ('Windows - Windows')
								            then 'GEP: Windows - Windows'


				            --====================================================================================
				            --====================================================================================
				            --Other Subprograms (programs with no corp entities and also programs that are superior and have some subprograms- they are also shown in Online Dashboard)
				            --====================================================================================
				            --====================================================================================
					
					            --=========================================================
					            --PARTNER RM
					            --=========================================================
									            ----------------------------------
									            --Old definition before 10/01/2015--exceptionally, Partner RM old taxonomy was extend till 2015-10-01 because Delhi team did not track it correctly via MR program level.
									            ----------------------------------
											            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.title like N'PRM[_]%' and  MRArea.AreaName = 'CorpSMSG' and MS.ServiceTypeID not IN (4,6,16,24,26,27,28,29)) OR (MRSub.SubsidiaryName = 'Partner-RM')))
																																			
									            ----------------------------------
									            --New definition after 10/01/2015
									            ----------------------------------								
													            OR (MP.ProgramName in ('PARTNER RM'))
											            then 'PARTNER RM'			

						            --=========================================================
						            --PARTNER AdHoc
						            --=========================================================
							            ----------------------------------
							            --Old definition before 10/01/2015--exceptionally, Partner AdHoc old taxonomy was extend till 2015-10-01 because Delhi team did not track it correctly via MR program level.
							            ----------------------------------
									            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.title like N'CPA[_]%' and  MRArea.AreaName = 'CorpSMSG' and MS.ServiceTypeID not IN (4,6,16,24,26,27,28,29)) OR (MRSub.SubsidiaryName = 'Partner-AdHoc')))
																																		
							            ----------------------------------
							            --New definition after 10/01/2015
							            ----------------------------------								
											            OR (MP.ProgramName in ('PARTNER AdHoc'))
									            then 'PARTNER AdHoc'
								
					            --=========================================================
					            --PARTNER Incentive
					            --=========================================================
								            when MP.ProgramName in ('PARTNER Incentive')
								            then 'PARTNER Incentive'

					            --=========================================================
					            --Corp AdHoc
					            --=========================================================
						            ----------------------------------
						            --Old definition before 10/01/2015
						            ----------------------------------
								            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND (((MST.title like N'CRP[_]%' OR MST.title like N'[#]CRP[_]%') and  MRArea.AreaName = 'CorpSMSG' and MS.ServiceTypeID not IN (4,6,16,24,26,27,28,29))
										            OR MST.TicketTag like N'%[#]CRP%'))
																																																							
						            ----------------------------------
						            --New definition after 10/01/2015
						            ----------------------------------								
										            OR (MP.ProgramName in ('Corp AdHoc'))
								            then 'Corp AdHoc'	

					            --=========================================================
					            --News Center
					            --=========================================================
						            -------------------------------------------------------------------
						            --Old definition before RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
								            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND (MST.Title like N'%[_]PRNEWS%' or MST.Title like N'%[#]PRNEWS%') AND MS.ServiceTypeID IN (4,6,16,24,26,27,28,29))
									
						            -------------------------------------------------------------------
						            --New definition after RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
										            OR MST.TicketTag like N'%[#]PRNEWS%'
								            then 'News Center'

					            --=========================================================
					            --Partner Marketing Center
					            --=========================================================
						            -------------------------------------------------------------------
						            --Old definition before RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
								            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.Title like N'%[_]PMC[_]%' OR MST.Title like N'%[#]PMC[_]%') AND MS.ServiceTypeID IN (4,6,16,24,26,27,28,29)))
									
						            -------------------------------------------------------------------
						            --New definition after RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
										            OR Site.SiteName in ('PMC')
								            then 'Partner Marketing Center'	

					            --=========================================================
					            --CashBack
					            --=========================================================
						            -------------------------------------------------------------------
						            --Old definition before RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
							            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND (MST.title like N'%[#]CashBack%' OR MST.title like N'%[_]CashBack%'))
									
						            -------------------------------------------------------------------
						            --New definition after RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
									            OR MST.TicketTag like N'%[#]CashBack%'
							            then 'CashBack'				
													
					            --=========================================================
					            --Investment Association Tool
					            --=========================================================
						            -------------------------------------------------------------------
						            --Old definition before RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
							            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND MST.title like N'%[#]IAT%')
									
						            -------------------------------------------------------------------
						            --New definition after RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
									            OR MST.TicketTag like N'%[#]IAT%' OR MSType.ServiceTypeName='IAT Association'
							            then 'Investment Association Tool'
										

					            --=========================================================
					            --Licensing Portal
					            --=========================================================
						            -------------------------------------------------------------------
						            --Old definition before RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
							            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.Title like '%!_Licensing%' escape '!' OR MST.Title like '%!_[#]Licensing%' escape '!') AND MS.ServiceTypeID IN (4,6,16,24,26,27,28,29,30)))
									
						            -------------------------------------------------------------------
						            --New definition after RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
									            OR Site.SiteName in ('Licensing')
							            then 'Licensing Portal'	
										
					            --=========================================================
					            --Student Portal
					            --=========================================================
						            -------------------------------------------------------------------
						            --Old definition before RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
							            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND ((MST.Title like '%!_Student%' escape '!' OR MST.Title like '%!_[#]Student%' escape '!') AND MS.ServiceTypeID IN (4,6,16,24,26,27,28,29,30)))
									
						            -------------------------------------------------------------------
						            --New definition after RTH was introduced (Report Tracking Header)
						            -------------------------------------------------------------------
									            OR Site.SiteName in ('Student')
							            then 'Student Portal'

					            --=========================================================
					            --Azure 
					            --=========================================================
								            when MP.ProgramName in ('Azure')
								            then 'Azure'

					            --=========================================================
					            --Digital Events
					            --=========================================================							
							            when MP.ProgramName in ('Digital Events')
							            then 'Digital Events'
									

				            else 'N/A'
			            end	
	
		            --===================
		            --SubProgram2
		            --===================
			            ,'SubProgram2'=	 case 
			
							            --=========================================================
							            --ME-Certain Urgent
							            --=========================================================		
								              when (MST.TicketTag like N'%[#]Certain%' or EventTypeE.EventTypeName in ('Certain') or EventTypeA.EventTypeName in ('Certain') or EventTypeM.EventTypeName in ('Certain')) and MST.TicketTag like N'%[#]eventhelp%' then 'ME-Certain Urgent'
							
							            --=========================================================
							            --ME-On24 Urgent
							            --=========================================================								
								              when (MST.TicketTag like N'%[#]On24%' or EventTypeE.EventTypeName in ('On24/Marketo') or EventTypeA.EventTypeName in ('On24/Marketo') or EventTypeM.EventTypeName in ('On24/Marketo')) and MST.TicketTag like N'%[#]eventhelp%' then 'ME-On24 Urgent'
							
							            --=========================================================
							            --ME-Certain T2
							            --=========================================================								  
								              when (MST.TicketTag like N'%[#]Certain%' or EventTypeE.EventTypeName in ('Certain') or EventTypeA.EventTypeName in ('Certain') or EventTypeM.EventTypeName in ('Certain')) and MST.TicketTag like N'%[#]t2%' then 'ME-Certain T2'
							
							            --=========================================================
							            --ME-Certain T3
							            --=========================================================		
								              when (MST.TicketTag like N'%[#]Certain%' or EventTypeE.EventTypeName in ('Certain') or EventTypeA.EventTypeName in ('Certain') or EventTypeM.EventTypeName in ('Certain')) then 'ME-Certain T3'
	
							            --=========================================================
							            --ME-On24 T2
							            --=========================================================	
								              when (MST.TicketTag like N'%[#]On24%' or EventTypeE.EventTypeName in ('On24/Marketo') or EventTypeA.EventTypeName in ('On24/Marketo') or EventTypeM.EventTypeName in ('On24/Marketo')) and MST.TicketTag like N'%[#]t2%' then 'ME-On24 T2'
							
							            --=========================================================
							            --ME-On24 T3
							            --=========================================================	
								              when (MST.TicketTag like N'%[#]On24%' or EventTypeE.EventTypeName in ('On24/Marketo') or EventTypeA.EventTypeName in ('On24/Marketo') or EventTypeM.EventTypeName in ('On24/Marketo')) then 'ME-On24 T3'
								  
							            --=========================================================
							            --WWE
							            --=========================================================	
								              when EventTypeE.EventTypeName in ('WWE') or EventTypeA.EventTypeName in ('WWE') or EventTypeM.EventTypeName in ('WWE') then 'WWE'

							            --=========================================================
							            --Upload to Marketo
							            --=========================================================	
								              when MST.TicketTag like N'%[#]MKTO%' then 'Upload to Marketo'

							            --=========================================================
							            --Upload from Marketo to CRMs
							            --=========================================================	
								              when MST.TicketTag like N'%[#]marketo2crm%' then 'Upload from Marketo to CRMs'

							            --=========================================================
							            --Marketo Email Deployment through ET     
							            --=========================================================	
								              when MST.TicketTag like N'%[#]email%' then 'Marketo Email Deployment through ET'
							
							            --=========================================================
							            --Responsive Design All Up
							            --=========================================================
								            -------------------------------------------------------------------
								            --Old definition before RTH was introduced (Report Tracking Header)
								            -------------------------------------------------------------------
									            when (MSTLatestUpdate.TicketLastUpdateDate<'2015-10-01 00:00:00.000' AND (MST.Title like N'%[#]RD%'))
									
								            -------------------------------------------------------------------
								            --New definition after RTH was introduced (Report Tracking Header)
								            -------------------------------------------------------------------
											            OR MST.TicketTag like N'%[#]RD%'
									            then 'Responsive Design All Up'	

								            else 'N/A'
							             end
	
		            ,'EventType'= case when MST.TicketTag like N'%[#]Certain%' or EventTypeE.EventTypeName in ('Certain') or EventTypeA.EventTypeName in ('Certain') or EventTypeM.EventTypeName in ('Certain') then 'Certain'
						               when MST.TicketTag like N'%[#]On24%' or EventTypeE.EventTypeName in ('On24/Marketo') or EventTypeA.EventTypeName in ('On24/Marketo') or EventTypeM.EventTypeName in ('On24/Marketo') then 'On24'
						               when EventTypeE.EventTypeName in ('WWE') or EventTypeA.EventTypeName in ('WWE') or EventTypeM.EventTypeName in ('WWE') then 'WWE'
						            else 'N/A'	 		
					              end
		            ,'EventId(Marketo)'=RTHe.EventId
						
		            --,RTHd.TemplateVersion
		            ,'TicketStatusName'= WSMST.WorkflowStatusName
		            --,'TicketStepName'= FWSMST.WorkflowStepName
		
		            --,'TicketLastUpdateDateKey'= replace(convert(varchar(20), convert(date, MSTLatestUpdate.TicketLastUpdateDate)),'-','')

		            --working minutes/hours per ticket without iteration (if build was rejected by QC or OA, WM put after that are not counted
			            --,MSTLatestUpdate.TicketWorkingMinutes
			            --,'TicketWorkingHours' = convert (float,MSTLatestUpdate.TicketWorkingMinutes)/60

		            --working minutes/hours per ticket including iteration (if build was rejected by QC or OA, WM put after that are inluded)
			            ,MSTLatestUpdate.TicketWorkingMinutesSUM
			            ,MSTLatestUpdate.TicketWorkingMinutesSUMProduction	
			            ,'TicketWorkingHoursSUM' = convert (float,MSTLatestUpdate.TicketWorkingMinutesSUM)/60
			            ,'TicketWorkingHoursSUM_Submit' = convert (float,MSTLatestUpdate.TicketWorkingMinutesSUM_Submit)/60
			            ,'TicketWorkingHoursSUM_Review' = convert (float,MSTLatestUpdate.TicketWorkingMinutesSUM_Review)/60
			            ,'TicketWorkingHoursSUM_Build' = convert (float,MSTLatestUpdate.TicketWorkingMinutesSUM_Build)/60
			            ,'TicketWorkingHoursSUM_QC' = convert (float,MSTLatestUpdate.TicketWorkingMinutesSUM_QC)/60
			            ,'TicketWorkingHoursSUM_Approve' = convert (float,MSTLatestUpdate.TicketWorkingMinutesSUM_Approve)/60
			            ,'TicketWorkingHoursSUM_Publish' = convert (float,MSTLatestUpdate.TicketWorkingMinutesSUM_Publish)/60
			
			            ,MSTLatestUpdate.Iteration --this column tells us how many times was most rejected build rejected by QC or OA. e.g. if Iteration is >=1 it means that at least one of the build on the ticket was once rejected
		            ,'BatchCount_recalc' = case
									            when MST.Batchcount is NULL or MST.Batchcount=0
									            then 1
									            else MST.Batchcount
								            end
		            ,AssignedUser.AssignedUser
		            ,'CurrentTicketAssignedUser'=CurrentTicketAssignedUser.CurrentTicketAssignedUser
		            ,'DueDate_PST' = MST.DueDate
		            ,'DueDate_SubTimezone' = (STORM.dbo.[svfConvertDateToTimeZone](MST.DueDate,MRSub.TimeZoneId,null))
	
	            --=================================================	
	            -- Resource section columns
	            --=================================================	
		            ,'CenterLocation'=case when GSGUserCenter.ServiceCenterName in ('Delhi Mid', 'Delhi Day','Delhi Night') then 'Delhi'
								            when GSGUserCenter.ServiceCenterName is null then 'N/A'
						            else GSGUserCenter.ServiceCenterName
						            end

	            --=================================================	
	            -- Ticket complexity and Priority columns
	            --=================================================	

		            ,'TicketComplexity' = MSTComplexity.ComplexityName
		            ,'TicketPiority' = MSTPriority.PriorityName 
	
	            --=================================================		
	            -- Date columns- all possible formats of date (left joined to the date when last task per ticket was completed)
	            --=================================================	
		            ,MSTLatestUpdate.TicketLastUpdateDate
		            ,DateFormats.[DateYYYYMMDDKey]
		            ,DateFormats.MonthName
		            ,DateFormats.FiscalMonth
		            ,DateFormats.FiscalYear
		            ,DateFormats.FiscalQuarter
		            ,DateFormats.FiscalYearMonthName
		            ,DateFormats.FiscalYearFiscalMonthForSortBy
		            --,'WeekCalendarNumber'=DATEPART(wk,MSTLatestUpdate.TicketLastUpdateDate)
		            --WeekFiscalNumber was just temporarily added for QBR for Ajay who was asked by Tomi
		            --,'WeekFiscalNumber'=case when datepart(wk,MSTLatestUpdate.TicketLastUpdateDate)>26 then datepart(wk,MSTLatestUpdate.TicketLastUpdateDate)-26
		            --	else datepart(wk,MSTLatestUpdate.TicketLastUpdateDate)+26 end

            --=================================================	
            -- Service Order Level columns
            --=================================================	

	            ,'Core/Digital' = case
						            when MSType.ServiceTypename in ('Banner Creation/Adaptation','Digital Tagging','Website/Page Production','Website/Page Maintenance','Image/Banner Replication/Resizing','Website Reports','Publish PR News Story','Redirect Or Retire Existing Website','Update Existing Website link', 'Update Existing Website Text', 'Resize Or Replicate Image') then 'Digital'
						            when MSType.ServiceTypename is NOT null then 'Core'
					             end
	            ,MSTypeC.ServiceTypeCategoryName
	            ,MSType.ServiceTypeName
	            --,'SOSubmitDate'= MS.SubmittedDate
	            --,MS.EPTacticId
	            --,MS.EPTrackingCode
	
	            ,'SOCreatedBy'= case 
						            when charindex('\',MS.CreatedBy) > 0 then Right(MS.CreatedBy,len(MS.CreatedBy)-(charindex('\',MS.CreatedBy)))
						            else ''
					            end 
	
	            ,'GAW_Triggered' =	case
							            when MS.IsAudienceApprovalRequired = 0 AND MS.IsBusinessApprovalRequired = 0 AND IsPrivacyApprovalRequired = 1 then 'Only Privacy'
							            when MS.IsAudienceApprovalRequired = 0 AND MS.IsBusinessApprovalRequired = 1 AND IsPrivacyApprovalRequired = 0 then 'Only Business'
							            when MS.IsAudienceApprovalRequired = 1 AND MS.IsBusinessApprovalRequired = 0 AND IsPrivacyApprovalRequired = 0 then 'Only Audience'
							            when MS.IsAudienceApprovalRequired = 0 AND MS.IsBusinessApprovalRequired = 1 AND IsPrivacyApprovalRequired = 1 then 'Privacy+Business'
							            when MS.IsAudienceApprovalRequired = 1 AND MS.IsBusinessApprovalRequired = 0 AND IsPrivacyApprovalRequired = 1 then 'Privacy+Audience'
							            when MS.IsAudienceApprovalRequired = 1 AND MS.IsBusinessApprovalRequired = 1 AND IsPrivacyApprovalRequired = 0 then 'Business+Audience'
							            when MS.IsAudienceApprovalRequired = 1 AND MS.IsBusinessApprovalRequired = 1 AND IsPrivacyApprovalRequired = 1 then 'Privacy+Business+Audience'
							            else 'No GAW'
						            end

	            ,'Email_SO_Type'=EmailType.EmailTypeName
	            ,EC.EventCategoryName
	            ,'TellUSHowYourContentWillBeGated'=ContentMarketingTypes.ContentReasonName

	            ,'GSGFactory'=
							
				            case
						            --=========================================================
						            --'Execution - Modern' factory: Events (Certain, On24)- (this is old definition before 2016-07-01) 
						            --=========================================================
							            when (MST.TicketTag like N'%[#]Certain%' or EventTypeE.EventTypeName in ('Certain') or EventTypeA.EventTypeName in ('Certain') or EventTypeM.EventTypeName in ('Certain') or MST.TicketTag like N'%[#]On24%' or EventTypeE.EventTypeName in ('On24/Marketo') or EventTypeA.EventTypeName in ('On24/Marketo') or EventTypeM.EventTypeName in ('On24/Marketo'))
					
						            --=========================================================
						            --'Execution - Modern' factory: Events (Certain, On24)- (this is new definition after 2016-07-01) 
						            --=========================================================
								            OR (MSType.ServiceTypeName in ('Event Creation/Management','Update Existing Event','Event Reports') and MSTLatestUpdate.TicketLastUpdateDate>'2016-07-01 00:00:00.000')
								            OR MSType.ServiceTypeName in ('Event Creation and Demand Generation')

						            --=========================================================
						            --'Execution - Modern' factory: Smartlist
						            --=========================================================
								            OR MST.TicketTag like N'%[#]smartlist%' --old definiation before 2016-11-17
								            --select MarketingServiceID, UseMarketo from SuperSOEventCommunication where UseMarketo=1 --this defines that smart list was used within the Event Creation and Demand Generation SO but it is not needed because as it is under that SO, it is automatically counted as modern marketing

						            --=========================================================
						            --'Execution - Modern' factory:  1st party Gated Experience
						            --=========================================================
								            OR (((MST.TicketTag like N'%[#]CP%' OR MST.TicketTag like N'%[#]CLE%') AND MSType.ServiceTypeName in ('Website/Page Maintenance','Website/Page Production')
									            AND MP.ProgramName not in ('Dynamics Portal','Enterprise Portal','GigJam','MPN Portal','Public Sector','SMB Portal','Windows Commercial') AND Site.SiteName not in ('Licensing')) AND Site.SiteName not in ('Student'))
										
								            OR ContentMarketingTypes.ContentReasonName='Marketo form (MSC to build the form)'

						            --=========================================================
						            --'Execution - Modern' factory: 3rd party gated Experience
						            --=========================================================
								            OR MSType.ServiceTypeName in ('Syndicated Content') --old definition for Rapid Request
								            OR ContentMarketingTypes.ContentReasonName='Content Syndication in Marketo form, built by approved vendor'

						            --=========================================================
						            --'Execution - Modern' factory: Activate Marketo Campaign
						            --=========================================================
								            OR MSType.ServiceTypeName in ('Activate Marketo Campaign')
								
							            then 'Execution - Modern' 
						
						            --=========================================================
						            --'Data - Modern' factory: Lead Match/Upload - List upload for Demand Centre (consists of subprograms: Upload to Marketo, Upload from Marketo to CRMs)
						            --=========================================================						
							            when MST.TicketTag like N'%[#]MKTO%' OR MST.TicketTag like N'%[#]marketo2crm%'
							
									            ---------------LeadManagementSO---------------(question in the form: all leads are matched againts MSX-if any of the radio button is selected it means it is modern marketing)
									            OR LeadManagement.IsLeadMatchHandled is not null
								
							
							            then 'Data - Modern' 
					            else GSGFactory.FactoryName
				            end
	
	            ,'Capability'=case	when MSType.ServiceTypeName in ('Activate Marketo Campaign') then 'Marketing Automation'
						            when MSType.ServiceTypeName in ('Ad hoc Initiative and IO Report') then 'Marketing Planning & Instrumentation'
						            when MSType.ServiceTypeName in ('Advanced Analysis') then 'Data & Response Management'
						            when MSType.ServiceTypeName in ('Banner Creation/Adaptation') then 'Digital'
						            when MSType.ServiceTypeName in ('Budget IO Creation/Management') then 'Marketing Planning & Instrumentation'
						            when MSType.ServiceTypeName in ('Campaign Counts') then 'Data & Response Management'
						            when MSType.ServiceTypeName in ('Campaign Reports') then 'Analytics & Reporting'
						            when MSType.ServiceTypeName in ('Cloud Prospecting Upload') then 'Content, PR & Media'
						            when MSType.ServiceTypeName in ('CloudDAM asset upload') then 'Content, PR & Media'
						            when MSType.ServiceTypeName in ('Content Localization') then 'Content, PR & Media'
						            when MSType.ServiceTypeName in ('Correct/Update RIO Reports') then ''
						            when MSType.ServiceTypeName in ('Customer/Partner List') then 'Data & Response Management'
						            when MSType.ServiceTypeName in ('Digital Tagging') then 'Digital'
						            when MSType.ServiceTypeName in ('Email Creation/Delivery') then 'Combination of Legacy Services'
						            when MSType.ServiceTypeName in ('Email Reports') then 'Analytics & Reporting'
						            when MSType.ServiceTypeName in ('Event Attendance Upload') then 'Events'
						            when MSType.ServiceTypeName in ('Event Creation/Management') then 'Events'
						            when MSType.ServiceTypeName in ('Event Reports') then 'Analytics & Reporting'
						            when MSType.ServiceTypeName in ('Generate Simple Report (Email, Event)') then 'Analytics & Reporting'
						            when MSType.ServiceTypeName in ('IAT Association') then 'Marketing Planning & Instrumentation'
						            when MSType.ServiceTypeName in ('Image/Banner Replication/Resizing') then 'Digital'
						            when MSType.ServiceTypeName in ('Initiative and IO Creation/Management') then 'Marketing Planning & Instrumentation'
						            when MSType.ServiceTypeName in ('Initiative Creation/Management') then 'Marketing Planning & Instrumentation'
						            when MSType.ServiceTypeName in ('Lead Match/Upload') then 'Marketing Automation'
						            when MSType.ServiceTypeName in ('Marketing Budget Transfer/Allocation') then 'Marketing Planning & Instrumentation'
						            when MSType.ServiceTypeName in ('Marketing Code Creation') then 'Marketing Planning & Instrumentation'
						            when MSType.ServiceTypeName in ('Marketing Forecast Upload') then 'Marketing Planning & Instrumentation'
						            when MSType.ServiceTypeName in ('Newsletter Subscription Upload') then 'Combination of Legacy Services'
						            when MSType.ServiceTypeName in ('Offline Evaluations Upload') then 'Events'
						            when MSType.ServiceTypeName in ('Online Profiling') then 'Combination of Legacy Services'
						            when MSType.ServiceTypeName in ('Publish PR News Story') then 'Digital'
						            when MSType.ServiceTypeName in ('Redirect Or Retire Existing Website') then 'Digital'
						            when MSType.ServiceTypeName in ('Resize Or Replicate Image') then 'Digital'
						            when MSType.ServiceTypeName in ('Update Existing Event') then 'Events'
						            when MSType.ServiceTypeName in ('Update Existing Website Link') then 'Digital'
						            when MSType.ServiceTypeName in ('Update Existing Website Text') then 'Digital'
						            when MSType.ServiceTypeName in ('Website Reports') then 'Analytics & Reporting'
						            when MSType.ServiceTypeName in ('Website/Page Maintenance') then 'Digital'
						            when MSType.ServiceTypeName in ('Website/Page Production') then 'Digital'
						            when MSType.ServiceTypeName in ('Wizard Reports') then 'Analytics & Reporting'
				            else 'Legacy services'
				            end
						

            --=================================================	
            -- Marketing Request Level columns
            --=================================================	

	            ,'MarketingRequestName' = MR.Title
	            ,MRSub.SubsidiaryName
	            ,MRArea.AreaName
                ,MRArea.EmailAddress
	            ,'MarketingRequestOwner' = MR.Owner
	            --,'MarketingRequestOwnerAlias' = case 
	            --									when charindex('\',MR.Owner) > 0 then Right(MR.Owner,len(MR.Owner)-(charindex('\',MR.Owner)))
	            --									else ''
	            --								end 
	            ,'EPCampaignRequired'= case
								            when MR.CampaignTypeSelectionId like '1' then 'Yes'
								            when MR.CampaignTypeSelectionId like '2' then 'No'
								            when MR.CampaignTypeSelectionId is null then 'N/A'
						               end
	            ,MR.EPDivisionName
	            ,MR.EPCampaignName
	            --,'EPJustification' = MR.NoCampaignNotes	--if marketer set that no EP campaign is needed for marketing request, this is the field where he/she needs to provide business justification
	            --,'MRStatus' = WSMR.WorkFlowStatusName
	            --,'MRCreatedBy' = MR.CreatedBy	--markterer/user who physically created marketing request via marketer portal
	            ,'MRCreatedByAlias' = case 
								            when charindex('\',MR.CreatedBy) > 0 then Right(MR.CreatedBy,len(MR.CreatedBy)-(charindex('\',MR.CreatedBy)))
								            else ''
						              end 
	


	            --=================================================	
	            -- MSD columns
	            --=================================================		
		            /*
			            These columns serve for MSD engagement report
				            MSD managed= all MRs where MSD was either on Owner line or he/she is creator of the MR
				            MSD on CC= all MRs where one of the MSD was part of IRRN line
		            */
		
		            --USE THIS section to get all list of MSD users
		            /*run these four selects once a month and update below aliases with those you get from the selects. We do this manually otherwise the automatic calculation takes too much time to refresh
			            --MSDs			
				            select distinct ''''+UserAlias+''''+',' from [RP_MSDUsersInfo] WITH (NOLOCK)
					            union
				            select ''''+ResourceAlias+''''+',' from [MSC].[!Archive_RP_WE_SDLs_AND_missing_CMPs] where Factory like 'SDL'
		
		            */
			            --,'SDL_CMP_Engagement'= case when 
			            --								--All MSDs who are owners or create MR
			            --									Right(MR.Owner,len(MR.Owner)-(charindex('\',MR.Owner))) in ('a-chyani',	'a-jotoh',	'a-juszuk',	'a-kabyeo',	'a-kalin',	'a-stleun',	'a-yanli',	'javier.peralta',	'Priscila.Vindas',	'Tatyana.Semenova',	'v-admir',	'v-ads',	'v-ahhall',	'v-aklong',	'v-alaria',	'v-aljorg',	'v-allech',	'v-anndun',	'v-anried',	'v-ansmo',	'v-antabu',	'v-arbejo',	'v-astra',	'v-auverm',	'v-aymuha',	'v-aytalu',	'v-azdast',	'v-bakas',	'v-behurs',	'v-beoi',	'v-beozer',	'v-betim',	'v-bewel',	'v-bltabb',	'v-brsaty',	'v-brunar',	'v-cabott',	'v-caflor',	'v-camiq',	'v-cara',	'v-carolc',	'v-cdeclo',	'v-cecart',	'v-cgiao',	'v-chesta',	'v-chkr',	'v-chmaja',	'v-chrstj',	'v-cillaw',	'v-cimarz',	'v-crmend',	'v-cyrco',	'v-dadasi',	'v-dafark',	'v-danmil',	'v-darodg',	'v-dermar',	'v-domaes',	'v-dorame',	'v-duzapa',	'v-elstim',	'v-erliao',	'v-esszab',	'v-evaso',	'v-evkopr',	'v-ezerda',	'v-fahami',	'v-gacas',	'v-gadach',	'v-gasaxe',	'v-gash',	'v-grchow',	'v-hanin',	'v-hapeli',	'v-hehops',	'v-hiclar',	'v-hiras',	'v-hs',	'v-ilkrav',	'v-ilyaf',	'v-imbend',	'v-irtsa',	'v-ismac',	'v-issala',	'v-iyakal',	'v-jagsin',	'v-janaan',	'v-jasaee',	'v-jawehb',	'v-jephie',	'v-jesaba',	'v-jesc',	'v-jessbe',	'v-jiyele',	'v-jizh',	'v-johuon',	'v-jomea',	'v-jowoud',	'v-juglav',	'v-juzuk',	'v-kabhat',	'v-kaboka',	'v-kadas',	'v-kahetz',	'v-kakrik',	'v-kankai',	'v-kathn',	'v-katmer',	'v-kecho',	'v-kefirg',	'v-kenjs',	'v-kewal',	'v-khajeb',	'v-knag',	'v-kriki',	'v-krmoe',	'v-krtopi',	'v-laazar',	'v-labell',	'v-labenh',	'Vladimir.rejlek',	'v-lagupt',	'v-lamsa',	'v-ldasca',	'v-leanmy',	'v-lehoog',	'v-leparo',	'v-liherz',	'v-limtin',	'v-lisans',	'v-loman',	'v-lorasa',	'v-luchlp',	'v-lukrin',	'v-lupape',	'v-luproc',	'v-lycarl',	'v-lykalo',	'v-maandu',	'v-maduff',	'v-maeto',	'v-mafres',	'v-malnik',	'v-maluri',	'v-mamaeg',	'v-mariea',	'v-marpie',	'v-marus',	'v-marwen',	'v-matmor',	'v-mavivo',	'v-mayag',	'v-miama',	'v-micyang',	'v-mihoub',	'v-mikoho',	'v-mimoug',	'v-misarm',	'v-mition',	'v-miusui',	'v-mmorac',	'v-momont',	'v-nalev',	'v-namilt',	'v-nasait',	'v-nistej',	'v-norubi',	'v-nyplan',	'v-olad',	'v-olcrok',	'v-oljaro',	'v-olkuus',	'v-olredm',	'v-onvasi',	'v-oselbo',	'v-oskhal',	'v-pankum',	'v-pataht',	'v-pathav',	'v-petroe',	'v-phchoo',	'v-piedma',	'v-priga',	'v-raashr',	'v-ramue',	'v-rapolo',	'v-rdwars',	'v-reblau',	'v-reburg',	'v-rekara',	'v-rfurla',	'v-robcas',	'v-rogli',	'v-rsfeir',	'v-ruraim',	'v-saccha',	'v-sadamm',	'v-sadelb',	'v-sakuyp',	'v-saprok',	'v-satsch',	'v-savyas',	'v-scbrei',	'v-sejeon',	'v-semata',	'v-serenk',	'v-shabba',	'v-silivp',	'v-skager',	'v-slawpo',	'v-slmelo',	'v-sobarb',	'v-sonab',	'v-sonija',	'v-spakl',	'v-ssarah',	'v-sshar',	'v-sthum',	'v-stmoo',	'v-sunkb',	'v-takurt',	'v-tasbou',	'v-temela',	'v-tetirk',	'v-thaise',	'v-thfig',	'v-thfige',	'v-thgrie',	'v-togerg',	'v-tomass',	'v-toshir',	'v-vamaal',	'v-vapham',	'v-vedia',	'v-vikore',	'v-vmuruz',	'v-wedeng',	'v-weihu',	'v-yakawa',	'v-ylain',	'v-yogek',	'v-zawaud',	'v-zhagak',	'v-zhweng',	'v-zubroz',	'v-zumaro')
			            --									OR Right(MR.CreatedBy,len(MR.CreatedBy)-(charindex('\',MR.CreatedBy))) in ('a-chyani',	'a-jotoh',	'a-juszuk',	'a-kabyeo',	'a-kalin',	'a-stleun',	'a-yanli',	'javier.peralta',	'Priscila.Vindas',	'Tatyana.Semenova',	'v-admir',	'v-ads',	'v-ahhall',	'v-aklong',	'v-alaria',	'v-aljorg',	'v-allech',	'v-anndun',	'v-anried',	'v-ansmo',	'v-antabu',	'v-arbejo',	'v-astra',	'v-auverm',	'v-aymuha',	'v-aytalu',	'v-azdast',	'v-bakas',	'v-behurs',	'v-beoi',	'v-beozer',	'v-betim',	'v-bewel',	'v-bltabb',	'v-brsaty',	'v-brunar',	'v-cabott',	'v-caflor',	'v-camiq',	'v-cara',	'v-carolc',	'v-cdeclo',	'v-cecart',	'v-cgiao',	'v-chesta',	'v-chkr',	'v-chmaja',	'v-chrstj',	'v-cillaw',	'v-cimarz',	'v-crmend',	'v-cyrco',	'v-dadasi',	'v-dafark',	'v-danmil',	'v-darodg',	'v-dermar',	'v-domaes',	'v-dorame',	'v-duzapa',	'v-elstim',	'v-erliao',	'v-esszab',	'v-evaso',	'v-evkopr',	'v-ezerda',	'v-fahami',	'v-gacas',	'v-gadach',	'v-gasaxe',	'v-gash',	'v-grchow',	'v-hanin',	'v-hapeli',	'v-hehops',	'v-hiclar',	'v-hiras',	'v-hs',	'v-ilkrav',	'v-ilyaf',	'v-imbend',	'v-irtsa',	'v-ismac',	'v-issala',	'v-iyakal',	'v-jagsin',	'v-janaan',	'v-jasaee',	'v-jawehb',	'v-jephie',	'v-jesaba',	'v-jesc',	'v-jessbe',	'v-jiyele',	'v-jizh',	'v-johuon',	'v-jomea',	'v-jowoud',	'v-juglav',	'v-juzuk',	'v-kabhat',	'v-kaboka',	'v-kadas',	'v-kahetz',	'v-kakrik',	'v-kankai',	'v-kathn',	'v-katmer',	'v-kecho',	'v-kefirg',	'v-kenjs',	'v-kewal',	'v-khajeb',	'v-knag',	'v-kriki',	'v-krmoe',	'v-krtopi',	'v-laazar',	'v-labell',	'v-labenh',	'Vladimir.rejlek',	'v-lagupt',	'v-lamsa',	'v-ldasca',	'v-leanmy',	'v-lehoog',	'v-leparo',	'v-liherz',	'v-limtin',	'v-lisans',	'v-loman',	'v-lorasa',	'v-luchlp',	'v-lukrin',	'v-lupape',	'v-luproc',	'v-lycarl',	'v-lykalo',	'v-maandu',	'v-maduff',	'v-maeto',	'v-mafres',	'v-malnik',	'v-maluri',	'v-mamaeg',	'v-mariea',	'v-marpie',	'v-marus',	'v-marwen',	'v-matmor',	'v-mavivo',	'v-mayag',	'v-miama',	'v-micyang',	'v-mihoub',	'v-mikoho',	'v-mimoug',	'v-misarm',	'v-mition',	'v-miusui',	'v-mmorac',	'v-momont',	'v-nalev',	'v-namilt',	'v-nasait',	'v-nistej',	'v-norubi',	'v-nyplan',	'v-olad',	'v-olcrok',	'v-oljaro',	'v-olkuus',	'v-olredm',	'v-onvasi',	'v-oselbo',	'v-oskhal',	'v-pankum',	'v-pataht',	'v-pathav',	'v-petroe',	'v-phchoo',	'v-piedma',	'v-priga',	'v-raashr',	'v-ramue',	'v-rapolo',	'v-rdwars',	'v-reblau',	'v-reburg',	'v-rekara',	'v-rfurla',	'v-robcas',	'v-rogli',	'v-rsfeir',	'v-ruraim',	'v-saccha',	'v-sadamm',	'v-sadelb',	'v-sakuyp',	'v-saprok',	'v-satsch',	'v-savyas',	'v-scbrei',	'v-sejeon',	'v-semata',	'v-serenk',	'v-shabba',	'v-silivp',	'v-skager',	'v-slawpo',	'v-slmelo',	'v-sobarb',	'v-sonab',	'v-sonija',	'v-spakl',	'v-ssarah',	'v-sshar',	'v-sthum',	'v-stmoo',	'v-sunkb',	'v-takurt',	'v-tasbou',	'v-temela',	'v-tetirk',	'v-thaise',	'v-thfig',	'v-thfige',	'v-thgrie',	'v-togerg',	'v-tomass',	'v-toshir',	'v-vamaal',	'v-vapham',	'v-vedia',	'v-vikore',	'v-vmuruz',	'v-wedeng',	'v-weihu',	'v-yakawa',	'v-ylain',	'v-yogek',	'v-zawaud',	'v-zhagak',	'v-zhweng',	'v-zubroz',	'v-zumaro')
			            --								then 'MSD managed'
										
			            --							when 
			            --								--All MSDs who created ticket (triaged ticket)
			            --									Right(MST.CreatedBy,len(MST.CreatedBy)-(charindex('\',MST.CreatedBy))) in ('a-chyani',	'a-jotoh',	'a-juszuk',	'a-kabyeo',	'a-kalin',	'a-stleun',	'a-yanli',	'javier.peralta',	'Priscila.Vindas',	'Tatyana.Semenova',	'v-admir',	'v-ads',	'v-ahhall',	'v-aklong',	'v-alaria',	'v-aljorg',	'v-allech',	'v-anndun',	'v-anried',	'v-ansmo',	'v-antabu',	'v-arbejo',	'v-astra',	'v-auverm',	'v-aymuha',	'v-aytalu',	'v-azdast',	'v-bakas',	'v-behurs',	'v-beoi',	'v-beozer',	'v-betim',	'v-bewel',	'v-bltabb',	'v-brsaty',	'v-brunar',	'v-cabott',	'v-caflor',	'v-camiq',	'v-cara',	'v-carolc',	'v-cdeclo',	'v-cecart',	'v-cgiao',	'v-chesta',	'v-chkr',	'v-chmaja',	'v-chrstj',	'v-cillaw',	'v-cimarz',	'v-crmend',	'v-cyrco',	'v-dadasi',	'v-dafark',	'v-danmil',	'v-darodg',	'v-dermar',	'v-domaes',	'v-dorame',	'v-duzapa',	'v-elstim',	'v-erliao',	'v-esszab',	'v-evaso',	'v-evkopr',	'v-ezerda',	'v-fahami',	'v-gacas',	'v-gadach',	'v-gasaxe',	'v-gash',	'v-grchow',	'v-hanin',	'v-hapeli',	'v-hehops',	'v-hiclar',	'v-hiras',	'v-hs',	'v-ilkrav',	'v-ilyaf',	'v-imbend',	'v-irtsa',	'v-ismac',	'v-issala',	'v-iyakal',	'v-jagsin',	'v-janaan',	'v-jasaee',	'v-jawehb',	'v-jephie',	'v-jesaba',	'v-jesc',	'v-jessbe',	'v-jiyele',	'v-jizh',	'v-johuon',	'v-jomea',	'v-jowoud',	'v-juglav',	'v-juzuk',	'v-kabhat',	'v-kaboka',	'v-kadas',	'v-kahetz',	'v-kakrik',	'v-kankai',	'v-kathn',	'v-katmer',	'v-kecho',	'v-kefirg',	'v-kenjs',	'v-kewal',	'v-khajeb',	'v-knag',	'v-kriki',	'v-krmoe',	'v-krtopi',	'v-laazar',	'v-labell',	'v-labenh',	'Vladimir.rejlek',	'v-lagupt',	'v-lamsa',	'v-ldasca',	'v-leanmy',	'v-lehoog',	'v-leparo',	'v-liherz',	'v-limtin',	'v-lisans',	'v-loman',	'v-lorasa',	'v-luchlp',	'v-lukrin',	'v-lupape',	'v-luproc',	'v-lycarl',	'v-lykalo',	'v-maandu',	'v-maduff',	'v-maeto',	'v-mafres',	'v-malnik',	'v-maluri',	'v-mamaeg',	'v-mariea',	'v-marpie',	'v-marus',	'v-marwen',	'v-matmor',	'v-mavivo',	'v-mayag',	'v-miama',	'v-micyang',	'v-mihoub',	'v-mikoho',	'v-mimoug',	'v-misarm',	'v-mition',	'v-miusui',	'v-mmorac',	'v-momont',	'v-nalev',	'v-namilt',	'v-nasait',	'v-nistej',	'v-norubi',	'v-nyplan',	'v-olad',	'v-olcrok',	'v-oljaro',	'v-olkuus',	'v-olredm',	'v-onvasi',	'v-oselbo',	'v-oskhal',	'v-pankum',	'v-pataht',	'v-pathav',	'v-petroe',	'v-phchoo',	'v-piedma',	'v-priga',	'v-raashr',	'v-ramue',	'v-rapolo',	'v-rdwars',	'v-reblau',	'v-reburg',	'v-rekara',	'v-rfurla',	'v-robcas',	'v-rogli',	'v-rsfeir',	'v-ruraim',	'v-saccha',	'v-sadamm',	'v-sadelb',	'v-sakuyp',	'v-saprok',	'v-satsch',	'v-savyas',	'v-scbrei',	'v-sejeon',	'v-semata',	'v-serenk',	'v-shabba',	'v-silivp',	'v-skager',	'v-slawpo',	'v-slmelo',	'v-sobarb',	'v-sonab',	'v-sonija',	'v-spakl',	'v-ssarah',	'v-sshar',	'v-sthum',	'v-stmoo',	'v-sunkb',	'v-takurt',	'v-tasbou',	'v-temela',	'v-tetirk',	'v-thaise',	'v-thfig',	'v-thfige',	'v-thgrie',	'v-togerg',	'v-tomass',	'v-toshir',	'v-vamaal',	'v-vapham',	'v-vedia',	'v-vikore',	'v-vmuruz',	'v-wedeng',	'v-weihu',	'v-yakawa',	'v-ylain',	'v-yogek',	'v-zawaud',	'v-zhagak',	'v-zhweng',	'v-zubroz',	'v-zumaro')
			            --								then 'MSD triaged'
			            --							when IRRN_SDL2.IsNotifiedUserSDL=1 then 'MSD on CC'
			            --						    else 'Marketer/AL'
			            --					  end
			            ,'TicketStepName'= FWSMST.WorkflowStepName
			            ,'TicketTriagedBy'= case 
									            when charindex('\',MST.CreatedBy) > 0 then Right(MST.CreatedBy,len(MST.CreatedBy)-(charindex('\',MST.CreatedBy)))
									            else ''
								            end 
			  	
            --=============================================================
            -- LTPS (Lead Time Per Service) columns
            --=============================================================	     
		            /*
			            1. If there is no NotStartBeforeDate set up on ticket, then we take SO_SubmittedDate (service order) as the first date for measuring production LTPS 
				            and last quality check per ticket as the last date for that calculation
				
			            2. If NotStartBeforeDate is set up on ticket, then Production lead time is summary of below periods:
					            a. Time from SO submission date till the time when SOP triage the ticket (MSTT.StatusID = 3 /*Complete*/ and MSTT.WorkflowStepId= 1 /*Triage*/)
						            +
					            b. Time from NotStartBeforeDate till last quality check per ticket

			            Notes: 1. For some tickets it happens that NotStartBeforeDate is lesser than SO_SubmittedDate. In that case we take SO_SubmittedDate as the first 
				                date for LTPS calculation. That happens because there is NO timestamp for NotStartBeforeDate in UI, therefore, if user fill it in one timezone,
					            it will put default time set by GSG and store it in PST. There is GSG enhancement to add timestamp to that date to comprehend different timezones
					            2. For some tickets it happened that TimeOfLastQualityCheck was lesser than NotStartBeforeDate. This happened because there was no restriction in GSG
					            to process build sooner than NSBD. From 2014 it is not possible to process build sooner than NSBD
		            */
		
		            ,'LTPS'= case 
					
					            --Issue corrections
						            --Removing couple of MRs from LTPS for Cairo for datauploads beacuase these were very long term reccurent tickets when NSBD was not set up
							            when MR.MarketingRequestId in (5763,	5759,	5757,	5753,	5765,	5758,	5761,	9147,	10170,	10181,	10191,	10192,	10215,	10218,	22532,	27122,	27124,	27128,	27137,	27138,	22532,	29874,	38870,2278,	2373,	2493,	3106,	3552,	6968,	11206,	12929,	12941,	15382,	19552,	22345,	23379,	30054,	30237,	37415,	40185,	16282)
								            then null
					            --End of Issue correction
					
					
					            when (LTPS.TicketNotStartBeforeDate_PST is null or LTPS.TicketNotStartBeforeDate_PST < LTPS.SOPAssignedDate_PST or LTPS.TicketNotStartBeforeDate_PST > LTPS.TimeOfLastQualityCheck_PST) 
						            and datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))))*2) < 3 --tickets with no DNSBD and Triage after SO submission
						            AND MS.ServiceTypeID not IN (21,22,23,24,25) --removing reports
						            then datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,MRSub.TimeZoneId,null))))*2) --datediff for week is used to remove weekends from LTPS
						  
					            when (LTPS.TicketNotStartBeforeDate_PST is null or LTPS.TicketNotStartBeforeDate_PST < LTPS.SOPAssignedDate_PST or LTPS.TicketNotStartBeforeDate_PST > LTPS.TimeOfLastQualityCheck_PST) 
						            and datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))))*2) > 2 --tickets with no DNSBD and Triage too late due to additional ticket
						            AND MS.ServiceTypeID not IN (21,22,23,24,25) --removing reports
						            then datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,MRSub.TimeZoneId,null))))*2) --datediff for week is used to remove weekends from LTPS
						  
					            when (LTPS.TicketNotStartBeforeDate_PST is not null AND LTPS.TicketNotStartBeforeDate_PST > LTPS.SOPAssignedDate_PST AND LTPS.TicketNotStartBeforeDate_PST < LTPS.TimeOfLastQualityCheck_PST) 
						            and datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))))*2) < 3 --tickets with DNSBD and Triage after SO submission
						            AND MS.ServiceTypeID not IN (21,22,23,24,25) --removing reports
						            then (datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))))*2)) --datediff for week is used to remove weekends from LTPS
								            +
							            (datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.TicketNotStartBeforeDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.TicketNotStartBeforeDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,MRSub.TimeZoneId,null))))*2)) --datediff for week is used to remove weekends from LTPS
					
					            when (LTPS.TicketNotStartBeforeDate_PST is not null AND LTPS.TicketNotStartBeforeDate_PST > LTPS.SOPAssignedDate_PST AND LTPS.TicketNotStartBeforeDate_PST < LTPS.TimeOfLastQualityCheck_PST) 
						            and datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastTriage_PST,MRSub.TimeZoneId,null))))*2) > 2 --tickets with DNSBD and Triage too late due to additional ticket 
						            AND MS.ServiceTypeID not IN (21,22,23,24,25) --removing reports
						            then (datediff (day,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.TicketNotStartBeforeDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,MRSub.TimeZoneId,null))) - ((datediff (week,(STORM.dbo.[svfConvertDateToTimeZone](LTPS.TicketNotStartBeforeDate_PST,MRSub.TimeZoneId,null)), (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,MRSub.TimeZoneId,null))))*2)) --datediff for week is used to remove weekends from LTPS

					            end
							                      
                --=================================================	
	            -- LTPS additional columns
	            --=================================================	
	                  
		            --,LTPS.SO_SubmittedDate_PST
		            --,LTPS.TimeOfLastTriage_PST
		            --,LTPS.TimeOfLastAssignment_PST
		            --,LTPS.TimeOfLastBuild_PST
		            --,LTPS.TimeOfLastQualityCheck_PST
       
		            ,'SO_SubmittedDate_SubTimezone' = (STORM.dbo.[svfConvertDateToTimeZone](LTPS.SO_SubmittedDate_PST,MRSub.TimeZoneId,null))
		            ,'SOPAssignedDate_SubTimezone' = (STORM.dbo.[svfConvertDateToTimeZone](LTPS.SOPAssignedDate_PST,MRSub.TimeZoneId,null))
		            ,'TimeofLastTriage_SubTimezone' = (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeofLastTriage_PST,MRSub.TimeZoneId,null))
		            ,'TimeofLastAssignment_SubTimezone' = (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeofLastAssignment_PST,MRSub.TimeZoneId,null))
		            ,'TimeofLastBuild_SubTimezone' = (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeofLastBuild_PST,MRSub.TimeZoneId,null))
		            ,'TimeOfLastQualityCheck_IST' = (STORM.dbo.[svfConvertDateToTimeZone](LTPS.TimeOfLastQualityCheck_PST,'India Standard Time',null))
       
		            ,'TicketCreatedDate_SubTimezone' = (STORM.dbo.[svfConvertDateToTimeZone](MST.CreatedDate,MRSub.TimeZoneId,null))
		            ,'TicketNotStartBeforeDate_SubTimezone'= (STORM.dbo.[svfConvertDateToTimeZone](MST.NotStartBeforeDate ,MRSub.TimeZoneId,null)) 
		   
		            --,'TicketCreatedDate_PST' = MST.CreatedDate 
		            --,'TicketNotStartBeforeDate_PST'= MST.NotStartBeforeDate     
		   

		            ,SO_submitted_before_or_after_2pm= case 
												            when convert(time,(STORM.dbo.[svfConvertDateToTimeZone](SO_SubmittedDate_PST,MRSub.TimeZoneId,null))) > '14:00:00' then 'SO submitted after 2pm' 
												            when convert(time,(STORM.dbo.[svfConvertDateToTimeZone](SO_SubmittedDate_PST,MRSub.TimeZoneId,null))) < '14:00:00' then 'SO submitted before 2pm' 
											            end 

		  
		            --,LTPS.TimeZoneId
 
             FROM
	            MarketingServiceTicket as MST WITH (NOLOCK)

            --=================================================	
            -- Service Order Level connection
            --=================================================	

	            left join MarketingService as MS WITH (NOLOCK)
		            on MST.MarketingServiceID = MS.MarketingServiceID
	            left join MarketingServiceType MSType WITH (NOLOCK)
		            on MS.ServiceTypeID = MSType.ServiceTypeID
	            left join MarketingServiceTypeCategory MSTypeC WITH (NOLOCK)
		            on MSType.ServiceTypeCategoryID = MSTypeC.ServiceTypeCategoryID
	            left join Factory GSGFactory WITH (NOLOCK)
		            on GSGFactory.FactoryID=MSType.factoryID

            ---------------SpecificServices---------------
	            left join MarketingEmailService MSEmail WITH (NOLOCK)
		            on MS.MarketingServiceId = MSEmail.MarketingServiceId
	            left join EmailType WITH (NOLOCK)
		            on MSEmail.EmailTypeId = EmailType.EmailTypeId
            ----------EventCategoryName-----------------------------------
	            left join EventCategory EC WITH (NOLOCK)
		            on EC.EventCategoryId = MS.EventTypeId

            ---------------SpecificServices---------------

            ---------------Content Marketing (tell us how your content will be gated- Marketo form (MSC to build the form), Content Syndication in Marketo form, built by approved vendor))------------
	            left join ContentMarketingInitialInformation ContentMarketing
		            on MS.MarketingServiceID=ContentMarketing.MarketingServiceId
	            left join ContentMarketingContentReason ContentMarketingTypes
		            on ContentMarketing.GatedContentReasonId=ContentMarketingTypes.ContentReasonId

            ---------------LeadManagementSO---------------(question in the form: all leads are matched againts MSX-if any of the radio button is selected it means it is modern marketing)
	            left join MarketingDataUploadService LeadManagement
		            on MS.MarketingServiceID=LeadManagement.MarketingServiceId


            --=================================================	
            -- Marketing Request level connection
            --=================================================	

	            left join MarketingRequest as MR WITH (NOLOCK)
		            on MS.MarketingRequestId = MR.MarketingRequestId
	            left join MarketingSubsidiary as MRSub WITH (NOLOCK)
		            on MR.SubsidiaryID = MRSub.SubsidiaryID
	            left join MarketingArea as MRArea WITH (NOLOCK)
		            on MRSub.AreaId=MRArea.AreaID
	            left join WorkFlowStatus as WSMR WITH (NOLOCK)
		            on MR.StatusId=WSMR.WorkFlowStatusId
	            left join MarketingServiceCenter MRCenter WITH (NOLOCK)
		            on MRSub.Servicecenterid = MRCenter.ServiceCenterId

		            --=================================================	
		            -- IRRN connection (Individuals To Receive Request Notifications)-transposed to column
		            --=================================================	
			            left join Storm_reporting.[MSC].[RP_v_IRRN_TransposedToColumn] IRRN WITH (NOLOCK)
				            on MR.MarketingRequestId = IRRN.MarketingRequestId
				
		            --=================================================	
		            -- Program connection
		            --=================================================	
			            left join [MarketingProgram] MP WITH (NOLOCK)
				            on MS.ProgramId = MP.ProgramId
				

	            --=================================================	
	            -- SDL columns for IRRN
	            --=================================================	
		            /*
			            These columns are here for MSD engagement report. Below select is finding out whether one of the MSDs was on IRRN line.
			            --run below two selects once a month and update below aliases with those you get from the selects. We do this manually otherwise the automatic calculation takes too much time to refresh
				            select distinct ''''+UserAlias+''''+',' from [RP_MSDUsersInfo] WITH (NOLOCK)
					            union
				            select ''''+ResourceAlias+''''+',' from [MSC].[!Archive_RP_WE_SDLs_AND_missing_CMPs] where Factory like 'SDL'
		            */
		            --left join (select MarketingRequestId,
		            --		  'IsNotifiedUserSDL'= max(case when Right(NotifiedUserAlias,len(NotifiedUserAlias)-(charindex('\',NotifiedUserAlias))) in ('a-chyani',	'a-jotoh',	'a-juszuk',	'a-kabyeo',	'a-kalin',	'a-stleun',	'a-yanli',	'javier.peralta',	'Priscila.Vindas',	'Tatyana.Semenova',	'v-admir',	'v-ads',	'v-ahhall',	'v-aklong',	'v-alaria',	'v-aljorg',	'v-allech',	'v-anndun',	'v-anried',	'v-ansmo',	'v-antabu',	'v-arbejo',	'v-astra',	'v-auverm',	'v-aymuha',	'v-aytalu',	'v-azdast',	'v-bakas',	'v-behurs',	'v-beoi',	'v-beozer',	'v-betim',	'v-bewel',	'v-bltabb',	'v-brsaty',	'v-brunar',	'v-cabott',	'v-caflor',	'v-camiq',	'v-cara',	'v-carolc',	'v-cdeclo',	'v-cecart',	'v-cgiao',	'v-chesta',	'v-chkr',	'v-chmaja',	'v-chrstj',	'v-cillaw',	'v-cimarz',	'v-crmend',	'v-cyrco',	'v-dadasi',	'v-dafark',	'v-danmil',	'v-darodg',	'v-dermar',	'v-domaes',	'v-dorame',	'v-duzapa',	'v-elstim',	'v-erliao',	'v-esszab',	'v-evaso',	'v-evkopr',	'v-ezerda',	'v-fahami',	'v-gacas',	'v-gadach',	'v-gasaxe',	'v-gash',	'v-grchow',	'v-hanin',	'v-hapeli',	'v-hehops',	'v-hiclar',	'v-hiras',	'v-hs',	'v-ilkrav',	'v-ilyaf',	'v-imbend',	'v-irtsa',	'v-ismac',	'v-issala',	'v-iyakal',	'v-jagsin',	'v-janaan',	'v-jasaee',	'v-jawehb',	'v-jephie',	'v-jesaba',	'v-jesc',	'v-jessbe',	'v-jiyele',	'v-jizh',	'v-johuon',	'v-jomea',	'v-jowoud',	'v-juglav',	'v-juzuk',	'v-kabhat',	'v-kaboka',	'v-kadas',	'v-kahetz',	'v-kakrik',	'v-kankai',	'v-kathn',	'v-katmer',	'v-kecho',	'v-kefirg',	'v-kenjs',	'v-kewal',	'v-khajeb',	'v-knag',	'v-kriki',	'v-krmoe',	'v-krtopi',	'v-laazar',	'v-labell',	'v-labenh',	'Vladimir.rejlek',	'v-lagupt',	'v-lamsa',	'v-ldasca',	'v-leanmy',	'v-lehoog',	'v-leparo',	'v-liherz',	'v-limtin',	'v-lisans',	'v-loman',	'v-lorasa',	'v-luchlp',	'v-lukrin',	'v-lupape',	'v-luproc',	'v-lycarl',	'v-lykalo',	'v-maandu',	'v-maduff',	'v-maeto',	'v-mafres',	'v-malnik',	'v-maluri',	'v-mamaeg',	'v-mariea',	'v-marpie',	'v-marus',	'v-marwen',	'v-matmor',	'v-mavivo',	'v-mayag',	'v-miama',	'v-micyang',	'v-mihoub',	'v-mikoho',	'v-mimoug',	'v-misarm',	'v-mition',	'v-miusui',	'v-mmorac',	'v-momont',	'v-nalev',	'v-namilt',	'v-nasait',	'v-nistej',	'v-norubi',	'v-nyplan',	'v-olad',	'v-olcrok',	'v-oljaro',	'v-olkuus',	'v-olredm',	'v-onvasi',	'v-oselbo',	'v-oskhal',	'v-pankum',	'v-pataht',	'v-pathav',	'v-petroe',	'v-phchoo',	'v-piedma',	'v-priga',	'v-raashr',	'v-ramue',	'v-rapolo',	'v-rdwars',	'v-reblau',	'v-reburg',	'v-rekara',	'v-rfurla',	'v-robcas',	'v-rogli',	'v-rsfeir',	'v-ruraim',	'v-saccha',	'v-sadamm',	'v-sadelb',	'v-sakuyp',	'v-saprok',	'v-satsch',	'v-savyas',	'v-scbrei',	'v-sejeon',	'v-semata',	'v-serenk',	'v-shabba',	'v-silivp',	'v-skager',	'v-slawpo',	'v-slmelo',	'v-sobarb',	'v-sonab',	'v-sonija',	'v-spakl',	'v-ssarah',	'v-sshar',	'v-sthum',	'v-stmoo',	'v-sunkb',	'v-takurt',	'v-tasbou',	'v-temela',	'v-tetirk',	'v-thaise',	'v-thfig',	'v-thfige',	'v-thgrie',	'v-togerg',	'v-tomass',	'v-toshir',	'v-vamaal',	'v-vapham',	'v-vedia',	'v-vikore',	'v-vmuruz',	'v-wedeng',	'v-weihu',	'v-yakawa',	'v-ylain',	'v-yogek',	'v-zawaud',	'v-zhagak',	'v-zhweng',	'v-zubroz',	'v-zumaro')
		            --			   then 1
		            --			   else 0
		            --			  end)
		            --		 from [MarketingRequestNotifiedUsers] WITH (NOLOCK)
		            --		 group by MarketingRequestId) IRRN_SDL2
		            --on MR.MarketingRequestId=IRRN_SDL2.MarketingRequestId	


            --=================================================	
            -- Ticket level connection
            --=================================================

	            left join WorkflowStatus as WSMST WITH (NOLOCK)
		            on MST.StatusId = WSMST.WorkflowStatusID	
	            left join FactoryWorkflowStep as FWSMST WITH (NOLOCK)
		            on FWSMST.WorkflowStepID = MST.CurrentWorkflowStepId
	
	            --=================================================	
	            -- TicketLastUpdateDate = date when last task per ticket was completed (for completed ticket, it is always date when publish step was completed)
	            -- +
	            -- Working Minutes Per Ticket= all working minutes per ticket without iteration + working minutes including iteration	
	            --=================================================
	
		            left join -- last updated date of ticket
				            (select MSTT.MarketingServiceTicketId--, FWSMSTT.WorkflowStepName
			            , max(TaskIteration.Iteration) as Iteration
			            , max(TaskEndTime) as TicketLastUpdateDate
			            , sum(ElaspedTime) as TicketWorkingMinutes
			            , sum(case when IterationWM is null then ElaspedTime else IterationWM+ElaspedTime end) as TicketWorkingMinutesSUM
			            , sum(case when IterationWM is null and FWSMSTT.WorkflowStepName in ('Review','Build','Quality Check') then ElaspedTime
					               when IterationWM is not null and FWSMSTT.WorkflowStepName in ('Review','Build','Quality Check') then IterationWM+ElaspedTime
					               else 0 end) as TicketWorkingMinutesSUMProduction --
			
			            --------------Additional Stage

			            , sum(case when IterationWM is null and FWSMSTT.WorkflowStepName in ('Review') then ElaspedTime
					               when IterationWM is not null and FWSMSTT.WorkflowStepName in ('Review') then IterationWM+ElaspedTime
					               else 0 end) as TicketWorkingMinutesSUM_Review --

					               , sum(case when IterationWM is null and FWSMSTT.WorkflowStepName in ('Build') then ElaspedTime
					               when IterationWM is not null and FWSMSTT.WorkflowStepName in ('Build') then IterationWM+ElaspedTime
					               else 0 end) as TicketWorkingMinutesSUM_Build --

			            , sum(case when IterationWM is null and FWSMSTT.WorkflowStepName in ('Quality Check') then ElaspedTime
					               when IterationWM is not null and FWSMSTT.WorkflowStepName in ('Quality Check') then IterationWM+ElaspedTime
					               else 0 end) as TicketWorkingMinutesSUM_QC --


			            , sum(case when IterationWM is null and FWSMSTT.WorkflowStepName in ('Approve') then ElaspedTime
					               when IterationWM is not null and FWSMSTT.WorkflowStepName in ('Approve') then IterationWM+ElaspedTime
					               else 0 end) as TicketWorkingMinutesSUM_Approve --

			            , sum(case when IterationWM is null and FWSMSTT.WorkflowStepName in ('Publish') then ElaspedTime
					               when IterationWM is not null and FWSMSTT.WorkflowStepName in ('Publish') then IterationWM+ElaspedTime
					               else 0 end) as TicketWorkingMinutesSUM_Publish --

			
			            , sum(case when IterationWM is null and FWSMSTT.WorkflowStepName in ('Submitted') then ElaspedTime
					               when IterationWM is not null and FWSMSTT.WorkflowStepName in ('Submitted') then IterationWM+ElaspedTime
					               else 0 end) as TicketWorkingMinutesSUM_submit --

			
					
			
				            from MarketingServiceTicketTask as MSTT WITH (NOLOCK)


				            -- Additional working minutes must also be added in case of QC or OA rejection
					            left join 
						            (select ServiceTicketTaskId
								            ,'IterationWM' = sum (ElapsedTime)
								            ,'Iteration' = max (IterationNumber) 
						            from WorkingMinutesHistory WITH (NOLOCK)
						            group by ServiceTicketTaskId
						            ) as TaskIteration
						            on MSTT.ServiceTicketTaskId=TaskIteration.ServiceTicketTaskId
					            left join FactoryWorkflowStep as FWSMSTT WITH (NOLOCK)--
						            on FWSMSTT.WorkflowStepID = MSTT.WorkflowStepId--
					

				            group by MSTT.MarketingServiceTicketId) as MSTLatestUpdate
			            on MST.MarketingServiceTicketID = MSTLatestUpdate.MarketingServiceTicketID
	
	
	            --=================================================
	            -- Resource connection
	            --=================================================
		            /* Assigned User. This section connects Ticket to billing sheets to find out 
		               what is MSC location, Factory, FY14PrimaryRole +additional info from the billing sheets like Shift (for Delhi) etc.
		               As one ticket has multiple tasks, we need to decide who is going to be the winner for that ticket. Will ticket belong to MSC Cairo whose SOP did triage or to
		               the builder from Manila who did last build or to publish specialist who published the task in Beirut? There are couple of possibilities. 
		               We decided to eliminate users in Quality Check (due to shorten steps in Rapid Requests etc.), Approve, Publish, Completed (4,5,6,7) as they may be already from different center. e.g. Digital requests that are built in Beirut but publish is done by Cairo users
		   
		               Rule for selecting the right Resource: 1. If there is still no build assigned, we pick assigned user for the latest step before build
														            a. If ticket is still not triaged, we pick up the person who created ticket
														            b. If ticket has been triaged, we pick up Factory Lead (FL- person to whom review step has been assigned)
												              2. If builds are already created and assigned, we pick up Factory Specialist (FS) who has highest Working minutes on the build per whole ticket
														            a. If we have two build tasks per ticket and user A has completed the build and put 20 WM, he will be picked up over the user B if user B put 10 WM
														            b. If builders have not finished the build yet and no WM are stored, latest assigned builder is picked up
												              3. If two FS put the same amount of WM, we will pick up the one who was the latest.
		            */
		            left join
				            (select MarketingServiceTicketId
					              ,'AssignedUser'= (select case 
													            when charindex('\',MU.UserAlias) > 0 then Right(MU.UserAlias,len(MU.useralias)-(charindex('\',MU.UserAlias)))
													            else ''
												             end 
										              from MarketingUser WITH (NOLOCK) where Alias.AssignedUserId=UserId)
					               ,Alias.TaskAssignedOREndTime
					               ,MU.UserId
					               --,TaskEndTime
					               --,WorkflowStepId
					               --,Ranking
				            from
							            (select MSTT.MarketingServiceTicketId
									            ,MSTT.ServiceTicketTaskId									
									            ,'Ranking'= rank() over (partition by MarketingServiceTicketId order by WorkflowStepId desc, case when TaskIteration.IterationWM is not null then ElaspedTime+TaskIteration.IterationWM else ElaspedTime end desc, case when TaskEndTime is null then TaskAssignedTime else TaskEndTime end desc, MSTT.ServiceTicketTaskId desc)
									            ,WorkflowStepId
									            ,AssignedUserId
									            ,TaskAssignedTime
									            ,TaskEndTime
									            ,'TaskAssignedOREndTime' = case when TaskEndTime is null then TaskAssignedTime else TaskEndTime end --in case task is still not finished, we will take TaskAssignedTime to report even ongoing, not finished tasks
									            ,ElaspedTime
									            ,TaskIteration.IterationWM
									            ,'WorkingMinutesSUM' = case when TaskIteration.IterationWM is not null then ElaspedTime+TaskIteration.IterationWM else ElaspedTime end
							             from MarketingServiceTicketTask as MSTT WITH (NOLOCK)
								             -- Additional working minutes must also be added in case of QC or OA rejection
										            left join 
											            (select ServiceTicketTaskId
													            ,'IterationWM' = sum (ElapsedTime)
													            ,'Iteration' = max (IterationNumber) 
											            from WorkingMinutesHistory WITH (NOLOCK)
											            group by ServiceTicketTaskId
											            ) as TaskIteration
											            on MSTT.ServiceTicketTaskId=TaskIteration.ServiceTicketTaskId
							             where MSTT.WorkflowStepId not in (4,5,6,7) --we do not want to pick up person from Quality Check, Approve Publish, Completed		
							            ) as Alias 
						            left join MarketingUser as MU WITH (NOLOCK)
							            on Alias.AssignedUserId=MU.UserId
				            where Ranking=1 /*picking up the best ranking based on the rules described above*/) as AssignedUser
				
				            left join (select distinct UserId,ServiceCenterId from MarketingServiceCenterUser WITH (NOLOCK)) GSGUser	
					            on GSGUser.UserId=AssignedUser.UserId
				            left join MarketingServiceCenter GSGUserCenter WITH (NOLOCK)
					            on GSGUser.ServiceCenterId=GSGUserCenter.ServiceCenterId
			
			            on MST.MarketingServiceTicketId=AssignedUser.MarketingServiceTicketId
	
		
		
	            --=================================================
	            -- Current Ticket assigned user
	            --=================================================
		            left join
			
			            (select MarketingServiceTicketId
						              ,'CurrentTicketAssignedUser'= (select case 
														            when charindex('\',MU.UserAlias) > 0 then Right(MU.UserAlias,len(MU.useralias)-(charindex('\',MU.UserAlias)))
														            else ''
													             end 
											              from MarketingUser WITH (NOLOCK) where Alias.AssignedUserId=UserId)
						               ,Alias.TaskAssignedOREndTime
						               ,MU.UserId
						               --,TaskEndTime
						               --,WorkflowStepId
						               --,Ranking
					            from
								            (select MSTT.MarketingServiceTicketId
										            ,MSTT.ServiceTicketTaskId									
										            ,'Ranking'= rank() over (partition by MarketingServiceTicketId order by WorkflowStepId desc, case when TaskEndTime is null then TaskAssignedTime else TaskEndTime end desc, MSTT.ServiceTicketTaskId desc)
										            ,WorkflowStepId
										            ,AssignedUserId
										            ,TaskAssignedTime
										            ,TaskEndTime
										            ,'TaskAssignedOREndTime' = case when TaskEndTime is null then TaskAssignedTime else TaskEndTime end --in case task is still not finished, we will take TaskAssignedTime to report even ongoing, not finished tasks
								             from MarketingServiceTicketTask as MSTT WITH (NOLOCK)
								            ) as Alias 
							            left join MarketingUser as MU WITH (NOLOCK)
								            on Alias.AssignedUserId=MU.UserId
					            where Ranking=1 /*picking up the best ranking based on the rules described above*/) as CurrentTicketAssignedUser
			            on MST.MarketingServiceTicketId=CurrentTicketAssignedUser.MarketingServiceTicketId
		
				
	            --=================================================			
	            -- Ticket complexity and priority connection
	            --=================================================	
		
		            left join ServiceTicketComplexity as MSTComplexity WITH (NOLOCK)
			            on MST.ComplexityId=MSTComplexity.ComplexityId
		            left join ServiceTicketPriority as MSTPriority WITH (NOLOCK)
			            on MST.PriorityId=MSTPriority.PriorityId
	
	            --=================================================	
	            -- Date formats connection
	            --=================================================	

		            left join Storm_reporting.msc.RP_DateFormats as DateFormats WITH (NOLOCK)
			            on replace(convert(varchar(20),convert(date,MSTLatestUpdate.TicketLastUpdateDate)),'-','')=DateFormats.[DateYYYYMMDDKey]

	            --=================================================	
	            -- LTPS connection
	            --=================================================	
				                     
                    left join 
                                (select       MSTT.MarketingServiceTicketId                                               
                                                ,MST.MarketingServiceID
                                                ,'TimeOfLastTriage_PST' = max(case when MSTT.StatusID = 3 /*Complete*/ and MSTT.WorkflowStepId= 1 /*Triage*/ then MSTT.TaskEndTime else null end)
                                                ,'TimeOfLastAssignment_PST' = max(case when MSTT.StatusID = 3 /*Complete*/ and MSTT.WorkflowStepId= 2 /*Assignment*/ then MSTT.TaskEndTime else null end)
                                                ,'TimeOfLastBuild_PST' = max(case when MSTT.StatusID = 3 /*Complete*/ and MSTT.WorkflowStepId= 3 /*Build*/ then MSTT.TaskEndTime else null end)
                                                ,'TimeOfLastQualityCheck_PST'= max(case when MSTT.StatusId = 3 /*Complete*/ and MSTT.WorkflowStepId= 4 /*Quality Check*/ then MSTT.TaskEndTime else null end)
									            ,'SO_SubmittedDate_PST' = case when MS.SubmittedDate is null then MS.CreatedDate else MS.SubmittedDate end
									            ,'SOPAssignedDate_PST' = MS.SOPAssignedDate
									            ,'TicketCreatedDate_PST'= MST.CreatedDate
                                                ,'TicketNotStartBeforeDate_PST'= MST.NotStartBeforeDate
                                         
                                                                                   
                                                                           
                                                ,MRArea.AreaName
                                                ,MRSub.SubsidiaryName
                                                ,MRSub.TimeZoneId
                                  
                                from MarketingServiceTicketTask MSTT WITH (NOLOCK)
                                        left join MarketingServiceTicket as MST WITH (NOLOCK)
                                                on MSTT.MarketingServiceTicketId=MST.MarketingServiceTicketId
                                        left join MarketingService as MS WITH (NOLOCK)
                                                on MST.MarketingServiceID = MS.MarketingServiceID
                                        left join MarketingRequest as MR WITH (NOLOCK)
                                                on MS.MarketingRequestId = MR.MarketingRequestId
                                        left join MarketingSubsidiary as MRSub WITH (NOLOCK)
                                                on MR.SubsidiaryID = MRSub.SubsidiaryID
                                        left join MarketingArea as MRArea WITH (NOLOCK)
                                                on MRSub.AreaId=MRArea.AreaID
                                        left join MarketingServiceCenter MRCenter WITH (NOLOCK)
                                                on MRSub.Servicecenterid = MRCenter.ServiceCenterId
                                group by MSTT.MarketingServiceTicketId
                                                    ,MST.MarketingServiceID
                                                    ,case when MS.SubmittedDate is null then MS.CreatedDate else MS.SubmittedDate end
                                                    ,MS.SOPAssignedDate
										            ,MST.CreatedDate
                                                    ,MST.NotStartBeforeDate
                                                
                                                    ,MRArea.AreaName
                                                    ,MRSub.SubsidiaryName
                                                    ,MRSub.TimeZoneId
                                                
                                ) LTPS
                            on MST.MarketingServiceTicketId=LTPS.MarketingServiceTicketId

	
	            --=================================================	
	            -- Ticket Report tracking header connection
	            --=================================================	
		            --=================================================	
		            -- RTH for DigitalFormMetaData (Website SOs- MSCOM, ENT, MPN, SMB, PMC and other sites)
		            --=================================================				
				
				            left join DigitalFormMetaData RTHd WITH (NOLOCK)
					            on MST.MarketingServiceTicketId=RTHd.MarketingServiceTicketId
				            left join Site WITH (NOLOCK)
					            on Site.SiteId=RTHd.SiteId
			
			
		            --=================================================	
		            -- RTH for ExecuteMarketingMetaData (Email Creation/Delivery, Online Profiling, Event Creation/Management SOs)
		            --=================================================				
				
				            left join ExecuteMarketingMetaData RTHe WITH (NOLOCK)
					            on MST.MarketingServiceTicketId=RTHe.MarketingServiceTicketId
				            left join EventType EventTypeE WITH (NOLOCK)
					            on EventTypeE.eventtypeid=RTHe.eventtypeid 

		            --=================================================	
		            -- RTH for AnalyzeResultMetaData (Email Reports, Event Reports, Wizard Reports, Website Reports, Campaign Reports)
		            --=================================================				
				
				            left join AnalyzeResultMetaData RTHa WITH (NOLOCK)
					            on MST.MarketingServiceTicketId=RTHa.MarketingServiceTicketId
				            left join EventType EventTypeA WITH (NOLOCK)
					            on EventTypeA.eventtypeid=RTHa.eventtypeid 				
	
		            --=================================================	
		            -- RTH for ManageResponseMetaData (Newsletter Subscription Upload, Cloud Prospecting Upload, Offline Evaluations Upload, Lead Match/Upload, Event Attendance Upload)
		            --=================================================				
				
				            left join ManageResponseMetaData RTHm WITH (NOLOCK)
					            on MST.MarketingServiceTicketId=RTHm.MarketingServiceTicketId
				            left join EventType EventTypeM WITH (NOLOCK)
					            on EventTypeM.eventtypeid=RTHm.eventtypeid 	
				

		            /*
			            select * from TargetYourCustomerMetaData
			            select * from ExecuteMarketingMetaData
			            select * from DigitalFormMetaData
			            select * from ManageResponseMetaData
			            select * from AnalyzeResultMetaData

			            --Event Reports  --analyze results
			            --Event Attendance Upload --manage responses
			            --Offline Evaluation Upload --manage responses
			            --Event Creation/Management --event creation/management

			            --TargetYourCustomerMetaData
				            select * from Program

			            --DigitalFormMetaData
				            select * from DigitalFormMetaData
				            select * from Site
				            select * from SiteAndSiteSection
				            select * from SitePageComponent
		            */


 
            ) abc
	
            left join storm..SuperSOFormEventDetails ed (nolock) on abc.MS_ID = ed.MarketingServiceId
            left join storm..SuperSOEventType et (nolock) on et.eventtypeid = ed.eventtypeid
            left join storm..[ContentMarketingAttachment] ATT (nolock) on ATT.[MarketingServiceId] = abc.MS_ID 
            left join storm..[Attachment] AttUrl (nolock) on AttUrl.[AttachmentId] = ATT.[AttachmentId]

            Left Join(select * from (Select  MSTT.MarketingServiceTicketId,MU.Fullname  as AssignedTo, K2SerialNum
            , ROW_NUMBER() OVER(PARTITION BY MarketingServiceTicketId ORDER BY TaskAssignedTime desc) num

             from 
            MarketingServiceTicketTask MSTT (Nolock) Left Outer Join MarketingUser (Nolock)MU on MSTT.AssignedUserId=MU.UserId
            Where MSTT.StatusId=13
            ) abc 
            where Num = 1) AssignedUser on abc.MST_ID=AssignedUser.MarketingServiceTicketID


            left join (select * from (Select  MSTT.MarketingServiceTicketId,MU.FirstName+' '+MU.LastName  as BuildExecutioner, MU.EmailAddress as [BuilderEmail]
            , ROW_NUMBER() OVER(PARTITION BY MarketingServiceTicketId ORDER BY TaskAssignedTime desc) num

             from 
            MarketingServiceTicketTask MSTT (Nolock) Left Outer Join MarketingUser (Nolock)MU on MSTT.AssignedUserId=MU.UserId
            Where MSTT.workflowstepid=3
            ) abc 
            where Num = 1
            ) BuildUser on abc.MST_ID=BuildUser.MarketingServiceTicketID

            left join (select * from (Select  MSTT.MarketingServiceTicketId,MU.FirstName+' '+MU.LastName  as PeerReviewer , MU.EmailAddress as [PeerEmail]
            , ROW_NUMBER() OVER(PARTITION BY MarketingServiceTicketId ORDER BY TaskAssignedTime desc) num

             from 
            MarketingServiceTicketTask MSTT (Nolock) Left Outer Join MarketingUser (Nolock)MU on MSTT.AssignedUserId=MU.UserId
            Where MSTT.workflowstepid=4
            ) abc 
            where Num = 1) PeerReviewer on abc.MST_ID= PeerReviewer.MarketingServiceTicketID
            
            
            left join (select * from (Select  MSTT.MarketingServiceTicketId,MU.FirstName+' '+MU.LastName  as FactoryLead , MU.EmailAddress as [FactoryLeadEmail]
            , ROW_NUMBER() OVER(PARTITION BY MarketingServiceTicketId ORDER BY TaskAssignedTime desc) num

             from 
            MarketingServiceTicketTask MSTT (Nolock) Left Outer Join MarketingUser (Nolock)MU on MSTT.AssignedUserId=MU.UserId
            Where MSTT.workflowstepid=2
            ) abc 
            where Num = 1) FactoryLead on abc.MST_ID= FactoryLead.MarketingServiceTicketID
           

            left join ContentMarketingInitialInformation CSII (nolock) on csii.MarketingServiceId = abc.MS_ID
            where 
            --abc.ServiceTypeName in ('Content Marketing')
            --and 
			abc.TicketStatusName =  'Assigned to Factory'
            and abc.[TicketStepName] = 'Approve'
            and abc.areaname not in ('SmokeTest-Area')
            --and AttUrl.[FileName] like '%.xlsm'
            and abc.TicketTag not like '%Hold%'
            and abc.program = 'Office Demand Center'
           -- and csii.GatedContentReasonId = 2 
            ) A where [Gap of Days] in (2,5,8)";
            return str;
        }
    }
}
