using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Collections.Specialized;
using System.Net.Mail;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelGeneration
{
    public partial class WebForm2 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void Log_Click(object sender, EventArgs e)
        {
            Response.Redirect("WebForm1.aspx");
        }
        protected void Cancel_Click(object sender, EventArgs e)
        {
            TextExcel.Text = "";
            TextEmail.Text = "";
           
        }

        protected void Start_Click(object sender, EventArgs e)
        {
            string strPAN;
            //int n;
            DataTable dtSupplier;
            string strFileName, strFilePath;
            string strSupplierCode, strName, strEmail;
            string strEmailBody;
            int iPAN;
            string quarter;
            string strfy;
            string strfname = TextExcel.Text + ".xlsx";
            //string res = "Yes";
            //string ress = "No";
            //string strFinancialYear = Convert.ToString(ConfigurationManager.AppSettings["FinancialYear"]);
            //string strQuarter = Convert.ToString(ConfigurationManager.AppSettings["Quarter"]);
            string strEmailTo = Convert.ToString(ConfigurationManager.AppSettings["EmailTo"]);
            string strEmailCc = Convert.ToString(ConfigurationManager.AppSettings["EmailCc"]);
            string strEmailBcc = Convert.ToString(ConfigurationManager.AppSettings["EmailBcc"]);
            string strTestEmail = Convert.ToString(ConfigurationManager.AppSettings["IsTestEmail"]);
            string strEmailSub = Convert.ToString(ConfigurationManager.AppSettings["EmailSubject"]);
            string strEmailDisp = Convert.ToString(ConfigurationManager.AppSettings["EmailDisplayName"]);
            //string strDoc = Convert.ToString(ConfigurationManager.AppSettings["DocumentPath"]);
            Microsoft.Office.Interop.Excel._Application oApp;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel._Workbook oBook;
            oApp = new Microsoft.Office.Interop.Excel.Application();
            oBook = oApp.Workbooks.Add();
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oBook.Worksheets.get_Item(1);
            oSheet.Cells[1, 1] = "SrNo";
            oSheet.Cells[1, 2] = "File Name";
            oSheet.Cells[1, 3] = "PAN";
            oSheet.Cells[1, 4] = "Supplier Code";
            oSheet.Cells[1, 5] = "Name";
            oSheet.Cells[1, 6] = "Email";
            oSheet.Cells[1, 7] = "IsEmailSent";
            if (FileUpload1.HasFile)
            {

                HttpFileCollection uploadedFiles = Request.Files;
                for (int i = 0; i < uploadedFiles.Count; i++)
                {
                    HttpPostedFile userPostedFile = uploadedFiles[i];
                    if (userPostedFile.ContentLength > 0)
                    {
                        string strfilename = Path.GetFileName(userPostedFile.FileName);

                        userPostedFile.SaveAs(Server.MapPath("~/Files/" + strfilename));
                        string[] filesPath = Directory.GetFiles(Server.MapPath("~/Files/"));
                        List<ListItem> files = new List<ListItem>();

                        foreach (string path in filesPath)
                        {
                            files.Add(new ListItem(Path.GetFileName(path)));


                        }

                        //gvDetails.DataSource = files;

                        //gvDetails.DataBind();
                        //string fname = @"C:\" + TextExcel.Text + ".xlsx";
                        //string fname = TextExcel.Text;
                        //string strfname = TextExcel.Text + ".xlsx";

                        if (File.Exists(strfname))
                        {
                            Page.ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('File Name Already Exists,Please choose another filename')");


                        }

                        for (int j = 1; j <= filesPath.Length; j++)
                        {
                            oSheet.Cells[j + 1, 1] = j;
                            strFilePath = filesPath[j - 1];
                            strFileName = Path.GetFileName(strFilePath);
                            iPAN = strFileName.IndexOf("_");

                            strPAN = strFileName.Substring(0, iPAN);
                            quarter = strFileName.Substring(iPAN, 13);
                            strfy = strFileName.Substring(13, '\0');
                            oSheet.Cells[j + 1, 2] = strFileName;
                            oSheet.Cells[j + 1, 3] = strPAN;

                            dtSupplier = getSupplierDetails(strPAN);

                            if (dtSupplier != null && dtSupplier.Rows.Count > 0)
                            {
                                strSupplierCode = Convert.ToString(dtSupplier.Rows[0]["SupplierCode"]);
                                strName = Convert.ToString(dtSupplier.Rows[0]["Name"]);


                                strEmail = Convert.ToString(dtSupplier.Rows[0]["Email"]);


                                oSheet.Cells[j + 1, 4] = strSupplierCode;
                        
                                
                                    oSheet.Cells[j + 1, 5] = strName;




                                    //if (int.TryParse(strEmail, out n))
                                    //{
                                    //    oSheet.Cells[j + 1, 6] = "Not Available";
                                    //}
                                    //else
                                    //{
                                        oSheet.Cells[j + 1, 6] = Convert.ToString(dtSupplier.Rows[0]["Email"]);

                                    //}
                                
                                    if (!string.IsNullOrEmpty(strEmail))
                                    {  
                                        List<string> lstEmailTo = new List<string>();
                                        List<string> lstAtt = new List<string>();
                                        List<string> lstEmailCc = new List<string>();
                                        List<string> lstEmailBcc = new List<string>();

                                        string[] EmailTo = strEmail.Split(',');

                                        foreach (string EmailAdd in EmailTo)
                                        {
                                            if (!string.IsNullOrEmpty(EmailAdd))
                                            {
                                                lstEmailTo.Add(EmailAdd);
                                            }
                                        }
                                        string[] EmailCc = strEmailCc.Split(',');

                                        foreach (string EmailAdd in EmailCc)
                                        {
                                            if (!string.IsNullOrEmpty(EmailAdd))
                                            {
                                                lstEmailCc.Add(EmailAdd);
                                                //lstEmailCc.Add("SHITAL.MERCHANT@larsentoubro.com");
                                                //lstEmailCc.Add("jayen.desai@larsentoubro.com");
                                                //lstEmailCc.Add("GBN.Babu@larsentoubro.com");
                                            }
                                        }

                                        string[] EmailBcc = strEmailBcc.Split(',');

                                        foreach (string EmailAdd in EmailBcc)
                                        {
                                            if (!string.IsNullOrEmpty(EmailAdd))
                                            {
                                                lstEmailBcc.Add(EmailAdd);
                                            }
                                        }
                                        strEmailBody = getMailTemplate().Replace("_customername_", strName).Replace("_Quareter_", quarter).Replace("_FY_", strfy);
                                        lstAtt.Add(strFilePath);

                                       
                                        bool IsSent = SendEmailwithAttachment(lstEmailTo, lstEmailBcc, lstEmailCc, lstAtt, strEmailSub, strEmailBody, strEmailBody, null, true, strEmailDisp);
                                        if (IsSent == true)
                                        {
                                            oSheet.Cells[j + 1, 7] = "yes";
                                        }
                                        else
                                        {
                                            oSheet.Cells[j + 1, 7] = "No";
                                        }
                                       
                                    
                                }
                                   
                            }
                            
                            if (oApp.Application.Sheets.Count < 1)
                            {
                                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oBook.Worksheets.Add();
                            }
                            else
                            {
                                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oApp.Worksheets[1];
                            }

                            //string strfname = TextExcel.Text + ".xlsx";


                        }
                       
                    }
                }
            }
                    oBook.SaveAs(Server.MapPath("~/excel/" + strfname));

                    //ClientScript.RegisterStartupScript(Page.GetType(), "validation", "<script language='javascript'>alert('Excel file created in solution folder Excel')</script>");
                    oBook.Close();
                   
                    oApp.Quit();

                    string[] filesexcel = Directory.GetFiles(Server.MapPath("~/excel/"));
                    //for (int k = 0; k <= filesexcel.Length; k++)
                    //{ string filemail =Path.GetFileName(strfname);
                    //}

                    //List<ListItem> filesmail = new List<ListItem>();
                    //foreach (string path in filesexcel)
                    //{
                    //    filesmail.Add(new ListItem(Path.GetFileName(path)));
                    //}
                    using (System.Net.Mail.MailMessage mm = new System.Net.Mail.MailMessage("SPSADMIN@HZW.LTINDIA.COM", TextEmail.Text))
                    {

                        mm.Body ="Attached Excel File";
                        //mm.Bcc.Add(new MailAddress("app.admin@larsentoubro.com"));
                        mm.IsBodyHtml = true;
                        mm.Subject = "Excel displaying files mailed to suppliers";
                        if (filesexcel != null)
                        {
                            if (filesexcel.Length>0)
                            {
                                foreach (string imageAttachment in filesexcel)
                                {
                                    FileInfo f1 = new FileInfo(imageAttachment);
                                    if (f1.Exists)
                                    {
                                        mm.Attachments.Add(new Attachment(imageAttachment));
                                    }
                                }
                            }
                        }
                       
                        SmtpClient smtp = new SmtpClient();
                        smtp.Host = "ssawf";
                        smtp.Port = 25;
                        smtp.Send(mm); 
                     
                    }
                    ClientScript.RegisterStartupScript(Page.GetType(), "validation", "<script language='javascript'>alert('Excel file mailed ')</script>");

                    //Label5.Visible = true;
                

            
        }


        public static bool SendEmailwithAttachment(List<String> lstEmailTo, List<String> lstEmailToBcc, List<String> lstEmailToCc, List<String> strAttachment, string subject, string mailBody, string strMailBodyPlainText, string strFromMailId, bool isbodyhtml, string strDisplayName)
        {
            bool valToReturn = true;
            try
            {
                System.Net.Mail.MailMessage Message = new System.Net.Mail.MailMessage();
                System.Net.Mail.SmtpClient sClient = new System.Net.Mail.SmtpClient();
                NameValueCollection nvc = new NameValueCollection();
                nvc.Add("MIME-Version", "1.0");
                nvc.Add("charset", "iso-8859-1");
                nvc.Add("From", Convert.ToString(ConfigurationManager.AppSettings["FromEmailID"]));
                DateTime dTime = DateTime.Now;
                Guid gd = Guid.NewGuid();
                string strMessageId = "<" + gd + "_" + dTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss");
                nvc.Add("Message-Id", strMessageId);
                Message.Headers.Add(nvc);
                //string[] ccem = { "shivangi.jhavar@larsentoubro.com", "JAYEN.DESAI@larsentoubro.com", "santosh.chitale@larsentoubro.com" };
                string[] pp = { "swapnil.shailee@sitpune.edu.in", "swapnilshailee27@gmail.com" };
                foreach (string addto in pp)
                        {
                           Message.To.Add(addto);
                        }
                        foreach (string addcc in pp)
                        {
                            Message.CC.Add(addcc);
                        }
                        foreach (string addbcc in pp)
                        {
                            Message.Bcc.Add(addbcc);
                        }
                //if (lstEmailTo != null && lstEmailTo.Count > 0)
                //{
                //    foreach (String addto in lstEmailTo)
                //    {
                //        if (!string.IsNullOrEmpty(addto))
                //        {
                //            Message.To.Add(addto);
                //        }
                //    }
                //}

                //if (lstEmailToBcc != null && lstEmailToBcc.Count > 0)
                //{
                //    foreach (String addtoBcc in lstEmailToBcc)
                //    {
                //        if (!string.IsNullOrEmpty(addtoBcc))
                //        {
                //            Message.Bcc.Add(addtoBcc);
                //        }
                //    }
                //}

                //if (lstEmailToCc != null && lstEmailToCc.Count > 0)
                //{
                //    foreach (String addtoCc in lstEmailToCc)
                //    {
                //        if (!string.IsNullOrEmpty(addtoCc))
                //        {
                //            Message.CC.Add(addtoCc);
                //        }
                //    }
                //}
                Message.From = new MailAddress("SPSADMIN@HZW.LTINDIA.COM");
                //Message.From = new MailAddress(ConfigurationManager.AppSettings["FromEmailID"], strDisplayName);
                Message.Subject = subject;
                Message.Body = mailBody;
                Message.IsBodyHtml = isbodyhtml;
                AlternateView plainView = AlternateView.CreateAlternateViewFromString(strMailBodyPlainText, System.Text.Encoding.UTF8, "text/plain");
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(mailBody, System.Text.Encoding.UTF8, "text/html");
                Message.AlternateViews.Add(plainView);
                Message.AlternateViews.Add(htmlView);
                if (strAttachment != null)
                {
                    if (strAttachment.Count > 0)
                    {
                        foreach (string imageAttachment in strAttachment)
                        {
                            FileInfo f1 = new FileInfo(imageAttachment);
                            if (f1.Exists)
                            {
                                Message.Attachments.Add(new Attachment(imageAttachment));
                            }
                        }
                    }
                }
                //for stream file
                //if (objStream != null)
                //{
                //    Message.Attachments.Add(new Attachment(objStream, DateTime.Now.Ticks.ToString() + ".pdf"));
                //}
                SmtpClient smtp = new SmtpClient();
                smtp.Host = "ssawf";
                smtp.Port = 25;

                //smtp.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["UseSSL"]);
                //smtp.Port = Convert.ToInt32(ConfigurationManager.AppSettings["MailPort"]);
                //smtp.Host = Convert.ToString(ConfigurationManager.AppSettings["MailHost"]);
                //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                // System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(Convert.ToString(ConfigurationSettings.AppSettings["NetworkCredentialEmail"]), Convert.ToString(ConfigurationSettings.AppSettings["NetworkCredentialPassword"]));
                //smtp.Credentials = basicAuthenticationInfo;
                //smtp.UseDefaultCredentials = true;
                foreach (var v in lstEmailTo)
                {
                    CreateLog("EmailSending.SendMailMultipleAttachmentsWithBody", "Before sending mail : " + v);
                    CreateLog("EmailSending.SendMailMultipleAttachmentsWithBody", "Mail body :" + mailBody);
                }

                smtp.Send(Message);

                foreach (var v in lstEmailTo)
                {
                    CreateLog("MailSender.SendMailMultipleAttachmentsWithBody", "Mail sent to : " + v);
                }
            }

            catch (Exception ex)
            {
                CreateLog("Mail sending failed", ex.Message);
                CreateLog("MailSender.SendMailMultipleAttachmentsWithBody", ex.Message);
                valToReturn = false;
            }
            return valToReturn;

        }


        public static void CreateLog(string strFunctionName, string strMessage)
        {
            try
            {
                if (ConfigurationManager.AppSettings["IsWriteException"] == "true")
                {
                    string strFilePath = ConfigurationManager.AppSettings["ExceptionFilePath"] + "ExceptionLog_" + DateTime.Now.ToString("ddMMyyyy") + ".txt";
                    FileStream fs = null;
                    StreamWriter sw = null;
                    if (!File.Exists(strFilePath))
                        fs = File.Create(strFilePath);
                    else
                        sw = File.AppendText(strFilePath);
                    sw = sw == null ? new StreamWriter(fs) : sw;
                    sw.WriteLine("Location : " + strFunctionName);
                    sw.WriteLine("DateTime : " + DateTime.Now);
                    sw.WriteLine("Message :" + strMessage);
                    sw.WriteLine("------------------------------------------------------------------------------------------------------------------------------");
                    sw.Close();
                    if (fs != null)
                        fs.Close();
                }
            }
            catch (Exception) { }
        }

        public string getMailTemplate()
        {
            return @"<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>
<html xmlns='http://www.w3.org/1999/xhtml'>
<head>
    <meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
    <title>Larsen &amp; Toubro</title>
</head>
<style>
    body
    {
        font-family: Verdana, Arial, Helvetica, sans-serif;
        font-size: 12px;
    }
    .table
    {
        border: solid 1px #999999;
    }
    td
    {
        padding: 10px;
    }
    .td
    {
        margin: 0px;
        padding: 0px;
        color: #FFFFFF;
        font-weight: bold;
    }
    p
    {
        margin-top: 10px;
    }
    .white_link
    {
        color: #FFFFFF;
    }
    .white_link:hover
    {
        color: #FFFFFF;
    }
</style>
<body>
    <table width='100%' border='0' cellspacing='0' cellpadding='0' class='table' align='center'>
        <tr>
            <td align='center'>
                <img src='http://www.larsentoubro.com/lntcorporate/Common/images/l&-T-logo.jpg' />
            </td>
        </tr>
        <tr>
            <td>
                <p>
                    &nbsp;</p>
                <strong>Dear _customername_,</strong>                
                <br />
                <p>
                Attached herewith TDS certificate for _Quareter_ of F.Y. _FY_. 
                Kindly arrange to take print of the certificate for your record after validation of signature.</p><br />
                <p>
                Attached herewith document for procedure to be followed for validation of signature.</p>
                
                <p>
                    &nbsp;</p>
                <p>
                    <b>Regards</b><br />    
                    Finance & Account<br />    
                    Heavy Engineering<br />     
                    L&T, Powai, Mumbai
                    </p>
            </td>
        </tr>
        <tr>
            <td>
                <br />
                <p>
                    <strong>Disclaimer:</strong></p>
                <p>
                    <i>This Email may contain confidential or privileged information for the intended recipient
                        (s).
                        <br />
                        If you are not the intended recipient, please do not use or disseminate the information,
                        notify the sender and delete it from your system. </i>
                </p>
            </td>
        </tr>
        <tr>
            <td bgcolor='#037afd'>
            </td>
        </tr>
    </table>
</body>
</html>

";
        }
        public DataTable getSupplierDetails(string pPAN)
        {

            try
            {
                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["lnapppddbConnectionString"].ConnectionString);

                string strSql = @" select distinct a.t_bpid SupplierCode,isnull(a.t_nama,'') Name,(case when c.t_info='' or (select count(d.t_bpid) from ttccom100180 as d left join ttctax400180 as e on  d.t_bpid=e.t_bpid where e.t_fovn=b.t_fovn group by e.t_fovn)>1 then 'Not Available'
else  c.t_info end) Email ,(select count(d.t_bpid) from ttccom100180 as d left join ttctax400180 as e on  d.t_bpid=e.t_bpid where e.t_fovn=b.t_fovn group by e.t_fovn)[count]
 from ttccom100180 as a 
                                left join ttctax400180 as b on a.t_bpid=b.t_bpid 
                                left join ttccom140180 as c on a.t_ccnt=c.t_ccnt where b.t_catg_l=1 and b.t_fovn= '" + pPAN + "' group by a.t_bpid,a.t_nama,c.t_info,b.t_fovn";
                SqlCommand cmd = new SqlCommand(strSql, conn);
                SqlDataAdapter Da = new SqlDataAdapter(cmd);
                DataTable Dt = new DataTable();
                Da.Fill(Dt);

                return Dt;

            }
            catch (SqlException)
            {
                throw;
            }
        }
    }
}
            


     
    