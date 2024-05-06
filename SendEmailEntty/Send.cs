using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Common;
using DAL;
using Attachment = System.Net.Mail.Attachment;
using Exception = System.Exception;

namespace SendEmailEntty
{
    public static class Send
    {



        private static COA_Report currentCoa;
        private static MailDetails mailDetails;
        private static string PHRASE_NAME = "EmailParameters";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="coaReportId">Specified COA ID</param>
        /// <param name="directly">Send directly or open outlook</param>
        /// 
        public static void SendEmail(string coaReportId, bool directly)
        {
            Debugger.Launch();

            //not used???
            var dal = new DataLayer();
            dal.Connect();

            currentCoa = dal.GetCoaReportById(Convert.ToInt32(coaReportId));

            string labName = currentCoa.Sdg.LabInfo.Name;
            string emails = currentCoa.Sdg.Emai;
            string[] emailsTo = emails.Split(';');


            if (currentCoa.ClientId != null)
            {
                // var toList = dal.GetAddresses("CLIENT", (long) currentCoa.ClientId);
                var emailParam = dal.GetPhraseByName(PHRASE_NAME);

                mailDetails = new MailDetails();

                SetMailDetails(mailDetails, emailParam, labName, currentCoa);

                foreach (var email in emailsTo)
                {
                    if (!string.IsNullOrEmpty(email))
                        mailDetails.To.Add(email.Trim());
                }



                string document = currentCoa.PdfPath;
                if (string.IsNullOrEmpty(document))
                {
                    document = currentCoa.DocPath;
                }
                mailDetails.AtachmentPathes = new List<string> { document };




                bool sent;
                if (directly)
                {
                    sent = MailService.Send(mailDetails);

                }
                else
                {
                    sent = MailService.OpenOutlook(mailDetails);
                }

                if (sent)
                {
                    currentCoa.SentOn = DateTime.Now;
                    dal.SaveChanges();
                }
                dal.Close();

            }
        }

        public static void SendMultepleCOA(List<string> coaReportIdList, bool bckgrndSend)
        {

            COA_Report coa = null;
            var dal = new DataLayer();
            dal.Connect();
            string labName = "";
            string emails = "";
            var listCoaReports = new List<COA_Report>();
            mailDetails = new MailDetails();
            var listOfAtach = new List<string>();
            foreach (var coaReportId in coaReportIdList)
            {
                coa = dal.GetCoaReportById(Convert.ToInt32(coaReportId));
                listCoaReports.Add(coa);
                labName = coa.Sdg.LabInfo.Name;
                emails = emails + coa.Sdg.Emai + ";";


                string document = coa.PdfPath;
                if (string.IsNullOrEmpty(document))
                {
                    document = coa.DocPath;

                }
                listOfAtach.Add(document);
            }
            string[] emailsTo = emails.Split(';');


            emailsTo = emailsTo.Distinct().ToArray();
            var emailParam = dal.GetPhraseByName(PHRASE_NAME);

         

            SetMailDetails(mailDetails, emailParam, labName, coa);

            foreach (var email in emailsTo)
            {
                if (!string.IsNullOrEmpty(email))
                    mailDetails.To.Add(email.Trim());
            }

            mailDetails.AtachmentPathes = listOfAtach;

    
            //Ashi  - Show popup when no mail addres in sdg  (8.2.2018)
            bool hasAddreDest = mailDetails.To.Count > 0 && mailDetails.To.All(IsValidEmail);
            if (!hasAddreDest)
            {
                string msg = " ";
                if (coa != null)
                {
                    msg += coa.Client.Name;
                    MessageBox.Show("לא קיימת כתובת מייל ללקוח!" + msg);
                    if (bckgrndSend || coaReportIdList.Count != 1)//אם השליחה היא דרך אאוטלוק וזה רק תעודה אחת צריך להמשיך בתהליך כדי שיפתח האאוטלוק 
                        return;
                }
            }

            bool sent;
            if (bckgrndSend)
            {
                sent = MailService.Send(mailDetails);

            }
            else
            {
                sent = MailService.OpenOutlook(mailDetails);
            }


            if (sent && hasAddreDest)
            {
                foreach (var coaReport in listCoaReports)
                {
                    coaReport.SentOn = DateTime.Now;
                }
                //Update sdg at sent              
                if (coa != null) coa.Sdg.COASent = "T";

                dal.SaveChanges();
            }
            dal.Close();


        }


        private static void SetMailDetails(MailDetails mailDetails, PhraseHeader emailParam, string labName, COA_Report currentCoa)
        {


            mailDetails.UserName = (from smtp in emailParam.PhraseEntries
                                    where smtp.PhraseName == "UserName_" + labName

                                    select smtp.PhraseDescription).FirstOrDefault();

            mailDetails.Password = (from smtp in emailParam.PhraseEntries
                                    where smtp.PhraseName == "Password_" + labName
                                    select smtp.PhraseDescription).FirstOrDefault();

            mailDetails.SmtpClient = (from smtp in emailParam.PhraseEntries
                                      where smtp.PhraseName == "exchange server"
                                      select smtp.PhraseDescription).FirstOrDefault();

            HeaderDetails mailDetailsHeader = CreateAndSetHeaderDtls(currentCoa.Sdg);              
            mailDetails.Subject = mailDetailsHeader.ToString();


            mailDetails.FromAddress = (from smtp in emailParam.PhraseEntries
                                       where smtp.PhraseName == "FromFile_" + labName
                                       select smtp.PhraseDescription).FirstOrDefault();

            var cc = (from smtp in emailParam.PhraseEntries
                      where smtp.PhraseName == "CC_" + labName
                      select smtp.PhraseDescription).FirstOrDefault();
            if (cc != null)
            {
                mailDetails.CC.Add(cc);
            }
        }

        private static HeaderDetails CreateAndSetHeaderDtls(Sdg currentSdg)
        {
            return new HeaderDetails
            {
                OrderName = currentSdg?.Name,
                ExternalRef = currentSdg?.ExternalReference,
                CoaFile = currentSdg?.SDG_USER?.FirstOrDefault()?.U_COA_FILE,
                FirstSampleDetails = currentSdg?.Samples?.FirstOrDefault()?.Description
            };

        }



        static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }



    }
    public class HeaderDetails
    {
        public string OrderName { get; set; }
        public string ExternalRef { get; set; }
        public string CoaFile { get; set; }
        public string FirstSampleDetails { get; set; }


        public override string ToString()
        {
            return string.Join(" - ", new[] { OrderName, ExternalRef, CoaFile, FirstSampleDetails }.Where(s => !string.IsNullOrEmpty(s)));
        }


    }
}

