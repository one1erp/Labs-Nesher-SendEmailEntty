using System;
using System.Runtime.InteropServices;
using Common;
using LSEXT;

namespace SendEmailEntty
{
    [ComVisible(true)]
    [ProgId("SendEmailEntty.SendEmailWf")]
    public class SendEmailWf : IWorkflowExtension//not using
    {


        public const bool BckgrndSend = true;
        public void Execute(ref LSExtensionParameters Parameters)
        {

            try
            {
                var sp = Parameters["SERVICE_PROVIDER"];
                var records = Parameters["RECORDS"];
                var id = records.Fields["U_COA_REPORT_ID"].Value;
                var ntlsCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlsCon);
                Send.SendEmail(id.ToString(), BckgrndSend);
            }
            catch (Exception ex)
            {

                Common.Logger.WriteLogFile(ex);
            }


        }
    }
}
