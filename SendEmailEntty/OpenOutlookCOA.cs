
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;

namespace SendEmailEntty
{


    [ComVisible(true)]
    [ProgId("SendEmailEntty.OpenOutlookCOA")]
    public class OpenOutlookCOA : IEntityExtension
    {
        private INautilusServiceProvider sp;
        public const bool BckgrndSend = false;
        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            sp = Parameters["SERVICE_PROVIDER"];
            var records = Parameters["RECORDS"];
            var ntlsCon = Utils.GetNtlsCon(sp);
            Utils.CreateConstring(ntlsCon);
            var CoaIds = new List<string>();

            while (!records.EOF)
            {
                var CoaId = records.Fields["U_COA_REPORT_ID"].Value;
                CoaIds.Add(CoaId);
                records.MoveNext();
            }


            var dal = new DataLayer();
            dal.Connect();
            var firstCoa = dal.GetCoaReportById(long.Parse(CoaIds.FirstOrDefault()));

            var firstCoaClientId = firstCoa.ClientId;

            foreach (var coaId in CoaIds)
            {
                var coa = dal.GetCoaReportById(long.Parse(coaId));
                var clientId = coa.ClientId;
                if (clientId != firstCoaClientId)
                {
                    dal.Close();
                    return ExecuteExtension.exDisabled;
                }
            }
            dal.Close();
            return ExecuteExtension.exEnabled;

        }
        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {


                sp = Parameters["SERVICE_PROVIDER"];
                var records = Parameters["RECORDS"];
                var ntlsCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlsCon);

                var CoaIds = new List<string>();

                while (!records.EOF)
                {
                    var CoaId = records.Fields["U_COA_REPORT_ID"].Value;
                    CoaIds.Add(CoaId);
                    records.MoveNext();
                }

                Send.SendMultepleCOA(CoaIds, BckgrndSend);
            }

            catch (Exception ex)
            {
                Common.Logger.WriteLogFile(ex);
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }
    }

}

