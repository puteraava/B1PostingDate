using B1PostingDataText.Functions;
using System;

namespace B1PostingDataText.B1Connection
{
    public class SdkConnection
    {
        public static SAPbobsCOM.Company MyCompany { get; set; }  
        public static SAPbobsCOM.Company GetCompany()
        {
            try
            {
                if (MyCompany == null)
                    MyCompany = new SAPbobsCOM.Company();
                if (MyCompany.Connected)
                {
                    return MyCompany;
                }
                MyCompany.Server = "DESKTOP-LBNT4J0";
                MyCompany.DbUserName = "sa";
                MyCompany.DbPassword = "devit";
                MyCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
                MyCompany.CompanyDB = "SBODemoAU";
                MyCompany.UserName = "B1SiteUser";
                MyCompany.Password = "Devit1234@";
                MyCompany.LicenseServer = "DESKTOP-LBNT4J0:40000";
                MyCompany.SLDServer = "https://DESKTOP-LBNT4J0:40000";

                if (MyCompany.Connect() != 0)
                {
                    throw new Exception(MyCompany.GetLastErrorDescription().ToString());
                }
            }
            catch (Exception ex)
            {
                Tracelog.TransWriteLine(ex.Message);
            }
            return MyCompany;
        }
    }
}