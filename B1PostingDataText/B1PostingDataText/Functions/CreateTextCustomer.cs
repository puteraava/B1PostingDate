using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace B1PostingDataText.Functions
{
    public class CreateTextCustomer
    {
        public static void CreateCustomer()
        {
            var json = new JavaScriptSerializer().Serialize(Get.Master.GetDataCustomer());
            using (var writer = File.AppendText("C:\\B11000_2405-70004131\\Customer.txt")) { writer.Write(json); };
            
            //Update Flag is_post from N to Y
            SqlConnection SQLCONNECT = new SqlConnection("Data Source=DESKTOP-LBNT4J0;Initial Catalog=SBODemoAU;User ID=sa;Password=devit");
            SqlCommand cmd = new SqlCommand("EXEC _IDU_UPDATE_CUSTOMER_POST",SQLCONNECT);
            SQLCONNECT.Open();
            cmd.ExecuteNonQuery();
            SQLCONNECT.Close();
            //SdkConnection.GetCompany();
        }
        //SdkConnection.GetCompany();
    }
}
