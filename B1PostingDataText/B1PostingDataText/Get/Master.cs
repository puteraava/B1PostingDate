using B1PostingDataText.Connection;
using B1PostingDataText.Model.MasterData.Customer;
using Dapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.IO;
using System.Data.SqlClient;
using System.Data;

namespace B1PostingDataText.Get
{
    public class Master:DapperConnection
    {
        public static List<Customer> GetDataCustomer()
        {
            try
            {
                //string queryheader  = Properties.QueryGet.CustomerHeader;
                //string querydetails = Properties.QueryGet.CustomerDetails;
                //string queryflag = Properties.QueryGet.Update_Flag;
                //string querycustomer = Properties.QueryGet.IDU_GET_Customer;

                SqlConnection SQLCONNECT = new SqlConnection("Data Source=DESKTOP-LBNT4J0;Initial Catalog=SBODemoAU;User ID=sa;Password=devit");
                using (var conn = SQLCONNECT.QueryMultiple("EXEC _IDU_GET_CUSTOMER"))
                {
                    //List<Customer> customer = conn.Query<Customer>(querycustomer).ToList();
                    var customer = conn.Read<Customer>().ToList();

                    var builder = new StringBuilder();
                    builder.AppendLine("Outlet Code,Outlet Name,Level 1,Level 2,Level 3,Level 4,Level 5,Owner,Phone 1,Phone 2,Address,Category,Credit Limit,Longitude,Latitude,Price Group,Area");
                    foreach (var item in customer)
                    {
                        builder.AppendLine($"{item.CardCode},{item.CardName},{item.U_IDU_Lv1},{item.U_IDU_Lv2},{item.U_IDU_Lv3},{item.U_IDU_Lv4},{item.U_IDU_Lv5},{item.CntctPrsn},{item.Phone1},{item.Phone2},{item.Address},{item.GroupName},{item.CreditLine},{item.U_Longitude},{item.U_Latitude},{item.U_PriceGroup},{item.Country}");
                    }
                    //if (File.Exists("C:\\B11000_2405-70004131\\Customer.csv"))
                    //{
                    //    File.Delete("C:\\B11000_2405-70004131\\Customer.csv");
                    //}
                    using (var writer = File.AppendText("C:\\B11000_2405-70004131\\Customer.csv")) { writer.Write(builder); };
                    
                    return customer;
                }
            }
            catch (Exception exc)
            {
                throw new Exception(string.Format("[{0}] [{1}]", MethodBase.GetCurrentMethod().Name, exc.Message));
            }
        }
    }
}
