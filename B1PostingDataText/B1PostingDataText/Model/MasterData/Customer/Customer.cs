using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace B1PostingDataText.Model.MasterData.Customer
{
    #region Model Customer
    public class Customer
    {
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string CreditLine { get; set; }
        public string U_IDU_Lv1 { get; set; }
        public string U_IDU_Lv2 { get; set; }
        public string U_IDU_Lv3 { get; set; }
        public string U_IDU_Lv4 { get; set; }
        public string U_IDU_Lv5 { get; set; }
        public string CntctPrsn { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string Address { get; set; }
        public string GroupName { get; set; }
        public string U_Longitude { get; set; }
        public string U_Latitude { get; set; }
        public string U_PriceGroup { get; set; }
        public string Country { get; set; }
    }
    #endregion
}
