using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WebAPI.Models;

namespace webAPI.Controllers
{
    public class BController : ApiController
    {
        // GET api/batch
        public class data
        {
            public string itemcode;
            public string whs;
            public string batchnumber;
            public string quantity;
            public int count;
        }
        public static int total;
        public static List<data> itembatch = new List<data>();
        public static List<data> DATA = new List<data>();


        //將批號等資料送到資料庫
        public IEnumerable<GG> GetBat()
        {
            int errorCode = 0;
            string errorMessage = "";
            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            try
            {
                oCompany.CompanyDB = "cyutdb";
                oCompany.Server = "WIN-IT8HPSMKSJR";
                oCompany.LicenseServer = "WIN-IT8HPSMKSJR";
                oCompany.DbUserName = "sa";
                oCompany.DbPassword = "2ixijklM";
                oCompany.UserName = "manager";
                oCompany.Password = "1234";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                oCompany.UseTrusted = false;
                int connectionResult = oCompany.Connect();

                if (connectionResult != 0)
                {
                    oCompany.GetLastError(out errorCode, out errorMessage);
                    List<GG> GGData = new List<GG>();
                    GG gg = new GG();
                    gg.Fail = "連接失敗";
                    GGData.Add(gg);
                    return GGData;
                }
                else
                {
                    string BAT = "";
                    string CountQ = "";
                    string Quantity = "";
                    foreach (var Peko in DATA)
                    {
                        BAT = BAT + Peko.batchnumber + ",";
                        CountQ = CountQ + Convert.ToString(Peko.count)+ ",";
                        Quantity = Quantity + Peko.quantity + ",";
                    }
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("UPDATE OINC SET U_Bat1 ='" + BAT.TrimEnd(',') + "', U_CountQ = '" + CountQ.TrimEnd(',') + "', U_Quantity = '" + Quantity.TrimEnd(',') + "'  WHERE DocEntry = " + Data.DocNum);
                    List<GG> GGData = new List<GG>();
                    GG gg = new GG();
                    gg.Success = "連接成功";
                    GGData.Add(gg);
                    return GGData;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }
        //GetInventory(抓盤點單)
        public IEnumerable<BInventory> GetInventory(int x)
        {

            int errorCode = 0;
            string errorMessage = "";
            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            try
            {
                oCompany.CompanyDB = "cyutdb";
                oCompany.Server = "WIN-IT8HPSMKSJR";
                oCompany.LicenseServer = "WIN-IT8HPSMKSJR";
                oCompany.DbUserName = "sa";
                oCompany.DbPassword = "2ixijklM";
                oCompany.UserName = "manager";
                oCompany.Password = "1234";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                oCompany.UseTrusted = false;
                int connectionResult = oCompany.Connect();

                if (connectionResult != 0)
                {
                    oCompany.GetLastError(out errorCode, out errorMessage);
                    List<BInventory> Inventory = new List<BInventory>();
                    BInventory Getdata = new BInventory();
                    Getdata.ItemCode = "連接失敗";
                    Inventory.Add(Getdata);
                    return Inventory;
                }
                else
                {
                    Data.DocNum = x;
                    List<BInventory> Inventory = new List<BInventory>();
                    BInventory Getdata = new BInventory();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T1.[ItemCode], T1.[WhsCode], T1.[ItemDesc] FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[DocNum] =" + x);
                    while (oRecordSet.EoF == false)
                    {
                        Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        Getdata.ItemDesc = oRecordSet.Fields.Item("ItemDesc").Value.ToString();
                        Getdata.WhsCode = oRecordSet.Fields.Item("WhsCode").Value.ToString();
                        Data.ItemCode = Getdata.ItemCode;
                        Data.WhsCode = Getdata.WhsCode;
                        oRecordSet.MoveNext();
                        Inventory.Add(Getdata);
                    }
                    return Inventory;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        //BNumber (用盤點單裡的商品編號抓序號)
        public IEnumerable<BNumber> GetBNumber()
        {

            int errorCode = 0;
            string errorMessage = "";
            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            try
            {
                oCompany.CompanyDB = "cyutdb";
                oCompany.Server = "WIN-IT8HPSMKSJR";
                oCompany.LicenseServer = "WIN-IT8HPSMKSJR";
                oCompany.DbUserName = "sa";
                oCompany.DbPassword = "2ixijklM";
                oCompany.UserName = "manager";
                oCompany.Password = "1234";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                oCompany.UseTrusted = false;
                int connectionResult = oCompany.Connect();

                if (connectionResult != 0)
                {
                    oCompany.GetLastError(out errorCode, out errorMessage);
                    List<BNumber> BNumber = new List<BNumber>();
                    BNumber Getdata = new BNumber();
                    Getdata.ItemCode = "連接失敗";

                    BNumber.Add(Getdata);
                    return BNumber;
                }
                else
                {
                    var ItemCode = Data.ItemCode;
                    var WhsCode = Data.WhsCode;
                    List<BNumber> BNumber = new List<BNumber>();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[ItemCode], T0.[BatchNum], T0.[Quantity] FROM  OIBT T0 WHERE T0.[ItemCode] =" + "'" + ItemCode + "'" + "and T0.[WhsCode]='" + WhsCode + "'");
                    while (oRecordSet.EoF == false)
                    {
                        BNumber Getdata = new BNumber();
                        Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        Getdata.BatchNumber = oRecordSet.Fields.Item("BatchNum").Value.ToString();
                        Getdata.Quantity = oRecordSet.Fields.Item("Quantity").Value.ToString();
                        Getdata.Count = 0;
                        oRecordSet.MoveNext();
                        BNumber.Add(Getdata);
                    }
                    return BNumber;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        //獲取盤點完的資料
        public IEnumerable<data> GetB(string ItemCode, string BatchNumber, int Count, string Quantity)
        {

            data Postdata = new data();
            Postdata.itemcode = ItemCode;
            Postdata.batchnumber = BatchNumber;
            Postdata.count = Count;
            Postdata.quantity = Quantity;
            DATA.Add(Postdata);
            return DATA;

        }

        //GetDATAClear 將DATA清除 測試會用到 成果不會用到
        public string GetDATAClear()
        {
            DATA.Clear();
            return "清除";
        }

        //將盤點完的資料丟入盤點單(未結)
        public IEnumerable<data> GetCount()
        {

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            oCompany.CompanyDB = "cyutdb";
            oCompany.Server = "WIN-IT8HPSMKSJR";
            oCompany.LicenseServer = "WIN-IT8HPSMKSJR";
            oCompany.DbUserName = "sa";
            oCompany.DbPassword = "2ixijklM";
            oCompany.UserName = "manager";
            oCompany.Password = "1234";
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
            oCompany.UseTrusted = false;
            int connectionResult = oCompany.Connect();
            if (connectionResult != 0)
            {
                itembatch.Add(new data() { itemcode = "error connecting db", whs = "error connecting db", batchnumber = "", quantity = "" });
                return itembatch;
            }
            else
            {
                total = 0;
                foreach (var Peko in DATA)
                {
                    total = total + Peko.count;
                }

                SAPbobsCOM.CompanyService oCS = oCompany.GetCompanyService();
                SAPbobsCOM.InventoryCountingsService oICS = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
                SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
                oICP.DocumentEntry = Data.DocNum;
                SAPbobsCOM.InventoryCounting oIC = oICS.Get(oICP) as SAPbobsCOM.InventoryCounting;
                SAPbobsCOM.InventoryCountingLine line = oIC.InventoryCountingLines.Item(0);
                line.CountedQuantity = total;
                line.Counted = SAPbobsCOM.BoYesNoEnum.tYES;
                oICS.Update(oIC);

                itembatch.Add(new data() { itemcode = "true" });
                return itembatch;
            }
        }

        //盤點資料丟到主管過帳
        public IEnumerable<BNumber> GetBtoSir(int x)
        {
            Data.DocNum = x;
            int errorCode = 0;
            string errorMessage = "";
            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            try
            {
                oCompany.CompanyDB = "cyutdb";
                oCompany.Server = "WIN-IT8HPSMKSJR";
                oCompany.LicenseServer = "WIN-IT8HPSMKSJR";
                oCompany.DbUserName = "sa";
                oCompany.DbPassword = "2ixijklM";
                oCompany.UserName = "manager";
                oCompany.Password = "1234";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                oCompany.UseTrusted = false;
                int connectionResult = oCompany.Connect();

                if (connectionResult != 0)
                {
                    oCompany.GetLastError(out errorCode, out errorMessage);
                    List<BNumber> Sir = new List<BNumber>();
                    BNumber Getdata = new BNumber();
                    Getdata.ItemCode = "連接失敗";
                    Sir.Add(Getdata);
                    return Sir;
                }
                else
                {
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("select T1.ItemCode,T0.U_Bat1,T0.U_CountQ,U_Quantity FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[DocEntry] =" + x);
                    string ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                    string Split = oRecordSet.Fields.Item("U_Bat1").Value.ToString();
                    string[] BATDATA = Split.Split(',');
                    Split = oRecordSet.Fields.Item("U_CountQ").Value.ToString();
                    string[] CountQDATA = Split.Split(',');
                    Split = oRecordSet.Fields.Item("U_Quantity").Value.ToString();
                    string[] QuantityDATA = Split.Split(',');

                    List<BNumber> Sir = new List<BNumber>();

                    for (var i = 0; i < BATDATA.Length; i++)
                    {
                        BNumber GetData = new BNumber();
                        GetData.ItemCode = ItemCode;
                        GetData.BatchNumber = BATDATA[i];
                        GetData.Count = Convert.ToInt32(CountQDATA[i]);
                        GetData.Quantity = QuantityDATA[i];
                        Sir.Add(GetData);
                    }
                    return Sir;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }


        }

        //過帳
        public IEnumerable<data> GetBPosting()
        {

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            oCompany.CompanyDB = "cyutdb";
            oCompany.Server = "WIN-IT8HPSMKSJR";
            oCompany.LicenseServer = "WIN-IT8HPSMKSJR";
            oCompany.DbUserName = "sa";
            oCompany.DbPassword = "2ixijklM";
            oCompany.UserName = "manager";
            oCompany.Password = "1234";
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
            oCompany.UseTrusted = false;
            int connectionResult = oCompany.Connect();
            if (connectionResult != 0)
            {
                itembatch.Add(new data() { itemcode = "error connecting db", whs = "error connecting db", batchnumber = "", quantity = "" });
                return itembatch;
            }
            else
            {
                SAPbobsCOM.CompanyService oCS = oCompany.GetCompanyService();
                SAPbobsCOM.InventoryCountingsService oICS = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
                SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
                oICP.DocumentEntry = Data.DocNum;
                SAPbobsCOM.InventoryCounting oIC = oICS.Get(oICP) as SAPbobsCOM.InventoryCounting;
                SAPbobsCOM.InventoryCountingLine line = oIC.InventoryCountingLines.Item(0);
                if (line.InWarehouseQuantity == total)
                {
                    oICS.Close(oICP);
                }
                else
                {
                    SAPbobsCOM.InventoryPostingsService oIPS = oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryPostingsService);
                    SAPbobsCOM.InventoryPosting oIP = oIPS.GetDataInterface(SAPbobsCOM.InventoryPostingsServiceDataInterfaces.ipsInventoryPosting);
                    oIP.CountDate = DateTime.Now;
                    SAPbobsCOM.InventoryPostingLines oIPLS = oIP.InventoryPostingLines;
                    SAPbobsCOM.InventoryPostingLine oIPL = oIPLS.Add();
                    oIPL.BaseEntry = oICP.DocumentEntry;
                    oIPL.BaseLine = 1;
                    SAPbobsCOM.IInventoryPostingBatchNumber oInventoryPostingBatchNumber;
                    foreach (var item in DATA)
                    {
                        if (item.count != 0)
                        {
                            oInventoryPostingBatchNumber = oIPL.InventoryPostingBatchNumbers.Add();
                            oInventoryPostingBatchNumber.BatchNumber = item.batchnumber;
                            oInventoryPostingBatchNumber.Quantity = item.count;
                        }
                        else
                        {
                            oInventoryPostingBatchNumber = oIPL.InventoryPostingBatchNumbers.Add();
                            oInventoryPostingBatchNumber.BatchNumber = item.batchnumber;
                            oInventoryPostingBatchNumber.Quantity = 0;
                        }
                    }
                    SAPbobsCOM.InventoryPostingParams oInventoryPostingParams = oIPS.Add(oIP);

                }
                itembatch.Add(new data() { itemcode = "true" });
                return itembatch;
            }
        }
    }


}



