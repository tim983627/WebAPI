using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WebAPI.Models;

namespace WebAPI.Controllers
{
    public class TESTController : ApiController
    {


        public static string ItemCode, WhsCode;
        public static int number;
        public static List<ANumber> ANumber = new List<ANumber>();
        public static List<GExit> GExitNumber = new List<GExit>();
        //GetGG 開始(檢查連接資料庫是否成功)
        public IEnumerable<GG> GetGGs()
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
                //盤點

                //過帳

                
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
                    List<GG> GGData = new List<GG>();
                    GG gg = new GG();
                    gg.Success = "連接成功";
                    GGData.Add(gg);

                    SAPbobsCOM.Documents oInvGenExit = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    var SN = oInvGenExit.Lines.SerialNumbers;
                    var obn = oInvGenExit.Lines.BatchNumbers;
                    oInvGenExit.Lines.ItemCode = "B10000";
                    oInvGenExit.Lines.Quantity = 8;
                    oInvGenExit.Lines.WarehouseCode = "A00商品";
                    obn.BatchNumber = "B-B1234";
                    obn.Quantity = 3;
                    obn.Add();
                    obn.BatchNumber = "BTEST01"; 
                    obn.Quantity = 5;
                    
                    oInvGenExit.Add();




                    return GGData;
                }
            }    
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }
        //GetGG 結束(檢查連接資料庫是否成功)
        //GetInventory開始(抓盤點單)
        public IEnumerable<Inventory> GetInventory(int x)
        {
            
            int errorCode = 0;
            string errorMessage = "";
            number = x;
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
                    List<Inventory> Inventory = new List<Inventory>();
                    Inventory Getdata = new Inventory();
                    Getdata.ItemCode = "連接失敗";

                    Inventory.Add(Getdata);
                    return Inventory;
                }
                else
                {
                    List<Inventory> Inventory = new List<Inventory>();
                    Inventory Getdata = new Inventory();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[DocNum], T1.[ItemCode], T1.[ItemDesc], T1.[WhsCode] FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[DocNum] =" + x);
                    while (oRecordSet.EoF == false)
                    {
                        
                        Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        Getdata.ItemDesc = oRecordSet.Fields.Item("ItemDesc").Value.ToString();
                        Getdata.WhsCode = oRecordSet.Fields.Item("WhsCode").Value.ToString();
                        ItemCode = Getdata.ItemCode;
                        WhsCode = Getdata.WhsCode;
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
        //GetInventory 結束(抓盤點單)
        //SELECT T1.[ItemCode], T0.[DistNumber] FROM OSRN T0  INNER JOIN OSRQ T1 ON T0.[AbsEntry] = T1.[MdAbsEntry] WHERE T0.[ItemCode] = 'A001' and T1.[WhsCode] = 'A00商品' and T1.[Quantity] = 1
        //ANumber 開始(用盤點單裡的商品編號抓序號)

        public IEnumerable<ANumber> GetANumber()
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
                    List<ANumber> ANumber = new List<ANumber>();
                    ANumber Getdata = new ANumber();
                    Getdata.ItemCode = "連接失敗";

                    ANumber.Add(Getdata);
                    return ANumber;
                }
                else
                {
                    List<ANumber> ANumber = new List<ANumber>();


                    if (ItemCode.Substring(0,1) == "A")
                    {
                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery("SELECT T1.[ItemCode], T0.[DistNumber] FROM OSRN T0  INNER JOIN OSRQ T1 ON T0.[AbsEntry] = T1.[MdAbsEntry] WHERE T0.[ItemCode] ='" + ItemCode + "'and T1.[WhsCode] ='" + WhsCode + "'and T1.[Quantity] = 1");
                        while (oRecordSet.EoF == false)
                        {
                            ANumber Getdata = new ANumber();
                            Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                            Getdata.DistNumber = oRecordSet.Fields.Item("DistNumber").Value.ToString();
                            Getdata.Quantity = "未清點";
                            oRecordSet.MoveNext();
                            ANumber.Add(Getdata);
                        }
                    }
                    else if (ItemCode.Substring(0,1) == "B")
                    {
                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery("SELECT T1.[ItemCode], T0.[DistNumber] ,T1.[Quantity] FROM OBTN T0  INNER JOIN OBTQ T1 ON T0.[AbsEntry] = T1.[MdAbsEntry] WHERE T0.[ItemCode] ='" + ItemCode + "'and T1.[WhsCode] ='" + WhsCode + "'");
                        while (oRecordSet.EoF == false)
                        {
                            ANumber Getdata = new ANumber();
                            Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                            Getdata.DistNumber = oRecordSet.Fields.Item("DistNumber").Value.ToString();
                            Getdata.Quantity = "0/" + oRecordSet.Fields.Item("Quantity").Value.ToString();
                            oRecordSet.MoveNext();
                            ANumber.Add(Getdata);
                        }
                    }
                    else 
                    {
                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery("SELECT T1.[ItemCode],T1.[ItemDesc], T1.[InWhsQty]  FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[DocNum] =" + number);
                        while (oRecordSet.EoF == false)
                        {
                            ANumber Getdata = new ANumber();
                            Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                            Getdata.DistNumber = oRecordSet.Fields.Item("ItemDesc").Value.ToString();
                            Getdata.Quantity = "0/" + oRecordSet.Fields.Item("InWhsQty").Value.ToString();
                            oRecordSet.MoveNext();
                            ANumber.Add(Getdata);
                        }                        
                    }
                    return ANumber;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }
        //ANumber 結束(用盤點單裡的商品編號抓序號)


        public string GET(string I, string D, string Q)
        {
    
                ANumber Postdata = new ANumber();
                Postdata.ItemCode = I;
                Postdata.DistNumber = D;
                Postdata.Quantity = Q;
                ANumber.Add(Postdata);
   
            return Postdata.ItemCode;

        }

        public string GExit(string N, int Q)
        {

            GExit GExitdata = new GExit();
            GExitdata.Number = N;
            GExitdata.Quantity = Q;
            GExitNumber.Add(GExitdata);

            return GExitdata.Number;

        }

        public int GETT(int Q=1)
        {

           

            return Q;

        }


        public string GetGenEntry(string itemcode,int quantity,string warehouse)
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

                    return "連接失敗";
                }
                else
                {
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[ManbtchNum], T0.[ManSerNum] FROM OITM T0 WHERE T0.[ItemCode] =" + "'" + itemcode + "'");
                    while (oRecordSet.EoF == false)
                    {
                        string Ser = oRecordSet.Fields.Item("ManSerNum").Value.ToString();
                        string Btch = oRecordSet.Fields.Item("ManbtchNum").Value.ToString();
                        int i = 1,j=1;
                        SAPbobsCOM.Documents oInvGenEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                        if (Ser == "Y" && Btch == "N")
                        {
                            var SN = oInvGenEntry.Lines.SerialNumbers;                            
                            oInvGenEntry.Lines.ItemCode = itemcode;
                            oInvGenEntry.Lines.Quantity = quantity;
                            oInvGenEntry.Lines.UnitPrice = 1;
                            oInvGenEntry.Lines.WarehouseCode = warehouse;
                            while ( i<= quantity)
                            {
                                oRecordSet.DoQuery("select T0.[IntrSerial]  from OSRI T0 where T0.[Status]=0 and T0.[IntrSerial]=" + "'" +itemcode + j.ToString("D2") + "'");
                                string TEST = oRecordSet.Fields.Item("IntrSerial").Value.ToString();
                                if (TEST == "") 
                                {
                                    SN.InternalSerialNumber = itemcode+ j.ToString("D2");
                                    SN.Add();
                                    i++;
                                    j++;
                                }
                                else 
                                {
                                    j++;
                                }

                            }
                            
                        }
                        else if (Ser == "N" && Btch == "Y")
                        {
                            var obn = oInvGenEntry.Lines.BatchNumbers;
                            oInvGenEntry.Lines.ItemCode = itemcode;
                            oInvGenEntry.Lines.Quantity = quantity;
                            oInvGenEntry.Lines.UnitPrice = 1;
                            oInvGenEntry.Lines.WarehouseCode = warehouse;
                            oRecordSet.DoQuery("select T0.[BatchNum] from OIBT T0 where Quantity=0 and T0.[BatchNum]=" + "'" + itemcode + j.ToString("D2") + "'");
                            string TEST = oRecordSet.Fields.Item("BatchNum").Value.ToString();
                            while (true)
                            {
                                if (TEST == "")
                                {
                                    obn.BatchNumber = itemcode + j.ToString("D2");
                                    obn.Quantity = quantity;                                    
                                    break;
                                }
                                else
                                {
                                    j++;
                                }
                            }
                        }
                        else
                        {
                            oInvGenEntry.Lines.ItemCode = itemcode;
                            oInvGenEntry.Lines.Quantity = quantity;
                            oInvGenEntry.Lines.UnitPrice = 1;
                            oInvGenEntry.Lines.WarehouseCode = warehouse;
                        }
                        oInvGenEntry.Add();
                        
                    }
                    return"成功";
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        public string GetGenExit(string itemcode, int quantity , int price, string warehouse)
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

                    return "連接失敗";
                }
                else
                {
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[ManbtchNum], T0.[ManSerNum] FROM OITM T0 WHERE T0.[ItemCode] =" + "'" + itemcode + "'");
                    while (oRecordSet.EoF == false)
                    {
                        string Ser = oRecordSet.Fields.Item("ManSerNum").Value.ToString();
                        string Btch = oRecordSet.Fields.Item("ManbtchNum").Value.ToString();
                        SAPbobsCOM.Documents oInvGenExit = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                        if (Ser == "Y" && Btch == "N")
                        {
                            var SN = oInvGenExit.Lines.SerialNumbers;
                            oInvGenExit.Lines.ItemCode = itemcode;
                            oInvGenExit.Lines.Quantity = quantity;
                            oInvGenExit.Lines.UnitPrice = price;
                            oInvGenExit.Lines.WarehouseCode = warehouse;                            
                            foreach (var item in GExitNumber)
                            {
                                SN.InternalSerialNumber = item.Number;
                                SN.Add();
                            }
                        }
                        else if (Ser == "N" && Btch == "Y")
                        {
                            var obn = oInvGenExit.Lines.BatchNumbers;
                            oInvGenExit.Lines.ItemCode = itemcode;
                            oInvGenExit.Lines.Quantity = quantity;
                            oInvGenExit.Lines.UnitPrice = price;
                            oInvGenExit.Lines.WarehouseCode = warehouse;
                            foreach (var item in GExitNumber)
                            {
                                obn.BatchNumber = item.Number;
                                obn.Quantity = item.Quantity;
                                obn.Add();
                            }
                            
                        }
                        else
                        {
                            oInvGenExit.Lines.ItemCode = itemcode;
                            oInvGenExit.Lines.Quantity = quantity;
                            oInvGenExit.Lines.UnitPrice = price;
                            oInvGenExit.Lines.WarehouseCode = warehouse;
                        }
                        oInvGenExit.Add();
                        GExitNumber.Clear();
                        oRecordSet.MoveNext();

                    }
                    return "成功";
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        /*
    //批次
    //盤點
        SAPbobsCOM.CompanyService oCS = oCompany.GetCompanyService();
        SAPbobsCOM.InventoryCountingsService oICS = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
        SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
        oICP.DocumentEntry = 盤點單單號;
        SAPbobsCOM.InventoryCounting oIC = oICS.Get(oICP) as SAPbobsCOM.InventoryCounting;
        SAPbobsCOM.InventoryCountingLine line = oIC.InventoryCountingLines.Item(0);
        line.CountedQuantity = 已清點數量;
        line.Counted = SAPbobsCOM.BoYesNoEnum.tYES;
        oICS.Update(oIC);   
        oICS.Close(oICP);//關閉
    //過帳
        SAPbobsCOM.InventoryPostingsService oIPS = oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryPostingsService);
        SAPbobsCOM.InventoryPosting oIP = oIPS.GetDataInterface(SAPbobsCOM.InventoryPostingsServiceDataInterfaces.ipsInventoryPosting);
        oIP.CountDate = DateTime.Now;
        SAPbobsCOM.InventoryPostingLines oIPLS = oIP.InventoryPostingLines;
        SAPbobsCOM.InventoryPostingLine oIPL = oIPLS.Add();
        oIPL.BaseEntry = 盤點單單號;
        oIPL.BaseLine = 1;
        SAPbobsCOM.InventoryPostingBatchNumber oInventoryPostingBatchNumber = oIPL.InventoryPostingBatchNumbers.Add();
        oInventoryPostingBatchNumber.BatchNumber = 批次號碼;
        oInventoryPostingBatchNumber.Quantity = 盤點清點數量;
        SAPbobsCOM.InventoryPostingParams oInventoryPostingParams = oIPS.Add(oIP);
         */


        /*
    //序號
    //盤點
        SAPbobsCOM.CompanyService oCS = oCompany.GetCompanyService();
        SAPbobsCOM.InventoryCountingsService oICS = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
        SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
        oICP.DocumentEntry = 盤點單單號;
        SAPbobsCOM.InventoryCounting oIC = oICS.Get(oICP) as SAPbobsCOM.InventoryCounting;
        SAPbobsCOM.InventoryCountingLine line = oIC.InventoryCountingLines.Item(0);
        line.CountedQuantity = 已清點數量;
        line.Counted = SAPbobsCOM.BoYesNoEnum.tYES;
        oICS.Update(oIC);   
        oICS.Close(oICP);//關閉
    //過帳
        SAPbobsCOM.InventoryPostingsService oIPS = oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryPostingsService);
        SAPbobsCOM.InventoryPosting oIP = oIPS.GetDataInterface(SAPbobsCOM.InventoryPostingsServiceDataInterfaces.ipsInventoryPosting);
        oIP.CountDate = DateTime.Now;
        SAPbobsCOM.InventoryPostingLines oIPLS = oIP.InventoryPostingLines;
        SAPbobsCOM.InventoryPostingLine oIPL = oIPLS.Add();
        oIPL.BaseEntry = oICP.DocumentEntry;
        oIPL.BaseLine = 1;
        SAPbobsCOM.InventoryPostingSerialNumber oInventoryPostingSerialNumber = oIPL.InventoryPostingSerialNumbers.Add();
        oInventoryPostingSerialNumber.InternalSerialNumber = "X08";
        oInventoryPostingSerialNumber = oIPL.InventoryPostingSerialNumbers.Add();
        oInventoryPostingSerialNumber.InternalSerialNumber = "T001";
        oInventoryPostingSerialNumber = oIPL.InventoryPostingSerialNumbers.Add();
        oInventoryPostingSerialNumber.InternalSerialNumber = "T002";
        oInventoryPostingSerialNumber = oIPL.InventoryPostingSerialNumbers.Add();
        oInventoryPostingSerialNumber.InternalSerialNumber = "T004";
        oInventoryPostingSerialNumber = oIPL.InventoryPostingSerialNumbers.Add();
        oInventoryPostingSerialNumber.InternalSerialNumber = "T005";
        SAPbobsCOM.InventoryPostingParams oInventoryPostingParams = oIPS.Add(oIP);
       */

        //收貨
        /*
        SAPbobsCOM.Documents oInvGenEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
        var SN = oInvGenEntry.Lines.SerialNumbers;
        var obn = oInvGenEntry.Lines.BatchNumbers;
        oInvGenEntry.Lines.ItemCode = "T01";
        oInvGenEntry.Lines.Quantity = 2;
        oInvGenEntry.Lines.UnitPrice = 5;
        oInvGenEntry.Lines.WarehouseCode = "A01半成品";
        SN.InternalSerialNumber = "TEST01";
        SN.Add();
        SN.InternalSerialNumber = "TEST02";
        oInvGenEntry.Add();
         */

        //發貨
        /*
        SAPbobsCOM.Documents oInvGenExit = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
        var SN = oInvGenExit.Lines.SerialNumbers;
        var obn = oInvGenExit.Lines.BatchNumbers;
        oInvGenExit.Lines.ItemCode = "T01";
        oInvGenExit.Lines.Quantity = 2;
        oInvGenExit.Lines.WarehouseCode = "A01半成品";
        SN.InternalSerialNumber = "TEST01";
        SN.Add();
        SN.InternalSerialNumber = "TEST02";
        oInvGenExit.Add();
         */
    }
}