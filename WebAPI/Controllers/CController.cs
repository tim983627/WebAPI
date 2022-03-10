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
    public class CController : ApiController
    {
        public static List<CInventory> DATA = new List<CInventory>();
        //Getamount(抓盤點單)
        public IEnumerable<CInventory> Getamount(int x)
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
                    List<CInventory> amount = new List<CInventory>();
                    CInventory Getdata = new CInventory();
                    Getdata.ItemCode = "連接失敗";
                    amount.Add(Getdata);
                    return amount;
                }
                else
                {
                    CData.DocNum = x;
                    List<CInventory> amount = new List<CInventory>();
                    CInventory Getdata = new CInventory();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T1.[ItemCode], T1.[InWhsQty], T1.[CountQty] FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[DocNum] =" + x);
                    while (oRecordSet.EoF == false)
                    {
                        Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        Getdata.InWhsQty = oRecordSet.Fields.Item("InWhsQty").Value.ToString();
                        Getdata.CountQty = oRecordSet.Fields.Item("CountQty").Value.ToString();
                        CData.ItemCode = Getdata.ItemCode;
                        CData.InWhsQty = Getdata.InWhsQty;
                        CData.CountQty = Getdata.CountQty;
                        oRecordSet.MoveNext();
                        amount.Add(Getdata);
                    }
                    return amount;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }
        //amount (用盤點單裡的商品編號抓盤點資料)
        public IEnumerable<CInventory> Getamount()
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
                    List<CInventory> amount = new List<CInventory>();
                    CInventory Getdata = new CInventory();
                    Getdata.ItemCode = "連接失敗";

                    amount.Add(Getdata);
                    return amount;
                }
                else
                {
                    var DocNum = CData.DocNum;
                    var ItemCode = CData.ItemCode;
                    var InWhsQty = CData.InWhsQty;
                    var CountQty = CData.CountQty;
                    List<CInventory> amount = new List<CInventory>();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T1.[ItemCode], T1.[InWhsQty], T1.[CountQty] FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[DocNum] =" + DocNum);
                    while (oRecordSet.EoF == false)
                    {
                        CInventory Getdata = new CInventory();
                        Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        Getdata.InWhsQty = oRecordSet.Fields.Item("InWhsQty").Value.ToString();
                        Getdata.CountQty = oRecordSet.Fields.Item("CountQty").Value.ToString();
                        oRecordSet.MoveNext();
                        amount.Add(Getdata);
                    }
                    return amount;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        //GetC (把剛剛APP盤點的資料回傳API)
        public string GetC(string ItemCode, string InWhsQty, string CountQty)
        {
            CInventory Postdata = new CInventory();
            Postdata.ItemCode = ItemCode;
            Postdata.InWhsQty = InWhsQty;
            Postdata.CountQty = CountQty;
            DATA.Add(Postdata);
            return "成功";
        }

        //GetDATAClear 將DATA清除 測試會用到 成果不會用到
        public string GetDATAClear()
        {
            DATA.Clear();
            return "清除";
        }


        //GetCtoSir (回傳GetC資料給主管看)
        public IEnumerable<CInventory> GetCtoSir(int x)
        {
            Data.DocNum = x;
            List<CInventory> Sir = new List<CInventory>();
            foreach (var Ic in DATA)
            {
                CInventory GetData = new CInventory();
                GetData.ItemCode = Ic.ItemCode;
                GetData.InWhsQty = Ic.InWhsQty;
                GetData.CountQty = Ic.CountQty;
                Sir.Add(GetData);
            }
            return Sir;

        }

        //GetCount把資料丟到盤點(還沒過帳)
        public IEnumerable<GG> GetCount()
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
                    gg.Fail = "失敗";
                    GGData.Add(gg);
                    return GGData;
                }
                else
                {
                    //把DATA整理，
                    List<CInventory> PostingData = new List<CInventory>();
                    CInventory amount = new CInventory();
                    foreach (var Peko in DATA)
                    {
                        amount.ItemCode = Peko.ItemCode;
                        amount.InWhsQty = Peko.InWhsQty;
                        amount.CountQty = Peko.CountQty;
                        PostingData.Add(amount);
                    }

                    SAPbobsCOM.CompanyService oCS = oCompany.GetCompanyService();
                    SAPbobsCOM.InventoryCountingsService oICS = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
                    SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
                    oICP.DocumentEntry = Data.DocNum;
                    SAPbobsCOM.InventoryCounting oIC = oICS.Get(oICP) as SAPbobsCOM.InventoryCounting;
                    SAPbobsCOM.InventoryCountingLine line = oIC.InventoryCountingLines.Item(0);
                    line.CountedQuantity = Convert.ToDouble(amount.InWhsQty);
                    line.Counted = SAPbobsCOM.BoYesNoEnum.tYES;
                    oICS.Update(oIC);
                    List<GG> GGData = new List<GG>();
                    GG gg = new GG();
                    gg.Success = "成功";
                    GGData.Add(gg);
                    return GGData;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }

        }

        //GetCPosting把GetA的資料丟入SAP進行過帳
        public IEnumerable<GG> GetCPosting()
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
                    gg.Fail = "失敗";
                    GGData.Add(gg);
                    return GGData;
                }
                else
                {
                    //把DATA整理，
                    List<CInventory> PostingData = new List<CInventory>();
                    CInventory amount = new CInventory();
                    foreach (var Peko in DATA)
                    {
                        amount.ItemCode = Peko.ItemCode;
                        amount.InWhsQty = Peko.InWhsQty;
                        amount.CountQty = Peko.CountQty;
                        PostingData.Add(amount);
                    }

                    SAPbobsCOM.CompanyService oCS = oCompany.GetCompanyService();
                    SAPbobsCOM.InventoryCountingsService oICS = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
                    SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
                    oICP.DocumentEntry = CData.DocNum;
                    SAPbobsCOM.InventoryCounting oIC = oICS.Get(oICP) as SAPbobsCOM.InventoryCounting;
                    SAPbobsCOM.InventoryCountingLine line = oIC.InventoryCountingLines.Item(0);
                    line.CountedQuantity = Convert.ToDouble(amount.InWhsQty);
                    line.Counted = SAPbobsCOM.BoYesNoEnum.tYES;
                    oICS.Update(oIC);
                    //過帳
                    if (line.InWarehouseQuantity == Convert.ToDouble(amount.InWhsQty))
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
                        SAPbobsCOM.InventoryPostingSerialNumber oInventoryPostingSerialNumber;
                        foreach (var item in PostingData)
                        {
                            oInventoryPostingSerialNumber = oIPL.InventoryPostingSerialNumbers.Add();
                            oInventoryPostingSerialNumber.InternalSerialNumber = item.InWhsQty;
                        }
                        SAPbobsCOM.InventoryPostingParams oInventoryPostingParams = oIPS.Add(oIP);
                        DATA.Clear();
                        PostingData.Clear();

                    }
                    List<GG> GGData = new List<GG>();
                    GG gg = new GG();
                    gg.Success = "成功";
                    GGData.Add(gg);
                    return GGData;

                }
            }
            catch (Exception errMsg)
            {
                DATA.Clear();
                throw errMsg;
            }

        }
    }
}