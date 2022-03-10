using SAPbobsCOM;
using System;
using System.Collections;
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
    public class AController : ApiController
    {
        public static List<ANumber> DATA = new List<ANumber>();

        //GetSirABC (用商品編號判斷主管選擇的盤點單為哪類商品)
        public IEnumerable<ABC> GetSirABC(string x)
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
                    List<ABC> ABC = new List<ABC>();
                    ABC Getdata = new ABC();
                    Getdata.ABCNumber = "連接失敗";
                    ABC.Add(Getdata);

                    return ABC;
                }
                else
                {
                    List<ABC> ABC = new List<ABC>();
                    ABC Getdata = new ABC();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[ManbtchNum], T0.[ManSerNum] FROM OITM T0 WHERE T0.[ItemCode] =" + "'" + x + "'");
                    while (oRecordSet.EoF == false)
                    {
                        string Ser = oRecordSet.Fields.Item("ManSerNum").Value.ToString();
                        string Btch = oRecordSet.Fields.Item("ManbtchNum").Value.ToString();
                        if (Ser == "Y" && Btch == "N")
                        {
                            Getdata.ABCNumber = "A";
                        }
                        else if (Ser == "N" && Btch == "Y")
                        {
                            Getdata.ABCNumber = "B";
                        }
                        else
                        {
                            Getdata.ABCNumber = "C";
                        }
                        oRecordSet.MoveNext();
                        ABC.Add(Getdata);
                    }
                    return ABC;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }
        //已清點未過帳盤點單單號
        public IEnumerable<AInventory> GetCountNToSir()
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
                    List<AInventory> Inventory = new List<AInventory>();
                    AInventory Getdata = new AInventory();
                    Getdata.ItemCode = "連接失敗";
                    Inventory.Add(Getdata);
                    return Inventory;
                }
                else
                {
                    List<AInventory> CountNumber = new List<AInventory>();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("select T0.DocEntry , T1.ItemCode,T1.ItemDesc FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.Status ='O' and T1.Counted='Y'");
                    while (oRecordSet.EoF == false)
                    {
                        AInventory Getdata = new AInventory();
                        Getdata.Entry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
                        Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        Getdata.ItemDesc = oRecordSet.Fields.Item("ItemDesc").Value.ToString();
                        oRecordSet.MoveNext();
                        CountNumber.Add(Getdata);
                    }
                    return CountNumber;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        //將序號等資料送到資料庫
        public IEnumerable<GG> GetSer()
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
                    string SER = "";
                    string Whether = "";
                    foreach (var Peko in DATA)
                    {
                        SER = SER + Peko.DistNumber + ",";
                        if (Peko.Whether == "GreenTick.png")
                        {
                            Whether = Whether + "1" + ",";
                        } 
                        else if (Peko.Whether == "NewTick.png")
                        {
                            Whether = Whether + "N1" + ",";
                        } 
                        else if (Peko.Whether == "RedCross.png")
                        {
                            Whether = Whether + "0" + ",";
                        } 
                    }
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("UPDATE OINC SET U_Ser1 ='"+ SER.TrimEnd(',') + "', U_Whether = '" + Whether.TrimEnd(',') + "'  WHERE DocEntry = " +Data.DocNum);
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
        //GetGG (檢查連接資料庫是否成功)
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

                if (connectionResult != 0)
                {
                    oCompany.GetLastError(out errorCode, out errorMessage);
                    List<GG> GGData = new List<GG>();
                    GG gg = new GG();
                    gg.Success = "連接失敗";
                    GGData.Add(gg);
                    return GGData;
                }
                else
                {
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
        public IEnumerable<AInventory> GetInventory(int x)
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
                    List<AInventory> Inventory = new List<AInventory>();
                    AInventory Getdata = new AInventory();
                    Getdata.ItemCode = "連接失敗";
                    Inventory.Add(Getdata);
                    return Inventory;
                }
                else
                {
                    Data.DocNum = x;
                    List<AInventory> Inventory = new List<AInventory>();
                    AInventory Getdata = new AInventory();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[DocNum], T1.[ItemCode], T1.[ItemDesc], T1.[WhsCode], T0.[Status] FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[DocNum] =" + x);
                    if (oRecordSet.EoF == true)
                    {
                        Data.ItemCode = "Error";
                        Inventory.Add(Getdata);
                    }
                    else
                    {
                        while (oRecordSet.EoF == false)
                        {
                            var status = oRecordSet.Fields.Item("Status").Value.ToString();
                            if (status == "O")
                            {
                                Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                                Getdata.ItemDesc = oRecordSet.Fields.Item("ItemDesc").Value.ToString();
                                Getdata.WhsCode = oRecordSet.Fields.Item("WhsCode").Value.ToString();
                                Data.ItemCode = Getdata.ItemCode;
                                Data.WhsCode = Getdata.WhsCode;
                                oRecordSet.MoveNext();
                                Inventory.Add(Getdata);
                            }
                            else if (status == "C")
                            {
                                oRecordSet.MoveNext();
                                Data.ItemCode = "Close";
                            }  
                        }
                    }
                    return Inventory;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }
        //GetABC (用商品編號去查此商品為ABC哪類)
        public IEnumerable<ABC> GetABC()
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
                    List<ABC> ABC = new List<ABC>();
                    ABC Getdata = new ABC();
                    Getdata.ABCNumber = "連接失敗";
                    ABC.Add(Getdata);

                    return ABC;
                }
                else
                {
                    var ItemCode = Data.ItemCode;
                    List<ABC> ABC = new List<ABC>();
                    ABC Getdata = new ABC();
                    if (ItemCode == "Close")
                    {
                        Getdata.ABCNumber = "Close";
                        ABC.Add(Getdata);
                    }
                    else
                    {
                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery("SELECT T0.[ManbtchNum], T0.[ManSerNum] FROM OITM T0 WHERE T0.[ItemCode] =" + "'" + ItemCode + "'");
                        if (ItemCode == "Error")
                        {
                            Getdata.ABCNumber = "Error";
                            ABC.Add(Getdata);
                        }
                        else
                        {
                            while (oRecordSet.EoF == false)
                            {
                                string Ser = oRecordSet.Fields.Item("ManSerNum").Value.ToString();
                                string Btch = oRecordSet.Fields.Item("ManbtchNum").Value.ToString();
                                if (Ser == "Y" && Btch == "N")
                                {
                                    Getdata.ABCNumber = "A";
                                }
                                else if (Ser == "N" && Btch == "Y")
                                {
                                    Getdata.ABCNumber = "B";
                                }
                                else
                                {
                                    Getdata.ABCNumber = "C";
                                }
                                oRecordSet.MoveNext();
                                ABC.Add(Getdata);
                            }
                        }
                    }
                    return ABC;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        //ANumber (用盤點單裡的商品編號抓序號)
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
                    var ItemCode = Data.ItemCode;
                    var WhsCode = Data.WhsCode;
                    List<ANumber> ANumber = new List<ANumber>();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT OSRI.[ItemCode],OSRI.[IntrSerial] from OSRI LEFT JOIN OITM on OSRI.[ItemCode] = OITM.[ItemCode] WHERE OSRI.[Status] = 0 and OSRI.[ItemCode] = '"+ItemCode +"'and OSRI.[WhsCode]='" + WhsCode+"'");
                    while (oRecordSet.EoF == false)
                    {
                        ANumber Getdata = new ANumber();
                        Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        Getdata.DistNumber = oRecordSet.Fields.Item("IntrSerial").Value.ToString();
                        Getdata.Whether = "RedCross.png";
                        oRecordSet.MoveNext();
                        ANumber.Add(Getdata);
                    }
                    return ANumber;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        //GetA (把剛剛APP盤點的資料回傳API)
        public string GetA(string ItemCode, string DistNumber, string Whether)
        {
            ANumber Postdata = new ANumber();
            Postdata.ItemCode = ItemCode;
            Postdata.DistNumber = DistNumber;
            Postdata.Whether = Whether;
            DATA.Add(Postdata); 
            return Postdata.DistNumber;
        }

        //GetAtoSir (回傳GetA資料給主管看)
        public IEnumerable<ANumber> GetAtoSir(int x)
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
                    List<ANumber> Sir = new List<ANumber>();
                    ANumber Getdata = new ANumber();
                    Getdata.ItemCode = "連接失敗";
                    Sir.Add(Getdata);
                    return Sir;
                }
                else
                {     
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("select T1.ItemCode,T0.U_Ser1,T0.U_Whether FROM OINC T0  INNER JOIN INC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[DocEntry] =" + x);
                    string ItemCode= oRecordSet.Fields.Item("ItemCode").Value.ToString();
                    string Split = oRecordSet.Fields.Item("U_Ser1").Value.ToString();
                    string[] SERDATA = Split.Split(',');
                    Split = oRecordSet.Fields.Item("U_Whether").Value.ToString();
                    string[] WhetherDATA = Split.Split(',');

                    List<ANumber> Sir = new List<ANumber>();

                    for (var i = 0; i < SERDATA.Length; i++)
                    {
                        ANumber GetData = new ANumber();
                        GetData.ItemCode = ItemCode;
                        GetData.DistNumber = SERDATA[i];
                        if (WhetherDATA[i] == "1")
                        {
                            GetData.Whether = "GreenTick.png";
                        } 
                        else if (WhetherDATA[i] == "N1")
                        {
                            GetData.Whether = "NewTick.png";
                        }
                        else if (WhetherDATA[i] == "0")
                        {
                            GetData.Whether = "RedCross.png";
                        }

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

        //GetDATAClear 將DATA清除
        public string GetDATAClear()
        {
            DATA.Clear();
            return "清除";
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
                    //把DATA整理，只要把GreenTick的丟回即可
                    int i = 0;
                    foreach (var Peko in DATA.Where(w => w.Whether == "GreenTick.png" || w.Whether == "NewTick.png"))
                    {
                        i++;
                    }

                    SAPbobsCOM.CompanyService oCS = oCompany.GetCompanyService();
                    SAPbobsCOM.InventoryCountingsService oICS = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
                    SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
                    oICP.DocumentEntry = Data.DocNum;
                    SAPbobsCOM.InventoryCounting oIC = oICS.Get(oICP) as SAPbobsCOM.InventoryCounting;
                    SAPbobsCOM.InventoryCountingLine line = oIC.InventoryCountingLines.Item(0);
                    line.CountedQuantity = i;
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

        //GetAPosting把GetA的資料丟入SAP進行過帳
        public IEnumerable<GG> GetAPosting()
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
                    //把DATA整理，只要把GreenTick的丟回即可
                    List<ANumber> PostingData = new List<ANumber>();

                    foreach (var Peko in DATA.Where(w => w.Whether == "GreenTick.png" || w.Whether == "NewTick.png"))
                    {
                        ANumber ANumber = new ANumber();
                        ANumber.ItemCode = Peko.ItemCode;
                        ANumber.DistNumber = Peko.DistNumber;
                        ANumber.Whether = Peko.Whether;
                        PostingData.Add(ANumber);
                    }

                    SAPbobsCOM.CompanyService oCS = oCompany.GetCompanyService();
                    SAPbobsCOM.InventoryCountingsService oICS = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
                    SAPbobsCOM.InventoryCountingParams oICP = (SAPbobsCOM.InventoryCountingParams)oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);
                    oICP.DocumentEntry = Data.DocNum;
                    SAPbobsCOM.InventoryCounting oIC = oICS.Get(oICP) as SAPbobsCOM.InventoryCounting;
                    SAPbobsCOM.InventoryCountingLine line = oIC.InventoryCountingLines.Item(0);
                    line.CountedQuantity = PostingData.Count();
                    line.Counted = SAPbobsCOM.BoYesNoEnum.tYES;
                    oICS.Update(oIC);
                    //過帳
                    if (line.InWarehouseQuantity == PostingData.Count)
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
                            oInventoryPostingSerialNumber.InternalSerialNumber = item.DistNumber;
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