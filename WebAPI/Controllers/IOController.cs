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
    
    public class IOController : ApiController
    {
        public static List<GExit> GExitNumber = new List<GExit>();
        //連線狀態
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
                    return GGData;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }
        //收貨
        public string GetImport(string itemcode, string warehouse, int quantity)
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
                        string Ser = oRecordSet.Fields.Item("ManSerNum").Value.ToString();
                        string Btch = oRecordSet.Fields.Item("ManbtchNum").Value.ToString();
                        SAPbobsCOM.Documents oInvGenEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                        if (Ser == "Y" && Btch == "N")
                        {
                            var SN = oInvGenEntry.Lines.SerialNumbers;
                            oInvGenEntry.Lines.ItemCode = itemcode;
                            oInvGenEntry.Lines.Quantity = quantity;
                            oInvGenEntry.Lines.WarehouseCode = warehouse;

                            foreach (var item in GExitNumber)
                            {
                                SN.InternalSerialNumber = item.Number;
                                SN.Add();
                            }

                        }
                        else if (Ser == "N" && Btch == "Y")
                        {
                            var obn = oInvGenEntry.Lines.BatchNumbers;
                            oInvGenEntry.Lines.ItemCode = itemcode;
                            oInvGenEntry.Lines.Quantity = quantity;
                            oInvGenEntry.Lines.WarehouseCode = warehouse;

                            foreach (var item in GExitNumber)
                            {
                                obn.BatchNumber = item.Number;
                                obn.Quantity = item.Quantity;
                            }

                        }
                        else
                        {
                            oInvGenEntry.Lines.ItemCode = itemcode;
                            oInvGenEntry.Lines.Quantity = quantity;
                            oInvGenEntry.Lines.WarehouseCode = warehouse;
                        }
                        oInvGenEntry.Add();
                        GExitNumber.Clear();
                        return "成功";
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }


        //發貨//////////////////////////////////////////////////////////////////////////////////////////////////////
        //透過使用者輸入的ItemCode抓到發貨商品為ABC哪類商品
        public IEnumerable<ABC> GetABC(string itemcode)
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
                    oRecordSet.DoQuery("SELECT T0.[ManbtchNum], T0.[ManSerNum] FROM OITM T0 WHERE T0.[ItemCode] =" + "'" + itemcode + "'");
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

        //抓序號
        public IEnumerable<ANumber> GetANumber(string ItemCode , string WhsCode)
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
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[IntrSerial] FROM OSRI T0 WHERE T0.[Status] =0 and T0.[ItemCode] ="+"'"+ItemCode+"'"+ "and T0.[WhsCode] ="+"'"+WhsCode+"'");
                    while (oRecordSet.EoF == false)
                    {
                        ANumber Getdata = new ANumber();
                        Getdata.ItemCode = ItemCode;
                        Getdata.DistNumber = oRecordSet.Fields.Item("IntrSerial").Value.ToString();
                        Getdata.Whether = "ffalse.png";
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

        //抓批號
        public IEnumerable<BNumber> GetBNumber(string ItemCode, string WhsCode)
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
                    List<BNumber> BNumber = new List<BNumber>();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[BatchNum], T0.[Quantity] FROM OIBT T0 WHERE T0.[Quantity] >0 and  T0.[ItemCode] =" + "'"+ItemCode+"'"+ "and  T0.[WhsCode] ="+"'"+WhsCode+"'");
                    while (oRecordSet.EoF == false)
                    {
                        BNumber Getdata = new BNumber();
                        Getdata.ItemCode = ItemCode;
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

        //抓C類商品
        public IEnumerable<CInventory> GetCNumber(string ItemCode, string WhsCode)
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
                    List<CInventory> CInventory = new List<CInventory>();
                    CInventory Getdata = new CInventory();
                    Getdata.ItemCode = "連接失敗";

                    CInventory.Add(Getdata);
                    return CInventory;
                }
                else
                {
                    List<CInventory> CInventory = new List<CInventory>();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("SELECT T0.[ItemCode], T0.[OnHand] FROM OITW T0 WHERE T0.[ItemCode] ="+"'"+ItemCode+"'"+ "and T0.[WhsCode] ="+"'"+WhsCode+"'");
                    while (oRecordSet.EoF == false)
                    {
                        CInventory Getdata = new CInventory();
                        Getdata.ItemCode = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        Getdata.InWhsQty = oRecordSet.Fields.Item("OnHand").Value.ToString();
                        Getdata.CountQty = "0";
                        oRecordSet.MoveNext();
                        CInventory.Add(Getdata);
                    }
                    return CInventory;
                }
            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        //將APP發貨資料丟回API
        public string GetExit(string N, int Q = 1)
        {
            GExit GExitdata = new GExit();
            GExitdata.Number = N;
            GExitdata.Quantity = Q;
            GExitNumber.Add(GExitdata);
            return GExitdata.Number;
        }

        //將抓回的資料丟到SAP進行發貨
        public string GetGenExit(string itemcode, int quantity, string warehouse)
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

        //GetDATAClear 將DATA清除
        public string GetDATAClear()
        {
            GExitNumber.Clear();
            return "清除";
        }
    }


    
    
}