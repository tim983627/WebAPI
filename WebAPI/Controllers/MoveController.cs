using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WebAPI.Models;

namespace WebAPI.Controllers
{
    public class MoveController : ApiController
    {
        //透過使用者輸入的ItemCode來判斷商品為ABC哪類商品
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

        public static List<GExit> Transfer = new List<GExit>();


        //抓序號
        public IEnumerable<ANumber> GetANumber(string ItemCode, string WhsCode)
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
                    oRecordSet.DoQuery("SELECT T0.[IntrSerial] FROM OSRI T0 WHERE T0.[Status] =0 and T0.[ItemCode] =" + "'" + ItemCode + "'" + "and T0.[WhsCode] =" + "'" + WhsCode + "'");
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
                    oRecordSet.DoQuery("SELECT T0.[BatchNum], T0.[Quantity] FROM OIBT T0 WHERE T0.[Quantity] >0 and  T0.[ItemCode] =" + "'" + ItemCode + "'" + "and  T0.[WhsCode] =" + "'" + WhsCode + "'");
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
                    oRecordSet.DoQuery("SELECT T0.[ItemCode], T0.[OnHand] FROM OITW T0 WHERE T0.[ItemCode] =" + "'" + ItemCode + "'" + "and T0.[WhsCode] =" + "'" + WhsCode + "'");
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
        public string GetTransfer(string N, int Q = 1)
        {
            GExit GExitTransfer = new GExit();
            GExitTransfer.Number = N;
            GExitTransfer.Quantity = Q;
            Transfer.Add(GExitTransfer);
            return GExitTransfer.Number;

        }
        public string GetTN(string itemcode, string FromWarehouse, string warehouse,int Quantity)
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
                    SAPbobsCOM.StockTransfer oStktransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                    if (Ser == "Y" && Btch == "N")
                    {
                        var oSN = oStktransfer.Lines.SerialNumbers;
                        oStktransfer.FromWarehouse = FromWarehouse;
                        oStktransfer.Lines.ItemCode = itemcode;
                        oStktransfer.Lines.WarehouseCode = warehouse;
                        oStktransfer.Lines.Quantity = Quantity;
                        foreach (var item in Transfer)
                        {
                            oSN.InternalSerialNumber = item.Number;
                            oSN.Add();
                        }
                    }
                    else if (Ser == "N" && Btch == "Y")
                    {
                        var obn = oStktransfer.Lines.BatchNumbers;
                        oStktransfer.FromWarehouse = FromWarehouse;
                        oStktransfer.Lines.ItemCode = itemcode;
                        oStktransfer.Lines.WarehouseCode = warehouse;
                        oStktransfer.Lines.Quantity = Quantity;
                        foreach (var item in Transfer)
                        {
                            obn.BatchNumber = item.Number;
                            obn.Quantity = item.Quantity;
                            obn.Add();
                        }

                    }
                    else
                    {
                        oStktransfer.FromWarehouse = FromWarehouse;
                        oStktransfer.Lines.ItemCode = itemcode;
                        oStktransfer.Lines.WarehouseCode = warehouse;
                        oStktransfer.Lines.Quantity = Quantity;
                    }
                    oStktransfer.Add();
                    Transfer.Clear();
                }
                return "成功";

            }
            catch (Exception errMsg)
            {
                throw errMsg;
            }
        }

        //GetDATAClear 將DATA清除
        public string GetDATAClear()
        {
            Transfer.Clear();
            return "清除";
        }
    }
}