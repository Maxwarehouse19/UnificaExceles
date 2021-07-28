﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Configuration;
using System.IO;
using System.Threading;

namespace UnificaExceles
{
    public partial class Form1 : Form
    {
        // variables para uso del programa
        // -------------------------------
        string RutaArchivoGeneracion            ="";
        string SalesOrderNumber                 ="";
        string HoldCode                         ="";
        string TotalSales                       ="";
        string SalesSku                         ="";
        string SalesCategoryAtTimeOfSale        ="";
        string UomCode                          ="";
        string UomQuantity                      ="";
        string SalesStatus                      ="";
        string SalesOrderDate                   ="";
        string SalesChannelName                 ="";
        string CustomerName                     ="";
        string FulfillmentSku                   ="";
        string FulfillmentChannelName           ="";
        string FulfillmentChannelType           ="";
        string LinkedFulfillmentChannelName     ="";
        string FulfillmentLocationName          ="";
        string FulfillmentOrderNumber           ="";
        string Quantity                         ="";
        string Sku                              ="";
        string Title                            ="";
        string TotalCost                        ="";
        string Commission                       ="";
        string InventoryCost                    ="";
        string UnitCost                         ="";
        string ServiceCost                      ="";
        string EstimatedShippingCost            ="";
        string ShippingCost                     ="";
        string ShippingPrice                    ="";
        string OverheadCost                     ="";
        string PackageCost                      ="";
        string ProfitLoss                       ="";
        string Carrier                          ="";
        string ShippingServiceLevel             ="";
        string ShippedByUser                    ="";
        string ShippingWeight                   ="";
        string Length                           ="";
        string varWidth                         ="";
        string varHeight                        ="";
        string Weight                           ="";
        string StateRegion                      ="";
        string TrackingNum                      ="";
        string MfrName                          ="";
        string PricingRule                      = "";
        string ActualShippingCost               = "";
        string ActualShipping                   = "";
        string ShippingCostDifference           = "";
        int counter = 0;
        string line;
        int cantidad = 0;
        bool Encontro = false;
        string PalabraCompleta = "";
        int ContadorProgreso = 0;
        string ArchivoLog = "";
        string ReporteLog = "";
        string LocaltextBox1 = "";
        string LocaltextBox2 = "";


        bool FlgSihayFedex = false;
        bool FlgSihayUSPS = false;
        bool FlgSihayUPS = false;
        bool EncontroRegistro = false;
        string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
        string pathString = "";

        int contador = 1;
        string BodyExcel = "<html>";

        List<PedidoFedex> listaPedido = new List<PedidoFedex>();

        List<PedidoUSPS> listaPedidoUSPS = new List<PedidoUSPS>();

        List<PedidoUPS> listaPedidoUPS = new List<PedidoUPS>();

        public Form1()
        {
            InitializeComponent();
        }

        private static string GetConnectionString(string file,string Tipo)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension.ToUpper() == "XLS"  )
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";

                if (Tipo== "MASTER")
                    props["Extended Properties"] = "Excel 8.0";
                else
                    props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension.ToUpper() == "XLSX")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";

                if (Tipo == "MASTER")
                    props["Extended Properties"] = "Excel 12.0 XML";
                else
                    props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private static DataSet GetDataSetFromExcelFile(string file,string connectionString)
        {
            DataSet ds = new DataSet();

            //string connectionString = GetConnectionString(file,);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        private static DataSet GetDataSetFromExcelFileDetalle(string file, string connectionString)
        {
            DataSet ds = new DataSet();

            //string connectionString = GetConnectionString(file,);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.Contains(" "))
                    {

                        if (!sheetName.EndsWith("$"))
                            continue;
                    }
                    else {
                        if (sheetName.Contains("FilterDatabase"))
                            continue;
                    }

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaEncabezadoReporte()
        {
            BodyExcel = "<html>";

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            RutaArchivoGeneracion = pathString +@"\"  ;//ConfigurationManager.AppSettings["RutaArchivoGeneracion"];
            RutaArchivoGeneracion = RutaArchivoGeneracion + "ReporteOutput" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".xls";
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                BodyExcel = BodyExcel + "<body>";
                BodyExcel = BodyExcel + "<table>";
                BodyExcel = BodyExcel + @"<tr bgcolor= ""#CA2229"" style=""color:#ffffff"">";
                // Archivo DW
                // ----------
                BodyExcel = BodyExcel + "<td> SalesOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> HoldCode</td>";
                BodyExcel = BodyExcel + "<td> TotalSales</td>";
                BodyExcel = BodyExcel + "<td> SalesSku</td>";
                BodyExcel = BodyExcel + "<td> SalesCategoryAtTimeOfSale</td>";
                BodyExcel = BodyExcel + "<td> UomCode</td>";
                BodyExcel = BodyExcel + "<td> UomQuantity</td>";
                BodyExcel = BodyExcel + "<td> SalesStatus</td>";
                BodyExcel = BodyExcel + "<td> SalesOrderDate</td>";
                BodyExcel = BodyExcel + "<td> SalesChannelName</td>";
                BodyExcel = BodyExcel + "<td> CustomerName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentSku</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelType</td>";
                BodyExcel = BodyExcel + "<td> LinkedFulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentLocationName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> Quantity</td>";
                BodyExcel = BodyExcel + "<td> Sku</td>";
                BodyExcel = BodyExcel + "<td> Title</td>";
                BodyExcel = BodyExcel + "<td> TotalCost</td>";
                BodyExcel = BodyExcel + "<td> Commission</td>";
                BodyExcel = BodyExcel + "<td> InventoryCost</td>";
                BodyExcel = BodyExcel + "<td> UnitCost</td>";
                BodyExcel = BodyExcel + "<td> ServiceCost</td>";
                BodyExcel = BodyExcel + "<td> EstimatedShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingPrice</td>";
                BodyExcel = BodyExcel + "<td> OverheadCost</td>";
                BodyExcel = BodyExcel + "<td> PackageCost</td>";
                BodyExcel = BodyExcel + "<td> ProfitLoss</td>";
                BodyExcel = BodyExcel + "<td> Carrier</td>";
                BodyExcel = BodyExcel + "<td> ShippingServiceLevel</td>";
                BodyExcel = BodyExcel + "<td> ShippedByUser</td>";
                BodyExcel = BodyExcel + "<td> ShippingWeight</td>";
                BodyExcel = BodyExcel + "<td> Length</td>";
                BodyExcel = BodyExcel + "<td> Width</td>";
                BodyExcel = BodyExcel + "<td> Height</td>";
                BodyExcel = BodyExcel + "<td> Weight</td>";
                BodyExcel = BodyExcel + "<td> StateRegion</td>";
                BodyExcel = BodyExcel + "<td> TrackingNum</td>";
                BodyExcel = BodyExcel + "<td> MfrName</td>";
                BodyExcel = BodyExcel + "<td> PricingRule</td>";


                // Archivo Fedex
                // -------------
                //BodyExcel = BodyExcel + "<td>FullTrakingId</td>";
                BodyExcel = BodyExcel + "<td> Ground Tracking ID Prefix</td>";
                BodyExcel = BodyExcel + "<td> Express or Ground Tracking ID</td>";
                BodyExcel = BodyExcel + "<td> Net Charge Amount</td>";
                BodyExcel = BodyExcel + "<td> Service Type</td>";
                BodyExcel = BodyExcel + "<td> Ground Service</td>";
                BodyExcel = BodyExcel + "<td> Shipment Date</td>";
                BodyExcel = BodyExcel + "<td> POD Delivery Date</td>";
                BodyExcel = BodyExcel + "<td> Actual Weight Amount</td>";
                BodyExcel = BodyExcel + "<td> Rated Weight Amount</td>";
                BodyExcel = BodyExcel + "<td> Dim Length</td>";
                BodyExcel = BodyExcel + "<td> Dim Width</td>";
                BodyExcel = BodyExcel + "<td> Dim Height</td>";
                BodyExcel = BodyExcel + "<td> Dim Divisor</td>";
                BodyExcel = BodyExcel + "<td> Shipper State</td>";
                BodyExcel = BodyExcel + "<td> Zone Code</td>";
                BodyExcel = BodyExcel + "<td> Tendered Date</td>";

                // cargos fijos
                // ------------
                BodyExcel = BodyExcel + "<td>Earned Discount</td>";
                BodyExcel = BodyExcel + "<td>Fuel Surcharge</td>";
                BodyExcel = BodyExcel + "<td>Performance Pricing</td>";
                BodyExcel = BodyExcel + "<td>Delivery Area Surcharge Extended</td>";
                BodyExcel = BodyExcel + "<td>Delivery Area Surcharge</td>";
                BodyExcel = BodyExcel + "<td>USPS Non-Mach Surcharge</td>";
                BodyExcel = BodyExcel + "<td>Residential</td>";
                BodyExcel = BodyExcel + "<td>Grace Discount</td>";
                BodyExcel = BodyExcel + "<td>Declared Value</td>";
                BodyExcel = BodyExcel + "<td>DAS Extended Resi</td>";
                BodyExcel = BodyExcel + "<td>Additional Handling</td>";
                BodyExcel = BodyExcel + "<td>Parcel Re-Label Charge</td>";
                BodyExcel = BodyExcel + "<td>Indirect Signature</td>";
                BodyExcel = BodyExcel + "<td>DAS Resi</td>";
                BodyExcel = BodyExcel + "<td>Address Correction</td>";
                BodyExcel = BodyExcel + "<td>DAS Extended Comm</td>";
                BodyExcel = BodyExcel + "<td>Oversize Charge</td>";
                BodyExcel = BodyExcel + "<td>AHS - Dimensions</td>";

                // dato USPS
                BodyExcel = BodyExcel + "<td>Ground Service </td>";
                BodyExcel = BodyExcel + "<td>Tracking Number </td>";
                BodyExcel = BodyExcel + "<td>Net Charge Amount </td>";
                BodyExcel = BodyExcel + "<td>POD Delivery Date </td>";
                BodyExcel = BodyExcel + "<td>Rated Weight Amount </td>";
                BodyExcel = BodyExcel + "<td>Zone Code </td>";

                BodyExcel = BodyExcel + "</tr>";

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";
            }

        }

        // realiza la impresion del cargo enviado si la tuviera el reporte de fedex
        // -------------------------------------------------------------------------------
        private void ColumnaCargo(string NombreCargo, string TrackingIDChargeDescription, string TrackingIDChargeAmount, string TrackingIDChargeDescription1, string TrackingIDChargeAmount1, string TrackingIDChargeDescription2, string TrackingIDChargeAmount2, string TrackingIDChargeDescription3, string TrackingIDChargeAmount3, string TrackingIDChargeDescription4, string TrackingIDChargeAmount4, string TrackingIDChargeDescription5, string TrackingIDChargeAmount5, string TrackingIDChargeDescription6, string TrackingIDChargeAmount6, string TrackingIDChargeDescription7, string TrackingIDChargeAmount7, string TrackingIDChargeDescription8, string TrackingIDChargeAmount8, string TrackingIDChargeDescription9, string TrackingIDChargeAmount9, string TrackingIDChargeDescription10, string TrackingIDChargeAmount10, string TrackingIDChargeDescription11, string TrackingIDChargeAmount11, string TrackingIDChargeDescription12, string TrackingIDChargeAmount12, string TrackingIDChargeDescription13, string TrackingIDChargeAmount13, string TrackingIDChargeDescription14, string TrackingIDChargeAmount14, string TrackingIDChargeDescription15, string TrackingIDChargeAmount15, string TrackingIDChargeDescription16, string TrackingIDChargeAmount16, string TrackingIDChargeDescription17, string TrackingIDChargeAmount17, string TrackingIDChargeDescription18, string TrackingIDChargeAmount18, string TrackingIDChargeDescription19, string TrackingIDChargeAmount19, string TrackingIDChargeDescription20, string TrackingIDChargeAmount20, string TrackingIDChargeDescription21, string TrackingIDChargeAmount21, string TrackingIDChargeDescription22, string TrackingIDChargeAmount22, string TrackingIDChargeDescription23, string TrackingIDChargeAmount23, string TrackingIDChargeDescription24, string TrackingIDChargeAmount24)
        {
            if (TrackingIDChargeDescription == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount + "</td>";
            else if (TrackingIDChargeDescription1 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount1 + "</td>";
            else if (TrackingIDChargeDescription2 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount2 + "</td>";
            else if (TrackingIDChargeDescription3 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount3 + "</td>";
            else if (TrackingIDChargeDescription4 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount4 + "</td>";
            else if (TrackingIDChargeDescription5 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount5 + "</td>";
            else if (TrackingIDChargeDescription6 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount6 + "</td>";
            else if (TrackingIDChargeDescription7 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount7 + "</td>";
            else if (TrackingIDChargeDescription8 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount8 + "</td>";
            else if (TrackingIDChargeDescription9 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount9 + "</td>";
            else if (TrackingIDChargeDescription10 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount10 + "</td>";
            else if (TrackingIDChargeDescription11 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount11 + "</td>";
            else if (TrackingIDChargeDescription12 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount12 + "</td>";
            else if (TrackingIDChargeDescription13 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount13 + "</td>";
            else if (TrackingIDChargeDescription14 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount14 + "</td>";
            else if (TrackingIDChargeDescription15 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount15 + "</td>";
            else if (TrackingIDChargeDescription16 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount16 + "</td>";
            else if (TrackingIDChargeDescription17 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount17 + "</td>";
            else if (TrackingIDChargeDescription18 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount18 + "</td>";
            else if (TrackingIDChargeDescription19 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount19 + "</td>";
            else if (TrackingIDChargeDescription20 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount20 + "</td>";
            else if (TrackingIDChargeDescription21 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount21 + "</td>";
            else if (TrackingIDChargeDescription22 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount22 + "</td>";
            else if (TrackingIDChargeDescription23 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount23 + "</td>";
            else if (TrackingIDChargeDescription24 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount24 + "</td>";
            else
                BodyExcel = BodyExcel + "<td>" + " " + "</td>";
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporte()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                var PedidoCollection = from s in listaPedido
                                       where s.FullTrakingId == TrackingNum
                                       select new
                                       {
                                           s.FullTrakingId,
                                           s.BilltoAccountNumber,
                                           s.InvoiceDate,
                                           s.InvoiceNumber,
                                           s.StoreID,
                                           s.OriginalAmountDue,
                                           s.CurrentBalance,
                                           s.Payor,
                                           s.GroundTrackingIDPrefix,
                                           s.ExpressorGroundTrackingID,
                                           s.TransportationChargeAmount,
                                           s.NetChargeAmount,
                                           s.ServiceType,
                                           s.GroundService,
                                           s.ShipmentDate,
                                           s.PODDeliveryDate,
                                           s.PODDeliveryTime,
                                           s.PODServiceAreaCode,
                                           s.PODSignatureDescription,
                                           s.ActualWeightAmount,
                                           s.ActualWeightUnits,
                                           s.RatedWeightAmount,
                                           s.RatedWeightUnits,
                                           s.NumberofPieces,
                                           s.BundleNumber,
                                           s.MeterNumber,
                                           s.TDMasterTrackingID,
                                           s.ServicePackaging,
                                           s.DimLength,
                                           s.DimWidth,
                                           s.DimHeight,
                                           s.DimDivisor,
                                           s.DimUnit,
                                           s.RecipientName,
                                           s.RecipientCompany,
                                           s.RecipientAddressLine1,
                                           s.RecipientAddressLine2,
                                           s.RecipientCity,
                                           s.RecipientState,
                                           s.RecipientZipCode,
                                           s.RecipientCountryTerritory,
                                           s.ShipperCompany,
                                           s.ShipperName,
                                           s.ShipperAddressLine1,
                                           s.ShipperAddressLine2,
                                           s.ShipperCity,
                                           s.ShipperState,
                                           s.ShipperZipCode,
                                           s.ShipperCountryTerritory,
                                           s.OriginalCustomerReference,
                                           s.OriginalRef2,
                                           s.OriginalRef3PONumber,
                                           s.OriginalDepartmentReferenceDescription,
                                           s.UpdatedCustomerReference,
                                           s.UpdatedRef2,
                                           s.UpdatedRef3PONumber,
                                           s.UpdatedDepartmentReferenceDescription,
                                           s.RMA,
                                           s.OriginalRecipientAddressLine1,
                                           s.OriginalRecipientAddressLine2,
                                           s.OriginalRecipientCity,
                                           s.OriginalRecipientState,
                                           s.OriginalRecipientZipCode,
                                           s.OriginalRecipientCountryTerritory,
                                           s.ZoneCode,
                                           s.CostAllocation,
                                           s.AlternateAddressLine1,
                                           s.AlternateAddressLine2,
                                           s.AlternateCity,
                                           s.AlternateStateProvince,
                                           s.AlternateZipCode,
                                           s.AlternateCountryTerritoryCode,
                                           s.CrossRefTrackingIDPrefix,
                                           s.CrossRefTrackingID,
                                           s.EntryDate,
                                           s.EntryNumber,
                                           s.CustomsValue,
                                           s.CustomsValueCurrencyCode,
                                           s.DeclaredValue,
                                           s.DeclaredValueCurrencyCode,
                                           s.CommodityDescription,
                                           s.CommodityCountryTerritoryCode,
                                           s.CommodityDescription1,
                                           s.CommodityCountryTerritoryCode1,
                                           s.CommodityDescription2,
                                           s.CommodityCountryTerritoryCode2,
                                           s.CommodityDescription3,
                                           s.CommodityCountryTerritoryCode3,
                                           s.CurrencyConversionDate,
                                           s.CurrencyConversionRate,
                                           s.MultiweightNumber,
                                           s.MultiweightTotalMultiweightUnits,
                                           s.MultiweightTotalMultiweightWeight,
                                           s.MultiweightTotalShipmentChargeAmount,
                                           s.MultiweightTotalShipmentWeight,
                                           s.GroundTrackingIDAddressCorrectionDiscountChargeAmount,
                                           s.GroundTrackingIDAddressCorrectionGrossChargeAmount,
                                           s.RatedMethod,
                                           s.SortHub,
                                           s.EstimatedWeight,
                                           s.EstimatedWeightUnit,
                                           s.PostalClass,
                                           s.ProcessCategory,
                                           s.PackageSize,
                                           s.DeliveryConfirmation,
                                           s.TenderedDate,
                                           s.TrackingIDChargeDescription,
                                           s.TrackingIDChargeAmount,
                                           s.TrackingIDChargeDescription1,
                                           s.TrackingIDChargeAmount1,
                                           s.TrackingIDChargeDescription2,
                                           s.TrackingIDChargeAmount2,
                                           s.TrackingIDChargeDescription3,
                                           s.TrackingIDChargeAmount3,
                                           s.TrackingIDChargeDescription4,
                                           s.TrackingIDChargeAmount4,
                                           s.TrackingIDChargeDescription5,
                                           s.TrackingIDChargeAmount5,
                                           s.TrackingIDChargeDescription6,
                                           s.TrackingIDChargeAmount6,
                                           s.TrackingIDChargeDescription7,
                                           s.TrackingIDChargeAmount7,
                                           s.TrackingIDChargeDescription8,
                                           s.TrackingIDChargeAmount8,
                                           s.TrackingIDChargeDescription9,
                                           s.TrackingIDChargeAmount9,
                                           s.TrackingIDChargeDescription10,
                                           s.TrackingIDChargeAmount10,
                                           s.TrackingIDChargeDescription11,
                                           s.TrackingIDChargeAmount11,
                                           s.TrackingIDChargeDescription12,
                                           s.TrackingIDChargeAmount12,
                                           s.TrackingIDChargeDescription13,
                                           s.TrackingIDChargeAmount13,
                                           s.TrackingIDChargeDescription14,
                                           s.TrackingIDChargeAmount14,
                                           s.TrackingIDChargeDescription15,
                                           s.TrackingIDChargeAmount15,
                                           s.TrackingIDChargeDescription16,
                                           s.TrackingIDChargeAmount16,
                                           s.TrackingIDChargeDescription17,
                                           s.TrackingIDChargeAmount17,
                                           s.TrackingIDChargeDescription18,
                                           s.TrackingIDChargeAmount18,
                                           s.TrackingIDChargeDescription19,
                                           s.TrackingIDChargeAmount19,
                                           s.TrackingIDChargeDescription20,
                                           s.TrackingIDChargeAmount20,
                                           s.TrackingIDChargeDescription21,
                                           s.TrackingIDChargeAmount21,
                                           s.TrackingIDChargeDescription22,
                                           s.TrackingIDChargeAmount22,
                                           s.TrackingIDChargeDescription23,
                                           s.TrackingIDChargeAmount23,
                                           s.TrackingIDChargeDescription24,
                                           s.TrackingIDChargeAmount24,
                                           s.ShipmentNotes
                                       };

                foreach (var Pedido in PedidoCollection)
                {
                    // arma la fila con el color de fondo que corresponde
                    // --------------------------------------------------
                    if (contador == 1)
                        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                    else
                        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                    // archivo base
                    // ------------
                    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                    // archivo fedex
                    // -------------
                    BodyExcel = BodyExcel + "<td>'" + Pedido.GroundTrackingIDPrefix + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + Pedido.ExpressorGroundTrackingID+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.NetChargeAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ServiceType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.GroundService + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ShipmentDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.PODDeliveryDate+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ActualWeightAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.RatedWeightAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimLength + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimDivisor + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ShipperState + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ZoneCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.TenderedDate + "</td>";

                    string NombreCargo = "Earned Discount";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Fuel Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Performance Pricing";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Delivery Area Surcharge Extended";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Delivery Area Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "USPS Non-Mach Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Residential";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Grace Discount";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Declared Value";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Extended Resi";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Additional Handling";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Parcel Re-Label Charge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Indirect Signature";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Address Correction";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Extended Comm";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Oversize Charge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "AHS - Dimensions";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    BodyExcel = BodyExcel + "</tr>";
                    break;
                }

                FileExcel.WriteLine(BodyExcel);
                BodyExcel= "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporteUSPS()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                var PedidoCollection = from s in listaPedidoUSPS
                                       where s.TrackingNumber == TrackingNum
                                       select new
                                       {
                                            s.GroundService,
                                            s.TrackingNumber,
                                            s.NetChargeAmount,
                                            s.PODDeliveryDate,
                                            s.RatedWeightAmount,
                                            s.ZoneCode
                                       };

                foreach (var Pedido in PedidoCollection)
                {
                    // arma la fila con el color de fondo que corresponde
                    // --------------------------------------------------
                    if (contador == 1)
                        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                    else
                        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                    // archivo base
                    // ------------
                    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                    // archivo fedex
                    // -------------
                    string vacio = ""; 
                    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";

                    string NombreCargo = "Earned Discount";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Fuel Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Performance Pricing";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Delivery Area Surcharge Extended";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Delivery Area Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "USPS Non-Mach Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Residential";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Grace Discount";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Declared Value";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Extended Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Additional Handling";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Parcel Re-Label Charge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Indirect Signature";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Address Correction";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Extended Comm";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Oversize Charge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "AHS - Dimensions";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    // dato USPS
                    BodyExcel = BodyExcel + "<td>" + Pedido.GroundService + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.TrackingNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.NetChargeAmount + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.PODDeliveryDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.RatedWeightAmount + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ZoneCode + "</td>";

                    BodyExcel = BodyExcel + "</tr>";
                    break;
                }

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporteUPS()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                var PedidoCollection = from s in listaPedidoUPS
                                       where s.Campo30 == TrackingNum
                                       select new
                                       {
                                           s.Campo12,
                                           s.Campo30,
                                           s.Campo39
                                       };

                foreach (var Pedido in PedidoCollection)
                {
                    // arma la fila con el color de fondo que corresponde
                    // --------------------------------------------------
                    if (contador == 1)
                        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                    else
                        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                    // archivo base
                    // ------------
                    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                    // archivo fedex
                    // -------------
                    string vacio = "";
                    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";

                    string NombreCargo = "Earned Discount";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Fuel Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Performance Pricing";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Delivery Area Surcharge Extended";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Delivery Area Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "USPS Non-Mach Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Residential";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Grace Discount";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Declared Value";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Extended Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Additional Handling";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Parcel Re-Label Charge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Indirect Signature";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Address Correction";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Extended Comm";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Oversize Charge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "AHS - Dimensions";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    // dato USPS
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";

                    // dato UPS
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo12 + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo30 + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo39 + "</td>";

                    BodyExcel = BodyExcel + "</tr>";
                    break;
                }

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistro(string[] valor/*DataRow row*/)
        {
            SalesOrderNumber             = valor[0];//Convert.ToString(row["SalesOrderNumber"]);
            HoldCode                     = valor[1];//Convert.ToString(row["HoldCode"]);
            TotalSales                   = valor[2];//Convert.ToString(row["TotalSales"]);
            SalesSku                     = valor[3];//Convert.ToString(row["SalesSku"]);
            SalesCategoryAtTimeOfSale    = valor[4];//Convert.ToString(row["SalesCategoryAtTimeOfSale"]);
            UomCode                      = valor[5];//Convert.ToString(row["UomCode"]);
            UomQuantity                  = valor[6];//Convert.ToString(row["UomQuantity"]);
            SalesStatus                  = valor[7];//Convert.ToString(row["SalesStatus"]);
            SalesOrderDate               = valor[8];//Convert.ToString(row["SalesOrderDate"]);
            SalesChannelName             = valor[9];//Convert.ToString(row["SalesChannelName"]);
            CustomerName                 = valor[10];//Convert.ToString(row["CustomerName"]);
            FulfillmentSku               = valor[11];//Convert.ToString(row["FulfillmentSku"]);
            FulfillmentChannelName       = valor[12];//Convert.ToString(row["FulfillmentChannelName"]);
            FulfillmentChannelType       = valor[13];//Convert.ToString(row["FulfillmentChannelType"]);
            LinkedFulfillmentChannelName = valor[14];//Convert.ToString(row["LinkedFulfillmentChannelName"]);
            FulfillmentLocationName      = valor[15];//Convert.ToString(row["FulfillmentLocationName"]);
            FulfillmentOrderNumber       = valor[16];//Convert.ToString(row["FulfillmentOrderNumber"]);
            Quantity                     = valor[17];//Convert.ToString(row["Quantity"]);
            Sku                          = valor[18];//Convert.ToString(row["Sku"]);
            Title                        = valor[19];//Convert.ToString(row["Title"]);
            TotalCost                    = valor[20];//Convert.ToString(row["TotalCost"]);
            Commission                   = valor[21];//Convert.ToString(row["Commission"]);
            InventoryCost                = valor[22];//Convert.ToString(row["InventoryCost"]);
            UnitCost                     = valor[23];//Convert.ToString(row["UnitCost"]);
            ServiceCost                  = valor[24];//Convert.ToString(row["ServiceCost"]);
            EstimatedShippingCost        = valor[25];//Convert.ToString(row["EstimatedShippingCost"]);
            ShippingCost                 = valor[26];//Convert.ToString(row["ShippingCost"]);
            ShippingPrice                = valor[27];//Convert.ToString(row["ShippingPrice"]);
            OverheadCost                 = valor[28];//Convert.ToString(row["OverheadCost"]);
            PackageCost                  = valor[29];//Convert.ToString(row["PackageCost"]);
            ProfitLoss                   = valor[30];//Convert.ToString(row["ProfitLoss"]);
            Carrier                      = valor[31];//Convert.ToString(row["Carrier"]);
            ShippingServiceLevel         = valor[32];//Convert.ToString(row["ShippingServiceLevel"]);
            ShippedByUser                = valor[33];//Convert.ToString(row["ShippedByUser"]);
            ShippingWeight               = valor[34];//Convert.ToString(row["ShippingWeight"]);
            Length                       = valor[35];//Convert.ToString(row["Length"]);
            varWidth                     = valor[36];//Convert.ToString(row["Width"]);
            varHeight                    = valor[37];//Convert.ToString(row["Height"]);
            Weight                       = valor[38];//Convert.ToString(row["Weight"]);
            StateRegion                  = valor[39];//Convert.ToString(row["StateRegion"]);
            TrackingNum                  = valor[40];//Convert.ToString(row["TrackingNum"]);
            MfrName                      = valor[41];//Convert.ToString(row["MfrName"]);
            PricingRule                  = valor[42];//Convert.ToString(row["PricingRule"]);
            ActualShippingCost           = valor[43];
            ActualShipping               = valor[44];
            ShippingCostDifference       = valor[45];
        }
        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalle(DataRow row, ref PedidoFedex clsPedido)
        {
            clsPedido.BilltoAccountNumber = Convert.ToString(row["Bill to Account Number"]);
            clsPedido.InvoiceDate = Convert.ToString(row["Invoice Date"]);
            clsPedido.InvoiceNumber = Convert.ToString(row["Invoice Number"]);
            clsPedido.StoreID = Convert.ToString(row["Store ID"]);
            clsPedido.OriginalAmountDue = Convert.ToString(row["Original Amount Due"]);
            clsPedido.CurrentBalance = Convert.ToString(row["Current Balance"]);
            clsPedido.Payor = Convert.ToString(row["Payor"]);
            clsPedido.GroundTrackingIDPrefix = Convert.ToString(row["Ground Tracking ID Prefix"]);
            clsPedido.ExpressorGroundTrackingID = Convert.ToString(row["Express or Ground Tracking ID"]);
            clsPedido.FullTrakingId = clsPedido.GroundTrackingIDPrefix + clsPedido.ExpressorGroundTrackingID;
            clsPedido.TransportationChargeAmount = Convert.ToString(row["Transportation Charge Amount"]);
            clsPedido.NetChargeAmount = Convert.ToString(row["Net Charge Amount"]);
            clsPedido.ServiceType = Convert.ToString(row["Service Type"]);
            clsPedido.GroundService = Convert.ToString(row["Ground Service"]);
            clsPedido.ShipmentDate = Convert.ToString(row["Shipment Date"]);
            clsPedido.PODDeliveryDate = Convert.ToString(row["POD Delivery Date"]);
            clsPedido.PODDeliveryTime = Convert.ToString(row["POD Delivery Time"]);
            clsPedido.PODServiceAreaCode = Convert.ToString(row["POD Service Area Code"]);
            clsPedido.PODSignatureDescription = Convert.ToString(row["POD Signature Description"]);
            clsPedido.ActualWeightAmount = Convert.ToString(row["Actual Weight Amount"]);
            clsPedido.ActualWeightUnits = Convert.ToString(row["Actual Weight Units"]);
            clsPedido.RatedWeightAmount = Convert.ToString(row["Rated Weight Amount"]);
            clsPedido.RatedWeightUnits = Convert.ToString(row["Rated Weight Units"]);
            clsPedido.NumberofPieces = Convert.ToString(row["Number of Pieces"]);
            clsPedido.BundleNumber = Convert.ToString(row["Bundle Number"]);
            clsPedido.MeterNumber = Convert.ToString(row["Meter Number"]);
            clsPedido.TDMasterTrackingID = Convert.ToString(row["TDMasterTrackingID"]);
            clsPedido.ServicePackaging = Convert.ToString(row["Service Packaging"]);
            clsPedido.DimLength = Convert.ToString(row["Dim Length"]);
            clsPedido.DimWidth = Convert.ToString(row["Dim Width"]);
            clsPedido.DimHeight = Convert.ToString(row["Dim Height"]);
            clsPedido.DimDivisor = Convert.ToString(row["Dim Divisor"]);
            clsPedido.DimUnit = Convert.ToString(row["Dim Unit"]);
            clsPedido.RecipientName = Convert.ToString(row["Recipient Name"]);
            clsPedido.RecipientCompany = Convert.ToString(row["Recipient Company"]);
            clsPedido.RecipientAddressLine1 = Convert.ToString(row["Recipient Address Line 1"]);
            clsPedido.RecipientAddressLine2 = Convert.ToString(row["Recipient Address Line 2"]);
            clsPedido.RecipientCity = Convert.ToString(row["Recipient City"]);
            clsPedido.RecipientState = Convert.ToString(row["Recipient State"]);
            clsPedido.RecipientZipCode = Convert.ToString(row["Recipient Zip Code"]);
            clsPedido.RecipientCountryTerritory = Convert.ToString(row["Recipient Country/Territory"]);
            clsPedido.ShipperCompany = Convert.ToString(row["Shipper Company"]);
            clsPedido.ShipperName = Convert.ToString(row["Shipper Name"]);
            clsPedido.ShipperAddressLine1 = Convert.ToString(row["Shipper Address Line 1"]);
            clsPedido.ShipperAddressLine2 = Convert.ToString(row["Shipper Address Line 2"]);
            clsPedido.ShipperCity = Convert.ToString(row["Shipper City"]);
            clsPedido.ShipperState = Convert.ToString(row["Shipper State"]);
            clsPedido.ShipperZipCode = Convert.ToString(row["Shipper Zip Code"]);
            clsPedido.ShipperCountryTerritory = Convert.ToString(row["Shipper Country/Territory"]);
            clsPedido.OriginalCustomerReference = Convert.ToString(row["Original Customer Reference"]);
            clsPedido.OriginalRef2 = Convert.ToString(row["Original Ref#2"]);
            clsPedido.OriginalRef3PONumber = Convert.ToString(row["Original Ref#3/PO Number"]);
            clsPedido.OriginalDepartmentReferenceDescription = Convert.ToString(row["Original Department Reference Description"]);
            clsPedido.UpdatedCustomerReference = Convert.ToString(row["Updated Customer Reference"]);
            clsPedido.UpdatedRef2 = Convert.ToString(row["Updated Ref#2"]);
            clsPedido.UpdatedRef3PONumber = Convert.ToString(row["Updated Ref#3/PO Number"]);
            clsPedido.UpdatedDepartmentReferenceDescription = Convert.ToString(row["Updated Department Reference Description"]);
            clsPedido.RMA = Convert.ToString(row["RMA#"]);
            clsPedido.OriginalRecipientAddressLine1 = Convert.ToString(row["Original Recipient Address Line 1"]);
            clsPedido.OriginalRecipientAddressLine2 = Convert.ToString(row["Original Recipient Address Line 2"]);
            clsPedido.OriginalRecipientCity = Convert.ToString(row["Original Recipient City"]);
            clsPedido.OriginalRecipientState = Convert.ToString(row["Original Recipient State"]);
            clsPedido.OriginalRecipientZipCode = Convert.ToString(row["Original Recipient Zip Code"]);
            clsPedido.OriginalRecipientCountryTerritory = Convert.ToString(row["Original Recipient Country/Territory"]);
            clsPedido.ZoneCode = Convert.ToString(row["Zone Code"]);
            clsPedido.CostAllocation = Convert.ToString(row["Cost Allocation"]);
            clsPedido.AlternateAddressLine1 = Convert.ToString(row["Alternate Address Line 1"]);
            clsPedido.AlternateAddressLine2 = Convert.ToString(row["Alternate Address Line 2"]);
            clsPedido.AlternateCity = Convert.ToString(row["Alternate City"]);
            clsPedido.AlternateStateProvince = Convert.ToString(row["Alternate State Province"]);
            clsPedido.AlternateZipCode = Convert.ToString(row["Alternate Zip Code"]);
            clsPedido.AlternateCountryTerritoryCode = Convert.ToString(row["Alternate Country/Territory Code"]);
            clsPedido.CrossRefTrackingIDPrefix = Convert.ToString(row["CrossRefTrackingID Prefix"]);
            clsPedido.CrossRefTrackingID = Convert.ToString(row["CrossRefTrackingID"]);
            clsPedido.EntryDate = Convert.ToString(row["Entry Date"]);
            clsPedido.EntryNumber = Convert.ToString(row["Entry Number"]);
            clsPedido.CustomsValue = Convert.ToString(row["Customs Value"]);
            clsPedido.CustomsValueCurrencyCode = Convert.ToString(row["Customs Value Currency Code"]);
            clsPedido.DeclaredValue = Convert.ToString(row["Declared Value"]);
            clsPedido.DeclaredValueCurrencyCode = Convert.ToString(row["Declared Value Currency Code"]);
            clsPedido.CommodityDescription = Convert.ToString(row["Commodity Description"]);
            clsPedido.CommodityCountryTerritoryCode = Convert.ToString(row["Commodity Country/Territory Code"]);
            clsPedido.CommodityDescription1 = Convert.ToString(row["Commodity Description"]);
            clsPedido.CommodityCountryTerritoryCode1 = Convert.ToString(row["Commodity Country/Territory Code1"]);
            clsPedido.CommodityDescription2 = Convert.ToString(row["Commodity Description1"]);
            clsPedido.CommodityCountryTerritoryCode2 = Convert.ToString(row["Commodity Country/Territory Code2"]);
            clsPedido.CommodityDescription3 = Convert.ToString(row["Commodity Description2"]);
            clsPedido.CommodityCountryTerritoryCode3 = Convert.ToString(row["Commodity Country/Territory Code3"]);
            clsPedido.CurrencyConversionDate = Convert.ToString(row["Currency Conversion Date"]);
            clsPedido.CurrencyConversionRate = Convert.ToString(row["Currency Conversion Rate"]);
            clsPedido.MultiweightNumber = Convert.ToString(row["Multiweight Number"]);
            clsPedido.MultiweightTotalMultiweightUnits = Convert.ToString(row["Multiweight Total Multiweight Units"]);
            clsPedido.MultiweightTotalMultiweightWeight = Convert.ToString(row["Multiweight Total Multiweight Weight"]);
            clsPedido.MultiweightTotalShipmentChargeAmount = Convert.ToString(row["Multiweight Total Shipment Charge Amount"]);
            clsPedido.MultiweightTotalShipmentWeight = Convert.ToString(row["Multiweight Total Shipment Weight"]);
            clsPedido.GroundTrackingIDAddressCorrectionDiscountChargeAmount = Convert.ToString(row["Ground Tracking ID Address Correction Discount Charge Amount"]);
            clsPedido.GroundTrackingIDAddressCorrectionGrossChargeAmount = Convert.ToString(row["Ground Tracking ID Address Correction Gross Charge Amount"]);
            clsPedido.RatedMethod = Convert.ToString(row["Rated Method"]);
            clsPedido.SortHub = Convert.ToString(row["Sort Hub"]);
            clsPedido.EstimatedWeight = Convert.ToString(row["Estimated Weight"]);
            clsPedido.EstimatedWeightUnit = Convert.ToString(row["Estimated Weight Unit"]);
            clsPedido.PostalClass = Convert.ToString(row["Postal Class"]);
            clsPedido.ProcessCategory = Convert.ToString(row["Process Category"]);
            clsPedido.PackageSize = Convert.ToString(row["Package Size"]);
            clsPedido.DeliveryConfirmation = Convert.ToString(row["Delivery Confirmation"]);
            clsPedido.TenderedDate = Convert.ToString(row["Tendered Date"]);
            clsPedido.TrackingIDChargeDescription = Convert.ToString(row["Tracking ID Charge Description"]);
            clsPedido.TrackingIDChargeAmount = Convert.ToString(row["Tracking ID Charge Amount"]);
            clsPedido.TrackingIDChargeDescription1 = Convert.ToString(row["Tracking ID Charge Description1"]);
            clsPedido.TrackingIDChargeAmount1 = Convert.ToString(row["Tracking ID Charge Amount1"]);
            clsPedido.TrackingIDChargeDescription2 = Convert.ToString(row["Tracking ID Charge Description2"]);
            clsPedido.TrackingIDChargeAmount2 = Convert.ToString(row["Tracking ID Charge Amount2"]);
            clsPedido.TrackingIDChargeDescription3 = Convert.ToString(row["Tracking ID Charge Description3"]);
            clsPedido.TrackingIDChargeAmount3 = Convert.ToString(row["Tracking ID Charge Amount3"]);
            clsPedido.TrackingIDChargeDescription4 = Convert.ToString(row["Tracking ID Charge Description4"]);
            clsPedido.TrackingIDChargeAmount4 = Convert.ToString(row["Tracking ID Charge Amount4"]);
            clsPedido.TrackingIDChargeDescription5 = Convert.ToString(row["Tracking ID Charge Description5"]);
            clsPedido.TrackingIDChargeAmount5 = Convert.ToString(row["Tracking ID Charge Amount5"]);
            clsPedido.TrackingIDChargeDescription6 = Convert.ToString(row["Tracking ID Charge Description6"]);
            clsPedido.TrackingIDChargeAmount6 = Convert.ToString(row["Tracking ID Charge Amount6"]);
            clsPedido.TrackingIDChargeDescription7 = Convert.ToString(row["Tracking ID Charge Description7"]);
            clsPedido.TrackingIDChargeAmount7 = Convert.ToString(row["Tracking ID Charge Amount7"]);
            clsPedido.TrackingIDChargeDescription8 = Convert.ToString(row["Tracking ID Charge Description8"]);
            clsPedido.TrackingIDChargeAmount8 = Convert.ToString(row["Tracking ID Charge Amount8"]);
            clsPedido.TrackingIDChargeDescription9 = Convert.ToString(row["Tracking ID Charge Description9"]);
            clsPedido.TrackingIDChargeAmount9 = Convert.ToString(row["Tracking ID Charge Amount9"]);
            clsPedido.TrackingIDChargeDescription10 = Convert.ToString(row["Tracking ID Charge Description10"]);
            clsPedido.TrackingIDChargeAmount10 = Convert.ToString(row["Tracking ID Charge Amount10"]);
            clsPedido.TrackingIDChargeDescription11 = Convert.ToString(row["Tracking ID Charge Description11"]);
            clsPedido.TrackingIDChargeAmount11 = Convert.ToString(row["Tracking ID Charge Amount11"]);
            clsPedido.TrackingIDChargeDescription12 = Convert.ToString(row["Tracking ID Charge Description12"]);
            clsPedido.TrackingIDChargeAmount12 = Convert.ToString(row["Tracking ID Charge Amount12"]);
            clsPedido.TrackingIDChargeDescription13 = Convert.ToString(row["Tracking ID Charge Description13"]);
            clsPedido.TrackingIDChargeAmount13 = Convert.ToString(row["Tracking ID Charge Amount13"]);
            clsPedido.TrackingIDChargeDescription14 = Convert.ToString(row["Tracking ID Charge Description14"]);
            clsPedido.TrackingIDChargeAmount14 = Convert.ToString(row["Tracking ID Charge Amount14"]);
            clsPedido.TrackingIDChargeDescription15 = Convert.ToString(row["Tracking ID Charge Description15"]);
            clsPedido.TrackingIDChargeAmount15 = Convert.ToString(row["Tracking ID Charge Amount15"]);
            clsPedido.TrackingIDChargeDescription16 = Convert.ToString(row["Tracking ID Charge Description16"]);
            clsPedido.TrackingIDChargeAmount16 = Convert.ToString(row["Tracking ID Charge Amount16"]);
            clsPedido.TrackingIDChargeDescription17 = Convert.ToString(row["Tracking ID Charge Description17"]);
            clsPedido.TrackingIDChargeAmount17 = Convert.ToString(row["Tracking ID Charge Amount17"]);
            clsPedido.TrackingIDChargeDescription18 = Convert.ToString(row["Tracking ID Charge Description18"]);
            clsPedido.TrackingIDChargeAmount18 = Convert.ToString(row["Tracking ID Charge Amount18"]);
            clsPedido.TrackingIDChargeDescription19 = Convert.ToString(row["Tracking ID Charge Description19"]);
            clsPedido.TrackingIDChargeAmount19 = Convert.ToString(row["Tracking ID Charge Amount19"]);
            clsPedido.TrackingIDChargeDescription20 = Convert.ToString(row["Tracking ID Charge Description20"]);
            clsPedido.TrackingIDChargeAmount20 = Convert.ToString(row["Tracking ID Charge Amount20"]);
            clsPedido.TrackingIDChargeDescription21 = Convert.ToString(row["Tracking ID Charge Description21"]);
            clsPedido.TrackingIDChargeAmount21 = Convert.ToString(row["Tracking ID Charge Amount21"]);
            clsPedido.TrackingIDChargeDescription22 = Convert.ToString(row["Tracking ID Charge Description22"]);
            clsPedido.TrackingIDChargeAmount22 = Convert.ToString(row["Tracking ID Charge Amount22"]);
            clsPedido.TrackingIDChargeDescription23 = Convert.ToString(row["Tracking ID Charge Description23"]);
            clsPedido.TrackingIDChargeAmount23 = Convert.ToString(row["Tracking ID Charge Amount23"]);
            clsPedido.TrackingIDChargeDescription24 = Convert.ToString(row["Tracking ID Charge Description24"]);
            clsPedido.TrackingIDChargeAmount24 = Convert.ToString(row["Tracking ID Charge Amount24"]);
            clsPedido.ShipmentNotes = Convert.ToString(row["Shipment Notes"]);
        }
        
        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleUSPS(DataRow row, ref PedidoUSPS clsPedido )
        {
            clsPedido.GroundService     = Convert.ToString(row["Ground Service"]);
            clsPedido.TrackingNumber    = Convert.ToString(row["Tracking Number"]);
            clsPedido.NetChargeAmount   = Convert.ToString(row["Net Charge Amount"]);
            clsPedido.PODDeliveryDate   = Convert.ToString(row["POD Delivery Date"]);
            clsPedido.RatedWeightAmount = Convert.ToString(row["Rated Weight Amount"]);
            clsPedido.ZoneCode          = Convert.ToString(row["Zone Code"]);
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleUPS(DataRow row, ref PedidoUPS clsPedido)
        {
            clsPedido.Campo1 = Convert.ToString(row[0]);
            clsPedido.Campo2 = Convert.ToString(row[1]);
            clsPedido.Campo3 = Convert.ToString(row[2]);
            clsPedido.Campo4 = Convert.ToString(row[3]);
            clsPedido.Campo5 = Convert.ToString(row[4]);
            clsPedido.Campo6 = Convert.ToString(row[5]);
            clsPedido.Campo7 = Convert.ToString(row[6]);
            clsPedido.Campo8 = Convert.ToString(row[7]);
            clsPedido.Campo9 = Convert.ToString(row[8]);
            clsPedido.Campo10 = Convert.ToString(row[9]);
            clsPedido.Campo11 = Convert.ToString(row[10]);
            clsPedido.Campo12 = Convert.ToString(row[11]);
            clsPedido.Campo13 = Convert.ToString(row[12]);
            clsPedido.Campo14 = Convert.ToString(row[13]);
            clsPedido.Campo15 = Convert.ToString(row[14]);
            clsPedido.Campo16 = Convert.ToString(row[15]);
            clsPedido.Campo17 = Convert.ToString(row[16]);
            clsPedido.Campo18 = Convert.ToString(row[17]);
            clsPedido.Campo19 = Convert.ToString(row[18]);
            clsPedido.Campo20 = Convert.ToString(row[19]);
            clsPedido.Campo21 = Convert.ToString(row[20]);
            clsPedido.Campo22 = Convert.ToString(row[21]);
            clsPedido.Campo23 = Convert.ToString(row[22]);
            clsPedido.Campo24 = Convert.ToString(row[23]);
            clsPedido.Campo25 = Convert.ToString(row[24]);
            clsPedido.Campo26 = Convert.ToString(row[25]);
            clsPedido.Campo27 = Convert.ToString(row[26]);
            clsPedido.Campo28 = Convert.ToString(row[27]);
            clsPedido.Campo29 = Convert.ToString(row[28]);
            clsPedido.Campo30 = Convert.ToString(row[29]);
            clsPedido.Campo31 = Convert.ToString(row[30]);
            clsPedido.Campo32 = Convert.ToString(row[31]);
            clsPedido.Campo33 = Convert.ToString(row[32]);
            clsPedido.Campo34 = Convert.ToString(row[33]);
            clsPedido.Campo35 = Convert.ToString(row[34]);
            clsPedido.Campo36 = Convert.ToString(row[35]);
            clsPedido.Campo37 = Convert.ToString(row[36]);
            clsPedido.Campo38 = Convert.ToString(row[37]);
            clsPedido.Campo39 = Convert.ToString(row[38]);
            clsPedido.Campo40 = Convert.ToString(row[39]);
            clsPedido.Campo41 = Convert.ToString(row[40]);
            clsPedido.Campo42 = Convert.ToString(row[41]);
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosFedex()
        {
            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString+@"\"+ ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios Fedex "+ DateTime.Now.ToString("yyyyMMddTHHmmss")+"\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }

        

            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: "+ Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }
                FlgSihayFedex = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name+" "+ DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PedidoFedex clsPedido = new PedidoFedex();
                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalle(row, ref clsPedido);

                    listaPedido.Add(clsPedido);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: "+ Archivos.Name+" " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString +@"\"+ Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos Fedex";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosUSPS()
        {
            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios USPS " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }



            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayUSPS = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PedidoUSPS clsPedido = new PedidoUSPS();
                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleUSPS(row, ref clsPedido);

                    listaPedidoUSPS.Add(clsPedido);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos USPS";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosUPS()
        {
            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios UPS " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUPS"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayUPS = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PedidoUPS clsPedido = new PedidoUPS();
                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleUPS(row, ref clsPedido);

                    listaPedidoUPS.Add(clsPedido);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos UPS";
            this.Refresh();
            this.Invalidate();
        }

        // Realiza accion de boton unifica reportes
        // ----------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            LocaltextBox1 = textBox1.Text;
            LocaltextBox2 = textBox2.Text;

            // Start BackgroundWorker
            backgroundWorker1.RunWorkerAsync(2000);
        }

        // Put all of background logic that is taking too much time      
        private int BackgroundProcessLogicMethod(BackgroundWorker bw, int a)
        {
            int result = 0;
            try
            {
                pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
                ReporteLog = "ReporteLog" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".txt";

                System.IO.Directory.CreateDirectory(pathString);

                LocaltextBox1 = "Inicio Preparación datos";
                this.Refresh();
                this.Invalidate();

                String fromFile = ConfigurationManager.AppSettings["NombreArchivoBase"];
                string toFile = pathString + @"\ArchivoProcesar.csv";


                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(fromFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // this does not throw exception if file doesnt exist
                File.Delete(toFile);

                wb.SaveAs(toFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, false, Type.Missing, Type.Missing, Type.Missing);

                wb.Close(false, Type.Missing, Type.Missing);

                app.Quit();


                LocaltextBox1 = "Inicio Proceso comparación";
                this.Refresh();
                this.Invalidate();

                //if (!backgroundWorker1.IsBusy)
                //{
                //    backgroundWorker1.RunWorkerAsync();
                //}

                LocaltextBox1 = "Obtener Datos Archivos Fedex";
                this.Refresh();
                this.Invalidate();
                FlgSihayFedex = false;

                ObtieneDatosFedex();

                LocaltextBox1 = "Obtener Datos Archivos USPS";
                this.Refresh();
                this.Invalidate();
                FlgSihayUSPS = false;

                ObtieneDatosUSPS();

                LocaltextBox1 = "Obtener Datos Archivos UPS";
                this.Refresh();
                this.Invalidate();
                FlgSihayUPS = false;

                ObtieneDatosUPS();

                if (FlgSihayFedex == true || FlgSihayUSPS == true)
                {
                    // obtiene datos del excel base
                    // ----------------------------
                    string file = toFile;

                    // obtiene record set
                    // ------------------
                    //progressBar1.Value = 10;
                    //timer1.Enabled = true;

                    LocaltextBox1 = "Comienzo de Carga Archivo Maestro";
                    this.Refresh();
                    this.Invalidate();

                    //// Registra inicio de procesamiento de archivo
                    //// ----------------------------------------
                    using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                    {
                        string Contenido = "Inicio Procesamiento Archivo Maestro: " + file + ": " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                        CreateText.WriteLine(Contenido);
                    }

                    //string connectionString = GetConnectionString(file, "MASTER");
                    // se comenta la lectura del excel para colocar lectura de CSV
                    // ------------------------------------------------------------
                    //var dataSet = GetDataSetFromExcelFile(file, connectionString);

                    //// Registra inicio de procesamiento de archivo
                    //// ----------------------------------------
                    //using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                    //{
                    //    string Contenido = "Cantidad Registros:" + dataSet.Tables("").Rows.Count() + ":" + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    //    CreateText.WriteLine(Contenido);
                    //}

                    // genera encabezado reporte
                    // -------------------------
                    InsertaEncabezadoReporte();

                    LocaltextBox1 = "Inicio de generación de Archivos";
                    this.Refresh();
                    this.Invalidate();
                    int conteoregistros = 0;
                    int conteoCoincidencias = 0;

                    //int counter = 0;
                    //string line;

                    // Read the file and display it line by line.  
                    System.IO.StreamReader fileRead = new System.IO.StreamReader(file);

                    string[] PalabrasUnificadas = new string[46]; ;
                    int contadorPala = 0;
                    while ((line = fileRead.ReadLine()) != null)
                    {
                        cantidad = 0;
                        //string LineaNueva= line.Replace("\",", " ");
                        string[] sa = line.Split(',');
                        if (sa.Length > 46)
                            cantidad = sa.Length - 46;

                        foreach (string s in sa)
                        {

                            if (cantidad > 0 && Encontro == false)
                            {
                                if (s.Length > 0)
                                {
                                    string tipo = s.Substring(0, 1);
                                    if (tipo == "\"")
                                    {
                                        Encontro = true;
                                        PalabraCompleta = s;
                                        continue;
                                    }
                                }
                            }

                            if (cantidad > 0 && Encontro == true)
                            {
                                if (s.Length > 0)
                                {
                                    if (s.Substring(s.Length - 1, 1) == "\"")
                                    {
                                        PalabraCompleta = PalabraCompleta + s;
                                        PalabrasUnificadas[contadorPala] = PalabraCompleta;
                                        contadorPala++;
                                        counter++;
                                        //cantidad = 0;
                                        Encontro = false;
                                        continue;

                                        try
                                        {
                                            PalabraCompleta = PalabraCompleta + s;
                                            PalabrasUnificadas[contadorPala] = PalabraCompleta;
                                            contadorPala++;
                                            counter++;
                                            //cantidad = 0;
                                            Encontro = false;
                                            continue;
                                        }
                                        catch (SystemException exp)
                                        {
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        PalabraCompleta = PalabraCompleta + s;
                                        continue;
                                    }
                                }
                                else
                                {
                                    PalabraCompleta = PalabraCompleta + s;
                                    continue;
                                }
                            }

                            try
                            {
                                PalabrasUnificadas[contadorPala] = s;
                                contadorPala++;
                            }
                            catch (SystemException exp)
                            {
                                continue;
                            }
                        }

                        contadorPala = 0;

                        // obtiene el valor del registro leido
                        // -----------------------------------
                        ObtieneValorRegistro(PalabrasUnificadas);
                        ContadorProgreso = ContadorProgreso + 1;
                        LocaltextBox1 = "Registro procesado: " + TrackingNum;
                        LocaltextBox2 = "Cantidad: " + ContadorProgreso;
                        this.Refresh();
                        this.Invalidate();

                        if (TrackingNum != "")
                        {
                            if (FlgSihayFedex == true)
                            {
                                if (listaPedido.Exists(x => x.FullTrakingId == TrackingNum))
                                {
                                    conteoCoincidencias += 1;
                                    // genera fila de reporte por coincidencia
                                    // ---------------------------------------
                                    InsertaFilaReporte();
                                    EncontroRegistro = true;
                                }
                            }

                            if (FlgSihayUSPS == true && EncontroRegistro == false)
                            {
                                if (listaPedidoUSPS.Exists(x => x.TrackingNumber == TrackingNum))
                                {
                                    conteoCoincidencias += 1;
                                    // genera fila de reporte por coincidencia
                                    // ---------------------------------------
                                    InsertaFilaReporteUSPS();
                                    EncontroRegistro = true;
                                }
                            }

                            if (FlgSihayUPS == true && EncontroRegistro == false)
                            {
                                if (listaPedidoUPS.Exists(x => x.Campo30 == TrackingNum))
                                {
                                    conteoCoincidencias += 1;
                                    // genera fila de reporte por coincidencia
                                    // ---------------------------------------
                                    InsertaFilaReporteUPS();
                                }
                            }

                            EncontroRegistro = false;

                        }
                        conteoregistros += 1;

                    }

                    //// Registra fin de procesamiento de archivo
                    //// ----------------------------------------
                    using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                    {
                        string Contenido = "Archivo Maestro: " + "Contiene: " + conteoregistros + " Registros";
                        CreateText.WriteLine(Contenido);

                        Contenido = "Existen: " + conteoCoincidencias + " Coincidencias";
                        CreateText.WriteLine(Contenido);
                    }

                    LocaltextBox1 = "finaliza de Carga Archivo Maestro";
                    this.Refresh();
                    this.Invalidate();

                    //// Registra inicio de procesamiento de archivo
                    //// ----------------------------------------
                    using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                    {
                        string Contenido = "finaliza Procesamiento Archivo Maestro: " + file + ": " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                        CreateText.WriteLine(Contenido);
                    }

                    //// realiza un archivo tipo Excel con la informacion del reporte
                    //// ------------------------------------------------------------
                    using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
                    {
                        LocaltextBox1 = "Fin Comparación de Archivos";
                        BodyExcel = "</body></html>";

                        FileExcel.WriteLine(BodyExcel);
                        BodyExcel = "";

                        // this does not throw exception if file doesnt exist
                        ///File.Delete(toFile);
                    }

                }
            }
            catch (SystemException exp)
            {
                MessageBox.Show("Error: " + exp.Message);
            }

            return result;
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
 
        }

         private void timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Increment(10);
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {
            ProgressBar Progebar = new ProgressBar();
            

        }

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker helperBW = sender as BackgroundWorker;
            int arg = (int)e.Argument;
            e.Result = BackgroundProcessLogicMethod(helperBW, arg);
            if (helperBW.CancellationPending)
            {
                e.Cancel = true;
            }
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
