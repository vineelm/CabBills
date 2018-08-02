using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace CabBills
{
    public static class Bill
    {
        public static void Generate()
        {
            string billPath = @"C:\TNow Documents\CabBills\CabBills\HTMLFormats\Ola_GST_OnlyFare_WithDriver.html";
            string excelPath = @"C:\TNow Documents\CabBills\CabBills\ExcelFiles\Ola_GST_OnlyFare.xlsx";

            List<BillData> lstBills = LoadExcel(excelPath);

            string dest = string.Empty;

            foreach (var data in lstBills)
            {
                string billHTML = File.ReadAllText(billPath);

                double d = double.Parse(data.Date);
                DateTime conv = DateTime.FromOADate(d);

                d = double.Parse(data.StartTime);
                DateTime dtStartTime = DateTime.FromOADate(d);

                d = double.Parse(data.EndTime);
                DateTime dtEndTime = DateTime.FromOADate(d);

                billHTML = billHTML.Replace("$CRN$", "CRN" + RandomNumber());
                billHTML = billHTML.Replace("$Day$", String.Format("{0:dddd}", conv));
                billHTML = billHTML.Replace("$TotalBill$", data.TotalBill);
                billHTML = billHTML.Replace("$Destination$", data.To);
                billHTML = billHTML.Replace("$MailedDate$", String.Format("{0:ddd, MMM dd, yyyy}", conv) + " at " + string.Format("{0:hh:mm tt}", conv));
                billHTML = billHTML.Replace("$TravelledDate$", String.Format("{0:dd MMM, yyyy}", conv));
                billHTML = billHTML.Replace("$BaseFare$", data.BaseFare);
                billHTML = billHTML.Replace("$DistanceKms$", data.DistanceKms);
                billHTML = billHTML.Replace("$DistanceFare$", data.DistanceFare);

                billHTML = billHTML.Replace("$RideTime$", data.RideTime);
                billHTML = billHTML.Replace("$RideFare$", data.RideFare);
                billHTML = billHTML.Replace("$AdvanceBooking$", data.AdvanceBooking);

                billHTML = billHTML.Replace("$Taxes$", data.Taxes);
                billHTML = billHTML.Replace("$TotalKms$", Convert.ToString(Convert.ToDecimal(data.DistanceKms) + 4));
                billHTML = billHTML.Replace("$RideTime$", data.RideTime);
                billHTML = billHTML.Replace("$StartTime$", String.Format("{0:hh:mm tt}", dtStartTime));
                billHTML = billHTML.Replace("$EndTime$", String.Format("{0:hh:mm tt}", dtEndTime));

                billHTML = billHTML.Replace("$SourceAddress$", data.SourceAddress);
                billHTML = billHTML.Replace("$DestinationAddress$", data.DestinationAddress);

                if (data.From == "Walamtari" && data.To == "Kothapet")
                {
                    billHTML = billHTML.Replace("$map$", "Maps\\Walamtari_Kothapet.png");
                }
                else if (data.From == "Kothapet" && data.To == "Walamtari")
                {
                    billHTML = billHTML.Replace("$map$", "Maps\\Kothapet_Walamtari.png");
                }
                else if (data.From == "peerancheruvu" && data.To == "Walamtari")
                {
                    billHTML = billHTML.Replace("$map$", "Maps\\Peerancheruvu_Walamtari.png");
                }
                else if (data.From == "Walamtari" && data.To == "peerancheruvu")
                {
                    billHTML = billHTML.Replace("$map$", "Maps\\Walamtari_Peerancheruvu.png");
                }
                else if (data.From == "peerancheruvu - tolichowki" && data.To == "Walamtari")
                {
                    billHTML = billHTML.Replace("$map$", "Maps\\Peerancheruvu_Walamtari_via_tolichowki.png");
                }

                billHTML = billHTML.Replace("$DriverPhoto$", "Photos\\" + data.DriverPhoto);
                billHTML = billHTML.Replace("$DriverName$", data.DriverName);
                billHTML = billHTML.Replace("$CarImage$", "Cars\\" + data.CarImage);
                billHTML = billHTML.Replace("$CarType$", data.CarType);

                dest = @"Bills\" + data.From + "_" + data.To + String.Format("{0:dd_MMM_yyyy}", conv) + ".html";

                using (FileStream fs = new FileStream(dest, FileMode.Create))
                {
                    using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
                    {
                        w.WriteLine(billHTML);
                    }
                }
            }
        }




        private static List<BillData> LoadExcel(string excelPath)
        {
            List<BillData> data = new List<BillData>();

            Application xlApp = null;
            Workbook xlWorkbook = null;
            Worksheet xlWorkSheet = null;

            try
            {
                xlApp = new Application();

                xlWorkbook = xlApp.Workbooks.Open(excelPath);
                xlWorkSheet = xlWorkbook.Sheets[1];
                var xlRange = xlWorkSheet.UsedRange;

                var rowCount = xlRange.Rows.Count;

                for (var i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        var cred = new BillData()
                        {
                            Date = xlRange.Cells[i, 1].Value2 != null ? xlRange.Cells[i, 1].Value2.ToString() : string.Empty,
                            From = xlRange.Cells[i, 2].Value2 != null ? xlRange.Cells[i, 2].Value2.ToString() : string.Empty,
                            To = xlRange.Cells[i, 3].Value2 != null ? xlRange.Cells[i, 3].Value2.ToString() : string.Empty,
                            RideFare = xlRange.Cells[i, 4].Value2 != null ? xlRange.Cells[i, 4].Value2.ToString() : string.Empty,
                            AdvanceBooking = xlRange.Cells[i, 5].Value2 != null ? xlRange.Cells[i, 5].Value2.ToString() : string.Empty,
                            Taxes = xlRange.Cells[i, 6].Value2 != null ? xlRange.Cells[i, 6].Value2.ToString() : string.Empty,
                            TotalBill = xlRange.Cells[i, 7].Value2 != null ? xlRange.Cells[i, 7].Value2.ToString() : string.Empty,
                            StartTime = xlRange.Cells[i, 8].Value2 != null ? xlRange.Cells[i, 8].Value2.ToString() : string.Empty,
                            EndTime = xlRange.Cells[i, 9].Value2 != null ? xlRange.Cells[i, 9].Value2.ToString() : string.Empty,
                            SourceAddress = xlRange.Cells[i, 10].Value2 != null ? xlRange.Cells[i, 10].Value2.ToString() : string.Empty,
                            DestinationAddress = xlRange.Cells[i, 11].Value2 != null ? xlRange.Cells[i, 11].Value2.ToString() : string.Empty,
                            DriverPhoto = xlRange.Cells[i, 12].Value2 != null ? xlRange.Cells[i, 12].Value2.ToString() : string.Empty,
                            DriverName = xlRange.Cells[i, 13].Value2 != null ? xlRange.Cells[i, 13].Value2.ToString() : string.Empty,
                            CarImage = xlRange.Cells[i, 14].Value2 != null ? xlRange.Cells[i, 14].Value2.ToString() : string.Empty,
                            CarType = xlRange.Cells[i, 15].Value2 != null ? xlRange.Cells[i, 15].Value2.ToString() : string.Empty,
                        };
                        data.Add(cred);
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkbook);
                // The Quit is done in the finally because we always
                // want to quit. It is no different than releasing RCWs.
                if (xlApp != null)
                {
                    xlApp.Quit();
                }
                releaseObject(xlApp);
            }

            return data;
        }

        private static void releaseObject(object obj) // note ref!
        {
            // Do not catch an exception from this.
            // You may want to remove these guards depending on
            // what you think the semantics should be.
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
            // Since passed "by ref" this assingment will be useful
            // (It was not useful in the original, and neither was the
            //  GC.Collect.)
            obj = null;
        }

        private static string RandomNumber()
        {
            var random = new Random();
            string s = string.Empty;
            for (int i = 0; i < 9; i++)
                s = String.Concat(s, random.Next(10).ToString());
            return s;
        }
    }
}
