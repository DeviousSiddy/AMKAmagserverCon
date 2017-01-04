using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace AMKAmagserverCon
{
    class Program
    {
        const int PORT_NO = 4444;
        const string SERVER_IP = "127.0.0.1";
        bool batch2 = false;

        
        static void Main(string[] args)
        {
            
            IPHostEntry host;
            string localIP = "?";
            host = Dns.GetHostEntry(Dns.GetHostName());
            String[] msgPacket = new String[] { "0", "in", "Amkamagwerker", "", "0" }; // 0 = ID#, 1 = in(0)/out(1), 2 = personeel, 3 = klant, 4 = mode (0 = normal, 1 = request info list, 2 = Continuous mode, 3 = Correction,  4 = close connection)

            ExcelInit();
            foreach (IPAddress ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    localIP = ip.ToString();
                }
            }
            Console.WriteLine("Sever IP: " + localIP);
            Bitmap qrcode = GenerateQR(300,300,localIP);
            if (File.Exists(@"C:\AMKA\qrip.jpeg"))
            {
                File.Delete(@"C:\AMKA\qrip.jpeg");
            }
            
            qrcode.Save(@"C:\AMKA\qrip.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);
            //Process.Start(@"C:\AMKA\qrip.jpeg");

            //---listen at the specified IP and port no.---
            IPAddress localAdd = IPAddress.Parse(localIP);
            TcpListener listener = new TcpListener(localAdd, PORT_NO);
            Console.WriteLine("Listening...\n");
            listener.Start();
            TcpClient client = null;
            client = listener.AcceptTcpClient();

            while ((true))
            {
                //---incoming client connected---

               

                    if (!client.Connected)
                    {

                        client.Close();
                        client = listener.AcceptTcpClient();
                        Console.WriteLine("Listening...\n");
                    }

                
                //---get the incoming data through a network stream---
                NetworkStream nwStream = client.GetStream();
                byte[] buffer = new byte[client.ReceiveBufferSize];

                //---read incoming stream---
                int bytesRead = nwStream.Read(buffer, 0, client.ReceiveBufferSize);
                nwStream.Flush();

                //---convert the data received into a string---
                string dataReceived = Encoding.ASCII.GetString(buffer, 0, bytesRead);
                msgPacket = dataReceived.Split('|');
                if (msgPacket.Length == 5)
                {
                    
                    string s = msgPacket[4];
                    Console.Write("Mode: "+s);
                    int i = Int32.Parse(s);
                    switch (i)
                        {
                            case 0:
                            string[] itemsi = msgPacket[0].Split(':');
                            Console.WriteLine("Received : " + msgPacket[0] + " " + msgPacket[1] + " " + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt") + " door " + msgPacket[2]);
                                //buffer = Encoding.ASCII.GetBytes("Received.");
                                //nwStream.Write(buffer, 0, buffer.Length);
                                nwStream.Write(buffer, 0, bytesRead);
                            //nwStream.Flush();
                            String lastitem = "";
                                //if (msgPacket[4].Contains("0"))
                                {
                                    lastitem = Excelwrite(itemsi[0], msgPacket[1], Int32.Parse(itemsi[1]));
                                }
                                //---write back the text to the client---

                                lastitem = lastitem + " " + msgPacket[1];
                                //buffer = Encoding.ASCII.GetBytes(lastitem);
                                //nwStream = client.GetStream();
                                //nwStream.Write(buffer, 0, bytesRead);
                                Console.WriteLine("Sending back : " + lastitem + " \n");
                                //---get the incoming data through a network stream---
                                break;
                        case 2: //TODO fix nwstream
                            string[] items = msgPacket[0].Split(';');
                            Console.WriteLine(items[0]);
                            string lastitem2 = "";
                            string[] item;
                            for (int x = 0; x < items.Length; x++)
                            {
                                item = items[x].Split(':');
                                Console.WriteLine("Received : " +item[1]+"x "+ item[0] + " " + msgPacket[1] + " " + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt") + " door " + msgPacket[2]);
                                buffer = Encoding.ASCII.GetBytes("Received.");
                                
                                nwStream.Write(buffer, 0, buffer.Length);
                                //nwStream.Write(buffer, 0, bytesRead);
                                //nwStream.Flush();
                                //if (msgPacket[4].Contains("0"))
                                {
                                    lastitem2 = Excelwrite(item[0], msgPacket[1], msgPacket[3] ,Int32.Parse(item[1]));
                                }
                                //---write back the text to the client---

                                lastitem2 = lastitem2 + " " + msgPacket[1];
                                buffer = Encoding.ASCII.GetBytes(lastitem2);
                                nwStream = client.GetStream();
                                nwStream.Write(buffer, 0, buffer.Length);
                                nwStream.Flush();
                                Console.WriteLine("Sending back : " + lastitem2 + " \n");
                                //---get the incoming data through a network stream---

                            }
                            
                            break;
                            case 4:
                                
                                client.Close();
                                listener.Stop();
                            Console.WriteLine("Connection closed.");
                            System.Threading.Thread.Sleep(25);
                                listener = new TcpListener(localAdd, PORT_NO);
                            listener.Start();
                                client = listener.AcceptTcpClient();
                            Console.WriteLine("Listening...\n");
                            break;
                            default:
                                break;
                        }
                }
                   
                }
            client.Close();
            listener.Stop();
            Console.ReadLine();



        }
        public static string Excelwrite(String idnum, String checkinout)
        {
            return Excelwrite(idnum,checkinout,"AMKA",1);
        }
        public static string Excelwrite(String idnum, String checkinout,int aantal)
        {
            return Excelwrite(idnum, checkinout, "AMKA", aantal);
        }
        public static string Excelwrite(String idnum, String checkinout,String klant_naam, int aantal)
        {
            
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Application xlAppin = null;
            Excel.Application xlAppuit = null;
            string la = "";

            Boolean inout = true; //true if in, false if out
            if (checkinout.Contains("in"))
            {
                inout = true;
            }
            else if (checkinout.Contains("out"))
            {
                inout = false;
                
            }

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return "Error";
            }
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Workbook xlWorkBookIn = null;
            Excel.Worksheet xlWorkSheetin = null;
            Excel.Workbook xlWorkBookUit = null;
            Excel.Worksheet xlWorkSheetuit = null;
            object misValue = System.Reflection.Missing.Value;
            string curFile = @"c:\AMKA\totaalmag.xls";
            string curFilein = @"c:\AMKA\In\00000000StandaardIn.xls";
            string curFileuit = @"c:\AMKA\Uit\00000000StandaardUit.xls";
            string in_xcel = @"c:\AMKA\In\" + DateTime.Now.ToString("yyyyMMdd") + "In.xls";
            string upperklant =  klant_naam.ToUpper();
            string uit_dir = @"c:\AMKA\Uit\Klant\" + upperklant + @"\";
            string uit_xcel = uit_dir + DateTime.Now.ToString("yyyyMMdd") + "Uit.xls";
            

            if (!File.Exists(curFile) || !File.Exists(curFilein) || !File.Exists(curFileuit))
            {
                ExcelInit();
            }
            if (inout)
            {
                if (File.Exists(in_xcel))
                {
                    curFilein = in_xcel;
                }
            }
            else
            {
                if (!Directory.Exists(uit_dir))
                {
                    Directory.CreateDirectory(uit_dir);
                }
                if (File.Exists(uit_xcel))
                {
                    curFileuit = uit_xcel;
                }
            }

            xlWorkBook = xlApp.Workbooks.Open(curFile);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range currentFind = null;
            // Excel.Range firstFind = null;
            Excel.Range editframe = null;
            // Excel.Range columnRange = null;
            //  Excel.Range rowRange = null;
            Excel.Range currentFindin = null;
            Excel.Range currentFinduit = null;
            Excel.Range editframein = null;
            Excel.Range editframeuit = null;
            Excel.Range lara = null;
            //int numberOfColumns = 0;
            //int numberOfRows = 0;
            Excel.Range Items = xlApp.get_Range("A1", "A300");
            Excel.Range Itemsin = null;
            Excel.Range Itemsuit = null;

            if (inout)
            {

                xlAppin = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBookIn = xlAppin.Workbooks.Open(curFilein);
                xlWorkSheetin = (Excel.Worksheet)xlWorkBookIn.Worksheets.get_Item(1);
                Itemsin = xlAppin.get_Range("A1", "A200");

            }
            else
            {

                xlAppuit = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBookUit = xlAppuit.Workbooks.Open(curFileuit);
                xlWorkSheetuit = (Excel.Worksheet)xlWorkBookUit.Worksheets.get_Item(1);
                Itemsuit = xlAppuit.get_Range("A1", "A200");

            }


            if (idnum != null)
            {
                try
                {
                    Regex regexObj = new Regex(@"[^\d]");
                    idnum = regexObj.Replace(idnum, "");
                }
                catch (ArgumentException ex)
                {
                    // Syntax error in the regular expression
                }
                currentFind = Items.Find(idnum, misValue,
                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                false, misValue);
                if (inout)
                {
                    currentFindin = Itemsin.Find(idnum, misValue,
                    Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                    false, misValue);
                }
                else
                {
                    currentFinduit = Itemsuit.Find(idnum, misValue,
                    Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                    false, misValue);
                }
            }

            //while (currentFind != null)
            //{
            // Keep track of the first range you find. 
            // if (firstFind == null)
            // {
            //     firstFind = currentFind;
            // }

            // If you didn't move to a new range, you are done.
            // else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
            //      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
            // {
            //break;
            // }
            lara = currentFind.Offset[0, 1];
            la = lara.Value2;
            editframe = currentFind.Offset[0, 2];


            Console.WriteLine(currentFind.Value2 + " " + currentFind.Offset[0, 1].Value2);



            // try
            // {

            // editframe = xlApp.Selection as Excel.Range;
            // columnRange = editframe.Columns;
            // rowRange = editframe.Rows;
            // numberOfColumns = columnRange.Count;
            // numberOfRows = rowRange.Count;

            //for (long iRow = 1; iRow <= numberOfRows; iRow++)
            //{
            //    for (long iCol = 1; iCol <= numberOfColumns; iCol++)
            //    {
            //Put the row and column address in the cell.
            if (inout)
            {
                Console.WriteLine("+" + aantal);
                
                editframe.Value2 = editframe.Value2 + aantal;
                editframein = currentFindin.Offset[0, 2];
                editframein.Value2 = editframein.Value2 + aantal;
                Console.WriteLine("Totaal in magazijn: " + editframe.Value2);
                Console.WriteLine(DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt"));

            }
            else
            {
                Console.WriteLine("-" + aantal);
                
                editframe.Value2 = editframe.Value2 - aantal;
                editframeuit = currentFinduit.Offset[0, 2];
                editframeuit.Value2 = editframeuit.Value2 + aantal;
                Console.WriteLine("Totaal in magazijn: "+ editframe.Value2);
                Console.WriteLine(DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt"));
            }
            //editframe[iRow,iCol].Value2 = editframe[iRow, iCol].Value2 + 1.0;
            //   }
            //}


            //}
            //finally
            // {
            // if (rowRange != null) Marshal.ReleaseComObject(rowRange);
            if (currentFind != null) Marshal.ReleaseComObject(currentFind);
            if (currentFindin != null) Marshal.ReleaseComObject(currentFindin);
            if (currentFinduit != null) Marshal.ReleaseComObject(currentFinduit);
            //if (firstFind != null) Marshal.ReleaseComObject(firstFind);
             //       if (columnRange != null) Marshal.ReleaseComObject(columnRange);
            if (editframe != null) Marshal.ReleaseComObject(editframe);
            if (editframein != null) Marshal.ReleaseComObject(editframein);
            if (editframeuit != null) Marshal.ReleaseComObject(editframeuit);
            // }

            // currentFind = Items.FindNext(currentFind);
            //}
            //xlWorkBook.SaveAs(@"c:\AMKA\totaalmag.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Save();
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            if (inout)
            {
                if (!File.Exists(in_xcel))
                {
                    xlWorkBookIn.SaveAs(in_xcel, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                else
                {
                    xlWorkBookIn.Save();
                }
                xlWorkBookIn.Close(true, misValue, misValue);
                xlAppin.Quit();
                Marshal.ReleaseComObject(xlWorkSheetin);
                Marshal.ReleaseComObject(xlWorkBookIn);
                Marshal.ReleaseComObject(xlAppin);
            }
            else
            {
                if (!File.Exists(uit_xcel))
                {
                    
                    xlWorkBookUit.SaveAs(uit_xcel, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                else
                {
                    xlWorkBookUit.Save();
                }
                xlWorkBookUit.Close(true, misValue, misValue);
                xlAppuit.Quit();
                Marshal.ReleaseComObject(xlWorkSheetuit);
                Marshal.ReleaseComObject(xlWorkBookUit);
                Marshal.ReleaseComObject(xlAppuit);
            }


            

            Console.WriteLine("Excel files edited.\n");

            return la;

        }
        public static Bitmap GenerateQR(int width, int height, string text)
        {
            var bw = new ZXing.BarcodeWriter();
            var encOptions = new ZXing.Common.EncodingOptions() { Width = width, Height = height, Margin = 0 };
            bw.Options = encOptions;
            bw.Format = ZXing.BarcodeFormat.QR_CODE;
            var result = new Bitmap(bw.Write(text));

            return result;
        }
        public static void ExcelInit()
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (!Directory.Exists(@"c:\AMKA\"))
            {
                Directory.CreateDirectory(@"c:\AMKA\");
                Directory.CreateDirectory(@"c:\AMKA\In");
                Directory.CreateDirectory(@"c:\AMKA\Uit");
                Directory.CreateDirectory(@"c:\AMKA\Uit\Klant");
                Directory.CreateDirectory(@"c:\AMKA\QRcode");
                Console.WriteLine("Created Directory");
            }

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            string curFile = @"c:\AMKA\totaalmag.xls";
            string curFilein = @"c:\AMKA\In\00000000StandaardIn.xls";
            string curFileuit = @"c:\AMKA\Uit\00000000StandaardUit.xls";

            if (!File.Exists(curFile)|| !File.Exists(curFilein) || !File.Exists(curFileuit))
            {
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "ID";     xlWorkSheet.Cells[1, 2] = "Item";                   xlWorkSheet.Cells[1, 3] = "Pakken";
                xlWorkSheet.Cells[2, 1] = "ST0001";   xlWorkSheet.Cells[2, 2] = "Cola Stroop 350ml";    xlWorkSheet.Cells[2, 3] = "0";
                xlWorkSheet.Cells[3, 1] = "AZ0002";   xlWorkSheet.Cells[3, 2] = "Azijn 350ml";          xlWorkSheet.Cells[3, 3] = "0";
                xlWorkSheet.Cells[4, 1] = "KJ0003";   xlWorkSheet.Cells[4, 2] = "Ketjap 350ml";         xlWorkSheet.Cells[4, 3] = "0";
                if (!File.Exists(curFile))
                {
                   xlWorkBook.SaveAs(curFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                if (!File.Exists(curFilein))
                {
                    xlWorkBook.SaveAs(curFilein, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                if (!File.Exists(curFileuit))
                {
                    xlWorkBook.SaveAs(curFileuit, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                for (int i = 2; i<=200; i++)
                {
                    if (xlWorkSheet.Cells[i, 1].Value != null)
                    {
                        //if (Regex.IsMatch(xlWorkSheet.Cells[i, 1].Value, "^[A - Z]{ 2}\\d{ 4}")) //TODO regex check + create missing QR.
                        {

                            String s = xlWorkSheet.Cells[i, 1].Value;
                            String s2 = xlWorkSheet.Cells[i, 2].Value;

                            Bitmap qrcode = GenerateQR(300, 300, s);
                            if (!File.Exists(@"c:\AMKA\QRcode\" + s + " " + s2 + ".jpeg"))
                            {
                                qrcode.Save(@"c:\AMKA\QRcode\" + s + " " + s2 + ".jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);
                            }
                        }
                    }
                }
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
            



        }

    }
}
