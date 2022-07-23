using System;
using IronXL;
using Jayrock;
using System.IO;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.Collections;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace Neurosky_Signal_Aquisition
{
    class Program
    {
        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy,MM,dd,HH,mm,ssffff");
            
        }
        static void Main(string[] args)
        {
            WorkBook workbook = WorkBook.Load("Example.xlsx");
            WorkSheet sheet = workbook.DefaultWorkSheet;
            WorkSheet sheet2 = workbook.DefaultWorkSheet;
            Range range = sheet["A2:M100"];
            Range range2 = sheet2["A2:M100"];

            IDictionary eegPower;
            IDictionary eSense;

            TcpClient client;
            TcpClient client2;

            Stream stream;
            Stream stream2;

            int r = 1;

            byte[] buffer = new byte[4096];
            byte[] buffer2 = new byte[4096];
            int bytesRead; // Building command to enable JSON output from ThinkGear Connector (TGC)  -> EEG Data
            int bytesRead2; // Building command to enable JSON output from ThinkGear Connector (TGC) -> Raw Data

            Console.WriteLine("Do you want normal EEG Signal data?(1/0) default is raw data.");

            var option = Console.ReadLine();

            var com = @"{""enableRawOutput"": true, ""format"": ""Json""}";
            var com2 = @"{""enableRawOutput"": true, ""format"": ""Json""}";

            if (option.ToString() == "1")
            {
                com = @"{""enableRawOutput"": false, ""format"": ""Json""}";
                com2 = @"{""enableRawOutput"": true, ""format"": ""Json""}";
            }
            else
                com = @"{""enableRawOutput"": true, ""format"": ""Json""}";

            byte[] myWriteBuffer = Encoding.ASCII.GetBytes(com);
            byte[] myWriteBuffer2 = Encoding.ASCII.GetBytes(com2);

            try
            {
                Console.WriteLine("Starting connection to Mindwave Mobile Headset.");
                client = new TcpClient("127.0.0.1", 13854);
                client2 = new TcpClient("127.0.0.1", 13854);
                stream = client.GetStream();
                stream2 = client2.GetStream();
                System.Threading.Thread.Sleep(500);
                client.Close();
                client2.Close();
                Console.WriteLine("Step 1 completed!!!");
            }
            catch (SocketException se)
            {
                Console.WriteLine("Error connecting to device.");
            }


            try
            {
                client = new TcpClient("127.0.0.1", 13854);
                client2 = new TcpClient("127.0.0.1", 13854);
                stream = client.GetStream();
                stream2 = client2.GetStream();

                Console.WriteLine("Sending configuration packet to device.");
                if (stream.CanWrite)
                    stream.Write(myWriteBuffer, 0, myWriteBuffer.Length);
                if (stream2.CanWrite)
                    stream2.Write(myWriteBuffer2, 0, myWriteBuffer2.Length);

                System.Threading.Thread.Sleep(500);
                client.Close();
                client2.Close();

                Console.WriteLine("Step 2 completed!!!");
            }

            catch (SocketException se)
            {
                Console.WriteLine("Error sending configuration packet to TGC.");
            }

            try
            {
                Console.WriteLine("Starting data collection.");

                client = new TcpClient("127.0.0.1", 13854);
                client2 = new TcpClient("127.0.0.1", 13854);

                stream = client.GetStream();
                stream2 = client2.GetStream();

                // Sending configuration packet to TGC                
                if (stream.CanWrite)
                    stream.Write(myWriteBuffer, 0, myWriteBuffer.Length);
                if (stream2.CanWrite)
                    stream2.Write(myWriteBuffer2, 0, myWriteBuffer2.Length);


                if (stream.CanRead && stream2.CanRead)
                {
                    //to check if device is ready
                    var ready = false;
                    var startRead = false;

                    //to note keyboard key press and note key press
                    Console.WriteLine("Enter any key to start.");
                    Console.WriteLine("Reading bytes");

                    Console.WriteLine("Enter Folder name");
                    var folName=Console.ReadLine();
                    Console.WriteLine("Enter session number");
                    var sessionNumber = Console.ReadLine();

                    Console.WriteLine("Enter CM or CB or N or T");
                    var op = Console.ReadLine();

                    while (true)
                    {
                        bytesRead = stream.Read(buffer, 0, 4096);
                        bytesRead2 = stream2.Read(buffer2, 0, 4096);

                        int i = 0;

                        string[] packets = Encoding.UTF8.GetString(buffer, 0, bytesRead).Split('\r');
                        string[] packets2 = Encoding.UTF8.GetString(buffer2, 0, bytesRead2).Split('\r');

                        foreach (string s2 in packets2)
                        {
                            string s = packets[i];
                            try
                            {
                                IDictionary data = Jayrock.Json.Conversion.JsonConvert.Import(typeof(IDictionary), s) as IDictionary;
                                IDictionary dataRaw = Jayrock.Json.Conversion.JsonConvert.Import(typeof(IDictionary), s2) as IDictionary;

                                //Check if device is ON/OFF
                                if (data.Contains("status"))
                                {
                                    Console.WriteLine("Device is Off.");
                                    ready = false;
                                    break;
                                }

                                //Check fitting (device on head or not)
                                if (data.Contains("eSense"))
                                {
                                    if (data["eSense"].ToString() == "{\"attention\":0,\"meditation\":0}")
                                    {
                                        Console.WriteLine("Check fitting.");
                                        ready = false;
                                        break;
                                    }
                                }

                                //check if device is ready
                                if (data.Contains("eegPower") && (ready == false) && dataRaw.Contains("rawEeg"))
                                {
                                    IDictionary d = (IDictionary)data["eSense"];
                                    if ((d["attention"].ToString() != "0") && (d["meditation"].ToString() != "0"))
                                    {
                                        ready = true;
                                        //Console.WriteLine("Device is ready.");
                                        Console.WriteLine("");
                                    }
                                    else
                                    {
                                        ready = false;
                                        break;
                                    }
                                }
                                else
                                {
                                    ready = false;
                                    break;
                                }

                                //start data reading only when device is ready.
                                if (ready)
                                {
                                    Console.WriteLine(data);
                                    Console.WriteLine(dataRaw);

                                    eSense = (IDictionary)data["eSense"];
                                    eegPower = (IDictionary)data["eegPower"];
                                    // dataRaw = (IDictionary)data["rawEeg"];


                                    String timeStamp = GetTimestamp(DateTime.Now);

                                    Console.WriteLine(timeStamp);

                                    sheet.SetCellValue(r, 0, eSense["attention"].ToString());
                                    sheet.SetCellValue(r, 1, eSense["meditation"].ToString());
                                    sheet.SetCellValue(r, 2, eegPower["delta"].ToString());
                                    sheet.SetCellValue(r, 3, eegPower["theta"].ToString());
                                    sheet.SetCellValue(r, 4, eegPower["lowAlpha"].ToString());
                                    sheet.SetCellValue(r, 5, eegPower["highAlpha"].ToString());
                                    sheet.SetCellValue(r, 6, eegPower["lowBeta"].ToString());
                                    sheet.SetCellValue(r, 7, eegPower["highBeta"].ToString());
                                    sheet.SetCellValue(r, 8, eegPower["lowGamma"].ToString());
                                    sheet.SetCellValue(r, 9, eegPower["highGamma"].ToString());
                                    sheet.SetCellValue(r, 10, dataRaw["rawEeg"].ToString());
                                    sheet.SetCellValue(r, 12, timeStamp);

                                    
                                    if (!Directory.Exists(folName))
                                    {
                                        Directory.CreateDirectory(folName);
                                    }
                                    var fullpath = "";

                                    if (op.ToString() == "cm" || op.ToString() == "CM")
                                    {
                                        sheet.SetCellValue(r, 11, "CheatingWithMobile");

                                        fullpath = $"./{folName}/CheatingWithMobile{sessionNumber}.xlsx";
                                        sheet.SaveAs(fullpath);
                                        

                                    }
                                    if (op.ToString() == "cb" || op.ToString() == "CB")
                                    {
                                        sheet.SetCellValue(r, 11, "CheatingWithBook");

                                        fullpath = $"./{folName}/CheatingWithBook{sessionNumber}.xlsx";
                                        sheet.SaveAs(fullpath);


                                    }
                                    if (op.ToString() == "n" || op.ToString() == "N")
                                    {
                                        sheet.SetCellValue(r, 11, "NotCheating");

                                        fullpath = $"./{folName}/notCheating{sessionNumber}.xlsx";
                                        sheet.SaveAs(fullpath);
   
                                       
                                    }

                                    r++;
                                    i++;
                                }
                            }
                            catch (Exception e)
                            {
                            }

                        }
                    }
                }

                //System.Threading.Thread.Sleep(100);
                client.Close();
                client2.Close();
            }
            catch (SocketException se)
            {
                Console.WriteLine("Error in data collection.");
            }
        }
    }
}