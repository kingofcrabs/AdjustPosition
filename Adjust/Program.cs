using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace Adjust
{
    class Program
    {
        static string concFile = ConfigurationManager.AppSettings["concFile"];
        static void Main(string[] args)
        {
            //try
            {
                Console.WriteLine(strings.version);
                
                DeleteCSVFiles(concFile);
                Console.WriteLine("Converting excel to csv file...");
                Convert2CSV();
                concFile = concFile.Replace(".xlsx", ".csv");
                List<double> concs = ReadConcs(concFile);
                
                Console.WriteLine("Converted successfully.");
                //FourBuffers fourBuffers = new FourBuffers();
                if(args.Length == 0)
                    throw new Exception("No buffer count has been specified!");
                
                int bufferCnt = int.Parse(args[0]);
                if (bufferCnt > 80)
                    throw new Exception("sample count must <= 80!");
                IPipettingInfosGenerator pipettingGenerator = BufferFactory.Create(bufferCnt);
                var simplifyPipettingInfos = pipettingGenerator.GetPipettingInfos(concs);
                simplifyPipettingInfos = simplifyPipettingInfos.OrderBy(x => x.dstPlateID * 100 + x.dstWellID).ToList();
                
                FileInfo fi = new FileInfo(concFile);
                string sDir = fi.Directory.FullName;
                string sSamplePipetting = sDir + "\\sample.csv";
                string sBufferPipetting = sDir + string.Format("\\buffer{0}.csv", bufferCnt);
                string sBufferPipettingGWL = sDir + string.Format("\\buffer{0}.gwl",bufferCnt);
                Common.Write2File(sSamplePipetting, simplifyPipettingInfos, true);
                Common.Write2File(sBufferPipetting, simplifyPipettingInfos, false);
                Worklist wklist = new Worklist();
                var completePipettingInfos = Common.Convert2CompletePipettingInfos(simplifyPipettingInfos);
                File.WriteAllLines(sBufferPipettingGWL, wklist.GenerateGWL(completePipettingInfos));
                Console.WriteLine(string.Format("worklist for sample & buffer has been generated at:{0}", sDir));
                //List<double> testConcs = new List<double>();
                //for (int i = 0; i < 81; i++)
                //{
                //    testConcs.Add(i + 1);
                //}
                //List<PipettingInfo> pipettingInfos = null;
                //string sFile = @"F:\temp\";
                ////FourBuffers fourBuffers = new FourBuffers();
                ////var pipettingInfos = fourBuffers.GetPipettingInfos(testConcs);
                //////pipettingInfos = pipettingInfos.OrderBy(x => x.dstPlateID * 100 + x.dstWellID).ToList();
                ////string sFile = @"F:\temp\";
                ////Common.Write2File(sFile + "fourBuffers.txt", pipettingInfos);

                //TwoBuffers twoBuffers = new TwoBuffers();
                //pipettingInfos = twoBuffers.GetPipettingInfos(testConcs);
                ////pipettingInfos = pipettingInfos.OrderBy(x => x.dstPlateID * 100 + x.dstWellID).ToList();
                //Common.Write2File(sFile + "twoBuffers.txt", pipettingInfos);

                //OneBuffer oneBuffer = new OneBuffer();
                //pipettingInfos = oneBuffer.GetPipettingInfos(testConcs);
                ////pipettingInfos = pipettingInfos.OrderBy(x => x.dstPlateID * 100 + x.dstWellID).ToList();
                //Common.Write2File(sFile + "oneBuffer.txt", pipettingInfos);
            }
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message + ex.StackTrace);
            //}
            Console.WriteLine("Press any key to exit!");
            Console.ReadKey();
        }

        private static void DeleteCSVFiles(string concFile)
        {
            string sCSV = concFile.Replace(".xlsx", ".csv");
            if (File.Exists(sCSV))
                File.Delete(sCSV);
        }

        

        private static List<double> ReadConcs(string concFile)
        {
            List<double> vals = new List<double>();
            var strs = File.ReadAllLines(concFile);
            for(int i = 2; i< strs.Length; i++)
            {
                string[] tmpStrs = strs[i].Split(',');
                if (tmpStrs[2] == "")
                    break;
                vals.Add(double.Parse(tmpStrs[2]));
            }
            return vals;
        }

        internal static void Convert2CSV()
        {
            Console.WriteLine("try to convert the excel to csv format.");
            string concFile = ConfigurationManager.AppSettings["concFile"];
            if(!File.Exists(concFile))
            {
                throw new Exception(string.Format("Cannot find concentration file at: {0}!",concFile));
            }
            List<string> files = new List<string>() { concFile };
            SaveAsCSV(files);
        }

        private static void SaveAsCSV(List<string> sheetPaths)
        {
            Application app = new Application();
            app.Visible = false;
            app.DisplayAlerts = false;
            foreach (string sheetPath in sheetPaths)
            {

                string sWithoutSuffix = "";
                int pos = sheetPath.IndexOf(".xls");
                if (pos == -1)
                    throw new Exception("Cannot find xls in file name!");
                sWithoutSuffix = sheetPath.Substring(0, pos);
                string sCSVFile = sWithoutSuffix + ".csv";
                if (File.Exists(sCSVFile))
                    continue;
                sCSVFile = sCSVFile.Replace("\\\\", "\\");
                Workbook wbWorkbook = app.Workbooks.Open(sheetPath);
                wbWorkbook.SaveAs(sCSVFile, XlFileFormat.xlCSV);
                wbWorkbook.Close();
                Console.WriteLine(sCSVFile);
            }
            app.Quit();
        }
    }



}
