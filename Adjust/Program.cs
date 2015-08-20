﻿using Microsoft.Office.Interop.Excel;
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
#if DEBUG
#else
            try
#endif
            {
                if (args.Length == 0)
                    throw new Exception("No buffer count has been specified!");
                Console.WriteLine(strings.version);
                DeleteCSVFiles(concFile);
                Console.WriteLine("Converting excel to csv file...");
                Convert2CSV();
                concFile = concFile.Replace(".xlsx", ".csv");
                List<double> concs = ReadConcs(concFile);
                Console.WriteLine("Converted successfully.");
                int bufferCnt = int.Parse(args[0]);
                if (bufferCnt <= 0)
                    throw new Exception("sample count must > 0!");
                Common.BufferCount = bufferCnt;
                Console.WriteLine(string.Format("buffer type count is :{0}", bufferCnt));
                FileInfo fi = new FileInfo(concFile);
                string sDir = fi.Directory.FullName;
                IPipettingInfosGenerator pipettingGenerator = BufferFactory.Create(bufferCnt);

                double pcConc = concs.Last();
                concs.RemoveAt(concs.Count - 1);
                int batchID = 0;
                while(concs.Count > 0)
                {
                    var thisBatchConcs = concs.Take(80).ToList();
                    batchID++;
                    concs = concs.Skip(thisBatchConcs.Count).ToList();
                    thisBatchConcs.Add(pcConc);
                    var simplifyPipettingInfos = pipettingGenerator.GetPipettingInfos(thisBatchConcs);
                    simplifyPipettingInfos = simplifyPipettingInfos.OrderBy(x => x.dstPlateID * 10000 + x.dstWellID).ToList();
                    string subFolder = sDir + string.Format("\\batch{0}\\",batchID);
                    if (!Directory.Exists(subFolder))
                        Directory.CreateDirectory(subFolder);
                    string sSamplePipetting = subFolder + string.Format("\\sample.csv", batchID);
                    string sBufferPipetting = subFolder + string.Format("\\buffer{0}.csv", bufferCnt);
                    string sBufferPipettingGWL = subFolder + string.Format("\\buffer{0}.gwl", bufferCnt);
                    Common.Write2File(sSamplePipetting, simplifyPipettingInfos,batchID, true);
                    Common.Write2File(sBufferPipetting, simplifyPipettingInfos,batchID, false);
                    Worklist wklist = new Worklist();
                    var completePipettingInfos = Common.Convert2CompletePipettingInfos(simplifyPipettingInfos,batchID);
                    File.WriteAllLines(sBufferPipettingGWL, wklist.GenerateGWL(completePipettingInfos));
                    Console.WriteLine(string.Format("worklist for sample&buffer for batch: {0} has been generated at:{1}",batchID, subFolder));
                }
                File.WriteAllText(sDir + "\\totalVolume.txt", Common.totalVolume.ToString());
                File.WriteAllText(sDir + "\\batchCount.txt", (batchID-1).ToString());
                
            }
#if DEBUG
#else
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
            }
#endif
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
