using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace Adjust
{
    class BufferFactory
    {
        public static IPipettingInfosGenerator Create(int bufferCnt)
        {
            switch(bufferCnt)
            {
                case 1:
                    return new OneBuffer();
                case 2:
                    return new TwoBuffers();
                case 4:
                    return new FourBuffers();
                default:
                    throw new Exception(string.Format("Buffer Count:{0} is not supported!", bufferCnt));
            }
        }
    }

    interface IPipettingInfosGenerator
    {
        List<PipettingInfo> GetPipettingInfos(List<double> concs);
        
    }


    class FourBuffers:IPipettingInfosGenerator
    {
        const int maxCntPerPlate = 22;
        Dictionary<int, double> specialPosConc = new Dictionary<int, double>();
        
        public List<PipettingInfo> GetPipettingInfos(List<double> concs)
        {
            double concPC = concs.Last();
            concs.RemoveAt(concs.Count - 1);
            List<int> specialIndexs = new List<int>() {22,23,46,47,70,71};
            for( int i = 0; i< concs.Count; i++)
            {
                if (specialIndexs.Contains(i))
                    specialPosConc.Add(i, concs[i]);
            }

            List<PipettingInfo> pipettingInfos = new List<PipettingInfo>();
            int dstPlateCnt = (concs.Count + maxCntPerPlate -1) / maxCntPerPlate;
            for (int dstPlateIndex = 0; dstPlateIndex < dstPlateCnt; dstPlateIndex++ )
            {
                int dstPlateStartIndex = dstPlateIndex * (maxCntPerPlate+2);
                
                int thisPlateCnt = maxCntPerPlate;
                bool isFinalPlate = dstPlateIndex == dstPlateCnt - 1;
                if (isFinalPlate)
                    thisPlateCnt = concs.Count - dstPlateIndex * (maxCntPerPlate+2);
                
                for (int dstWellIndex = 0; dstWellIndex < thisPlateCnt; dstWellIndex++)
                {
                    int colShift = dstWellIndex / 8;
                    int srcWellID = dstPlateStartIndex + dstWellIndex + 1;
                    double conc = concs[srcWellID-1];
                    
                    if (srcWellID > 40) //one blank column
                        srcWellID += 8;
                    int thisColIndex = dstWellIndex - colShift * 8;
                    pipettingInfos.AddRange(GetPipettingInfosThisWell(srcWellID, dstPlateIndex + 1, thisColIndex + 1 + colShift * 32, conc));
                }

                //add the wells be pushed to final plate
                int curPlateCnt = thisPlateCnt; //first, we add normal samples, then we add samples pushed back by NC & PC
                if(isFinalPlate)
                {
                    foreach(KeyValuePair<int,double> pair in specialPosConc)
                    {
                        int colShift = curPlateCnt / 8;
                        int colIndex = curPlateCnt - colShift * 8;
                        pipettingInfos.AddRange(GetPipettingInfosThisWell(pair.Key + 1, dstPlateIndex + 1, colIndex + 1 + colShift * 32, pair.Value));
                        curPlateCnt++;
                    }
                }

                //normal + nc + pc + pushed back
                int pcColShift = curPlateCnt / 8;
                int pcColIndex = curPlateCnt - pcColShift * 8;
                pipettingInfos.AddRange(GetPipettingInfosThisWell(Common.GetWellID("A6"), dstPlateIndex + 1, pcColIndex + 1 + pcColShift * 32, concPC));
                curPlateCnt++;
                int ncColShift = curPlateCnt / 8;
                int ncColIndex = curPlateCnt - ncColShift * 8;
                pipettingInfos.AddRange(GetPipettingInfosThisWell(Common.GetWellID("E6"), dstPlateIndex + 1, ncColIndex + 1 + ncColShift * 32, 0));

            }
            return pipettingInfos;
        }

        private List<PipettingInfo> GetPipettingInfosThisWell(int curWellID,int plateID, int wellID, double conc)
        {
            List<PipettingInfo> pipettingInfos = new List<PipettingInfo>();
            for(int i = 0; i< 4; i++)
            {
                pipettingInfos.Add(new PipettingInfo(curWellID, plateID, wellID + i * 8, conc));
            }
            return pipettingInfos;
        }

    }

    class TwoBuffers : IPipettingInfosGenerator
    {
        const int maxCntPerPlate = 40;
        public List<PipettingInfo> GetPipettingInfos(List<double> concs)
        {
            double concPC = concs.Last();
            concs.RemoveAt(concs.Count - 1);
            List<PipettingInfo> pipettingInfos = new List<PipettingInfo>();
            int plateCnt = (concs.Count + maxCntPerPlate - 1) / maxCntPerPlate;
            for (int plateIndex = 0; plateIndex < plateCnt; plateIndex++)
            {
                int startIndex = plateIndex * maxCntPerPlate;
                int thisPlateCnt = concs.Count;
                while (thisPlateCnt > maxCntPerPlate)
                    thisPlateCnt -= maxCntPerPlate;
                int endIndex = startIndex + thisPlateCnt - 1;
                for (int wellIndex = startIndex; wellIndex <= endIndex; wellIndex++)
                {
                    double conc = concs[wellIndex];
                    int curWellID = wellIndex + 1;
                    if (curWellID > 40) //one blank column
                        curWellID += 8;
                    pipettingInfos.AddRange(GetPipettingInfosThisWell(curWellID, plateIndex + 1, wellIndex - startIndex + 1, conc));
                }
                pipettingInfos.AddRange(GetPCNCPipettings4Plate(plateIndex,concPC));
            }
            return pipettingInfos;
        }

        private List<PipettingInfo> GetPCNCPipettings4Plate(int plateIndex,double concPC)
        {
            List<PipettingInfo> pipettingInfos = new List<PipettingInfo>();
            pipettingInfos.Add(new PipettingInfo(Common.GetWellID("A6"), plateIndex + 1, Common.GetWellID("A6"), concPC));
            pipettingInfos.Add(new PipettingInfo(Common.GetWellID("A6"), plateIndex + 1, Common.GetWellID("A12"), concPC));
            pipettingInfos.Add(new PipettingInfo(Common.GetWellID("E6"), plateIndex + 1, Common.GetWellID("E6"), 0));
            pipettingInfos.Add(new PipettingInfo(Common.GetWellID("E6"), plateIndex + 1, Common.GetWellID("E12"), 0));
            return pipettingInfos;
        }

        private List<PipettingInfo> GetPipettingInfosThisWell(int curWellID, int plateID, int wellID, double conc)
        {
            List<PipettingInfo> pipettingInfos = new List<PipettingInfo>();
            for (int i = 0; i < 2; i++)
            {
                int dstWellID = wellID + i*48;
                pipettingInfos.Add(new PipettingInfo(curWellID, plateID, dstWellID, conc));
            }
            return pipettingInfos;
        }

    }


    class OneBuffer : IPipettingInfosGenerator
    {
        const int maxCntPerPlate = 80;
        public List<PipettingInfo> GetPipettingInfos(List<double> concs)
        {
            double concPC = concs.Last();
            concs.RemoveAt(concs.Count - 1);
            List<PipettingInfo> pipettingInfos = new List<PipettingInfo>();
            int plateCnt = (concs.Count + maxCntPerPlate - 1) / maxCntPerPlate;
           
            int startIndex = 0;
            int thisPlateCnt = concs.Count;

            int endIndex = startIndex + maxCntPerPlate - 1;
            for (int wellIndex = startIndex; wellIndex <= endIndex; wellIndex++)
            {
                double conc = concs[wellIndex];
                int curWellID = wellIndex + 1;
                if (curWellID > 40) //one blank column
                    curWellID += 8;
                pipettingInfos.AddRange(GetPipettingInfosThisWell(curWellID, 1, conc));
            }
            pipettingInfos.AddRange(GetPCNCPipettings4Plate(concPC));
            return pipettingInfos;
        }
        private List<PipettingInfo> GetPCNCPipettings4Plate( double concPC)
        {
            List<PipettingInfo> pipettingInfos = new List<PipettingInfo>();
            int dstStartID = Common.GetWellID("A6");
            double conc = concPC / 8.0;
            for (int i = 0; i < 6; i++ )
            {
                pipettingInfos.Add(new PipettingInfo(Common.GetWellID("A6"), 1, dstStartID++, Math.Round(conc)));
                conc = conc * 2;
            }
            pipettingInfos.Add(new PipettingInfo(Common.GetWellID("E6"), 1, Common.GetWellID("G6"), 0));
            pipettingInfos.Add(new PipettingInfo(Common.GetWellID("E6"), 1, Common.GetWellID("H6"), 0));
            return pipettingInfos;
        }

        private List<PipettingInfo> GetPipettingInfosThisWell(int curWellID, int plateID, double conc)
        {
            List<PipettingInfo> pipettingInfos = new List<PipettingInfo>();

            pipettingInfos.Add(new PipettingInfo(curWellID, plateID, curWellID, conc));
            return pipettingInfos;
        }

    }


 
    class PipettingInfo
    {
        public int srcWellID;
        public int dstPlateID;
        public int dstWellID;
        public double conc;
        public PipettingInfo(int srcWellID, int dstPlateID, int dstWellID, double conc)
        {
            this.srcWellID = srcWellID;
            this.dstPlateID = dstPlateID;
            this.dstWellID = dstWellID;
            this.conc = conc;
        }
    }

    class Common
    {
        public static int rows = 8;
        public static int cols = 12;
        private static int dstConc = int.Parse(ConfigurationManager.AppSettings["dstConc"]);
        private static int totalVolume = int.Parse(ConfigurationManager.AppSettings["totalVolume"]);
        public static int GetWellID(int rowIndex, int colIndex)
        {
            return colIndex * 8 + rowIndex + 1;
        }

        public static string GetWellDesc(int wellID)
        {
            int colIndex = (wellID - 1) / 8;
            int rowIndex = wellID - colIndex * 8 - 1;
            return string.Format("{0}{1}", (char)('A' + rowIndex), colIndex + 1);
        }

        internal static int GetWellID(string sWell)
        {
            int rowIndex = sWell.First() - 'A';
            int colIndex = int.Parse(sWell.Substring(1)) - 1;
            return GetWellID(rowIndex, colIndex);
        }


        public static void Write2File(string sFile,List<PipettingInfo> pipettingInfos, bool addingSample)
        {
            List<string> strs = new List<string>();
            //pipettingInfos.ForEach(x=>strs.Add(Format(x,addingSample)));
            foreach(var pipettingInfo in pipettingInfos)
            {
                string s = Format(pipettingInfo, addingSample);
                if (s != "")
                    strs.Add(s);
            }
            File.WriteAllLines(sFile, strs);
        }

        private static string Format(PipettingInfo pipettingInfo, bool addingSample)
        {
            if (!addingSample && pipettingInfo.conc == 0) //no buffer for nc
                return "";
            string srcLabware = addingSample ? "sample" : "buffer";
            int volume = CalculateVolume(pipettingInfo.conc, addingSample);
            return string.Format("{0},{1},plate{2},{3},{4}",
                                            srcLabware,
                                            GetWellDesc(pipettingInfo.srcWellID),
                                            pipettingInfo.dstPlateID,
                                            GetWellDesc(pipettingInfo.dstWellID),
                                            volume);
        }

        private static int CalculateVolume(double conc, bool addingSample)
        {
            if (conc == 0) // 10ul for nc
                return 10;
            double volume = totalVolume * dstConc / conc;
            if (volume < 1)
                volume = 1;
            if (!addingSample)
                volume = totalVolume - volume;
            return (int)Math.Round(volume);

        }
    }
}
