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
        List<PipettingInfoSimplify> GetPipettingInfos(List<double> concs);
        
    }


    class FourBuffers:IPipettingInfosGenerator
    {
        const int maxCntPerPlate = 22;
        Dictionary<int, double> specialPosConc = new Dictionary<int, double>();
        
        public List<PipettingInfoSimplify> GetPipettingInfos(List<double> concs)
        {
            double concPC = concs.Last();
            concs.RemoveAt(concs.Count - 1);
            int totalSampleCnt = concs.Count;
            List<int> specialIndexs = new List<int>() {22,23,46,47,70,71};
            int totalCnt = concs.Count;
            specialPosConc.Clear();
            for (int i = totalCnt - 1; i >= 0; i--)
            {
                if (specialIndexs.Contains(i))
                {
                    specialPosConc.Add(i, concs[i]);
                    concs.RemoveAt(i);
                }
            }
            int normalCnt = concs.Count;
            concs.AddRange(specialPosConc.Values.Reverse()); //move special position to end
            List<PipettingInfoSimplify> pipettingInfos = new List<PipettingInfoSimplify>();
            
            for( int i = 0; i< concs.Count; i++)
            {
                bool isSpecial = i >= normalCnt;
                int dstPlateIndex = i / maxCntPerPlate;
                int indexInDstPlate = i - dstPlateIndex * maxCntPerPlate;
                int startWellIDInDst = GetStartWellID(indexInDstPlate);
                int srcWellID = i + 1 + dstPlateIndex * 2; //add 2 pushed back
                if( isSpecial)
                {
                    srcWellID = specialIndexs[i-normalCnt] + 1;
                }
                if (srcWellID > 40)//one blank column
                    srcWellID += 8;
                pipettingInfos.AddRange(GetPipettingInfosThisWell(srcWellID, dstPlateIndex + 1, startWellIDInDst, concs[i]));
            }

            int dstPlateCnt = (totalSampleCnt + maxCntPerPlate - 1) / maxCntPerPlate;
            for (int dstPlateIndex = 0; dstPlateIndex < dstPlateCnt; dstPlateIndex++)
            {
                bool isFinalPlate = dstPlateIndex == dstPlateCnt - 1;
                int curPlateCnt = maxCntPerPlate;
                if (isFinalPlate)
                {
                    curPlateCnt = totalSampleCnt - dstPlateIndex * maxCntPerPlate;
                    //normal + nc + pc + pushed back
                    int curIndex = curPlateCnt;
                    int startID = GetStartWellID(curIndex);

                    pipettingInfos.AddRange(GetPipettingInfosThisWell(Common.GetWellID("A6"), dstPlateIndex + 1, startID, concPC));
                    curIndex++;
                    startID = GetStartWellID(curIndex);
                    pipettingInfos.AddRange(GetPipettingInfosThisWell(Common.GetWellID("E6"), dstPlateIndex + 1, startID, 0));
                }
                else
                {
                    pipettingInfos.AddRange(GetPipettingInfosThisWell(Common.GetWellID("A6"), dstPlateIndex + 1,Common.GetWellID("G9") , concPC));
                    pipettingInfos.AddRange(GetPipettingInfosThisWell(Common.GetWellID("E6"), dstPlateIndex + 1, Common.GetWellID("H9"), 0));
                }
               
            }
            return pipettingInfos;
        }

        private int GetStartWellID(int curIndex)
        {
            int colShift = curIndex / 8;
            int colIndex = curIndex - colShift * 8;
            return colIndex + 1 + colShift * 32;
        }

        private List<PipettingInfoSimplify> GetPipettingInfosThisWell(int curWellID,int plateID, int wellID, double conc)
        {
            List<PipettingInfoSimplify> pipettingInfos = new List<PipettingInfoSimplify>();
            for(int i = 0; i< 4; i++)
            {
                int bufferType =  i + 1;
                pipettingInfos.Add(new PipettingInfoSimplify(bufferType, curWellID, plateID, wellID + i * 8, conc));
            }
            return pipettingInfos;
        }

    }

    class TwoBuffers : IPipettingInfosGenerator
    {
        const int maxCntPerPlate = 40;
        public List<PipettingInfoSimplify> GetPipettingInfos(List<double> concs)
        {
            double concPC = concs.Last();
            concs.RemoveAt(concs.Count - 1);
            List<PipettingInfoSimplify> pipettingInfos = new List<PipettingInfoSimplify>();
            int dstPlateCnt = (concs.Count + maxCntPerPlate - 1) / maxCntPerPlate;
            for (int dstPlateIndex = 0; dstPlateIndex < dstPlateCnt; dstPlateIndex++)
            {
              
                int startIndex = dstPlateIndex * maxCntPerPlate;
                int thisPlateCnt = maxCntPerPlate;
                bool isFinalPlate = dstPlateIndex == dstPlateCnt - 1;
                if (isFinalPlate)
                    thisPlateCnt = concs.Count - dstPlateIndex * maxCntPerPlate;

                int endIndex = startIndex + thisPlateCnt - 1;
                for (int wellIndex = startIndex; wellIndex <= endIndex; wellIndex++)
                {
                    double conc = concs[wellIndex];
                    int curWellID = wellIndex + 1;
                    if (curWellID > 40) //one blank column
                        curWellID += 8;
                    pipettingInfos.AddRange(GetPipettingInfosThisWell(curWellID, dstPlateIndex + 1, wellIndex - startIndex + 1, conc));
                }
                pipettingInfos.AddRange(GetPCNCPipettings4Plate(dstPlateIndex, concPC));
            }
            return pipettingInfos;
        }

        private List<PipettingInfoSimplify> GetPCNCPipettings4Plate(int plateIndex, double concPC)
        {
            List<PipettingInfoSimplify> pipettingInfos = new List<PipettingInfoSimplify>();
            int bufferType = plateIndex + 1;
            pipettingInfos.Add(new PipettingInfoSimplify(1,Common.GetWellID("A6"), plateIndex + 1, Common.GetWellID("A6"), concPC));
            pipettingInfos.Add(new PipettingInfoSimplify(2,Common.GetWellID("A6"), plateIndex + 1, Common.GetWellID("A12"), concPC));
            pipettingInfos.Add(new PipettingInfoSimplify(1,Common.GetWellID("E6"), plateIndex + 1, Common.GetWellID("E6"), 0));
            pipettingInfos.Add(new PipettingInfoSimplify(2,Common.GetWellID("E6"), plateIndex + 1, Common.GetWellID("E12"), 0));
            return pipettingInfos;
        }

        private List<PipettingInfoSimplify> GetPipettingInfosThisWell(int curWellID, int plateID, int wellID, double conc)
        {
            List<PipettingInfoSimplify> pipettingInfos = new List<PipettingInfoSimplify>();
            for (int i = 0; i < 2; i++)
            {
                int dstWellID = wellID + i*48;
                pipettingInfos.Add(new PipettingInfoSimplify(i+1,curWellID, plateID, dstWellID, conc));
            }
            return pipettingInfos;
        }

    }


    class OneBuffer : IPipettingInfosGenerator
    {
        const int maxCntPerPlate = 80;
        public List<PipettingInfoSimplify> GetPipettingInfos(List<double> concs)
        {
            double concPC = concs.Last();
            concs.RemoveAt(concs.Count - 1);
            List<PipettingInfoSimplify> pipettingInfos = new List<PipettingInfoSimplify>();
            int plateCnt = (concs.Count + maxCntPerPlate - 1) / maxCntPerPlate;
           
            int startIndex = 0;
            int thisPlateCnt = concs.Count;

            int endIndex = startIndex + thisPlateCnt - 1;
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
        private List<PipettingInfoSimplify> GetPCNCPipettings4Plate( double concPC)
        {
            List<PipettingInfoSimplify> pipettingInfos = new List<PipettingInfoSimplify>();
            int dstStartID = Common.GetWellID("A6");
            //double conc = concPC / 8.0;
            //for (int i = 0; i < 6; i++ )
            //{
            //    pipettingInfos.Add(new PipettingInfo(Common.GetWellID("A6"), 1, dstStartID++, Math.Round(conc)));
            //    conc = conc * 2;
            //}
            pipettingInfos.Add(new PipettingInfoSimplify(1,Common.GetWellID("E6"), 1, Common.GetWellID("G6"), 0));
            pipettingInfos.Add(new PipettingInfoSimplify(1,Common.GetWellID("E6"), 1, Common.GetWellID("H6"), 0));
            return pipettingInfos;
        }

        private List<PipettingInfoSimplify> GetPipettingInfosThisWell(int curWellID, int plateID, double conc)
        {
            List<PipettingInfoSimplify> pipettingInfos = new List<PipettingInfoSimplify>();

            pipettingInfos.Add(new PipettingInfoSimplify(1,curWellID, plateID, curWellID, conc));
            return pipettingInfos;
        }

    }

    class PipettingInfo
    {
        public string srcLabware;
        public int srcWellID;
        public double volume;
        public string dstLabware;
        public int dstWellID;

        public PipettingInfo(string srcLabware, int srcWellID, double volume, string dstLabware, int dstWellID)
        {
            this.srcLabware = srcLabware;
            this.srcWellID = srcWellID;
            this.volume = volume;
            this.dstLabware = dstLabware;
            this.dstWellID = dstWellID;
        }
       
        public PipettingInfo( PipettingInfoSimplify simplify, double vol, string srcLabware)
        {
            this.srcLabware = srcLabware;
            this.srcWellID = simplify.srcWellID;
            this.volume = vol;
            this.dstLabware = string.Format("plate{0}",simplify.dstPlateID);
            this.dstWellID = simplify.dstWellID;
        }
    }
 
    class PipettingInfoSimplify
    {
        public int srcWellID;
        public int dstPlateID;
        public int dstWellID;
        public double conc;
        public int bufferType;

        public PipettingInfoSimplify(int bufferType, int srcWellID, int dstPlateID, int dstWellID, double conc)
        {
            this.srcWellID = srcWellID;
            this.dstPlateID = dstPlateID;
            this.dstWellID = dstWellID;
            this.conc = conc;
            this.bufferType = bufferType;
        }
    }

    class Common
    {
        public static int rows = 8;
        public static int cols = 12;
        private static int dstConc = int.Parse(ConfigurationManager.AppSettings["dstConc"]);
        public static int totalVolume = int.Parse(ConfigurationManager.AppSettings["totalVolume"]);
        private static int ncVolume = int.Parse(ConfigurationManager.AppSettings["ncVolume"]);
        private static int maxTransferVolume = int.Parse(ConfigurationManager.AppSettings["maxTransferVolume"]);
        public static int BufferCount { get; set; }
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


        public static void Write2File(string sFile,List<PipettingInfoSimplify> pipettingInfos,int batchID, bool addingSample)
        {
            List<string> strs = new List<string>();
            foreach(var pipettingInfo in pipettingInfos)
            {
                string s = Format(pipettingInfo,batchID, addingSample);
                if (s != "")
                    strs.Add(s);
            }
            File.WriteAllLines(sFile, strs);
        }

        private static string Format(PipettingInfoSimplify pipettingInfo,int batchID, bool addingSample)
        {
            string srcLabware = addingSample ? string.Format("sample{0}",batchID) : string.Format("buffer{0}",pipettingInfo.bufferType);
            int volume = CalculateVolume(pipettingInfo.conc, addingSample);
            if (volume == 0)
                return "";
            return string.Format("{0},{1},plate{2},{3},{4}",
                                            srcLabware,
                                            GetWellDesc(pipettingInfo.srcWellID),
                                            pipettingInfo.dstPlateID + (batchID-1) * BufferCount,
                                            GetWellDesc(pipettingInfo.dstWellID),
                                            volume);
        }

        private static int CalculateVolume(double conc, bool addingSample)
        {
            
            double volume = totalVolume * dstConc / conc;
            if (conc == 0) // for nc
            {
                volume = ncVolume;
            }
            //limit volume into 1-10
            if (volume < 1)
                volume = 1;
            if (volume > maxTransferVolume)
                volume = maxTransferVolume;
            if (!addingSample)
                volume = totalVolume - volume;
            return (int)Math.Round(volume);

        }

        internal static List<PipettingInfo> Convert2CompletePipettingInfos(List<PipettingInfoSimplify> pipettingInfos, int batchID)
        {
            List<PipettingInfo> completePipettingInfos = new List<PipettingInfo>();
            pipettingInfos.ForEach(x => AddBuffer(x,batchID,completePipettingInfos));
            return completePipettingInfos;
        }

        private static void AddBuffer(PipettingInfoSimplify x,int batchID, List<PipettingInfo> completePipettingInfos)
        {
            x.dstPlateID += (batchID - 1) * BufferCount;
            completePipettingInfos.Add(new PipettingInfo(x, CalculateVolume(x.conc, false), string.Format("buffer{0}", x.bufferType)));
        }
    }
}
