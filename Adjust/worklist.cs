using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Adjust
{
    class Worklist
    {
        int tipCountLiha = int.Parse(ConfigurationManager.AppSettings["tipCount"]);
        public List<string> GenerateGWL( List<PipettingInfo> pipettingInfos)
        {
            int tipVol = int.Parse(ConfigurationManager.AppSettings["bufferTipVolume"]);
            List<string> scripts = new List<string>();
            scripts.Add(string.Format("C;Pipetting with tip:{0}", tipVol));
            
          
            var allBufferPipettingInfos = pipettingInfos.GroupBy(x => x.srcLabware).ToList();
            foreach(var thisBufferPipettingInfo in allBufferPipettingInfos)
            {
                var tmpInfos = thisBufferPipettingInfo.ToList();
                AddScript4ThisBuffer(scripts, tmpInfos);
            }
            return scripts;
        }

        private void AddScript4ThisBuffer(List<string> scripts, List<PipettingInfo> tmpInfos)
        {
            int tipReusedTimes = 0;
            int maxTipReuseTimes = int.Parse(ConfigurationManager.AppSettings["reuseTimes"]);
            List<List<PipettingInfo>> allTipPipettingInfos = SplitByTipIndex(tmpInfos);
            foreach (var thisTipPipettingInfos in allTipPipettingInfos)
            {
                tipReusedTimes = 0;
                bool bNeedWash = false;
                foreach (var pipettingInfo in thisTipPipettingInfos)
                {
                    scripts.Add(GetAspirate(pipettingInfo.srcLabware, pipettingInfo.srcWellID, pipettingInfo.volume));
                    scripts.Add(GetDispense(pipettingInfo.dstLabware, pipettingInfo.dstWellID, pipettingInfo.volume));
                    tipReusedTimes++;
                    bNeedWash = tipReusedTimes % maxTipReuseTimes == 0;
                    if (bNeedWash)
                        scripts.Add("W;");
                }
                scripts.Add("W;");
            }
            scripts.Add("B;");
        }

        private List<List<PipettingInfo>> SplitByTipIndex(List<PipettingInfo> pipettingInfos)
        {
            List<List<PipettingInfo>> groups = new List<List<PipettingInfo>>();
            for (int i = 0; i < 8; i++ )
            {
                groups.Add(new List<PipettingInfo>());
            }
            for (int i = 0; i < pipettingInfos.Count; i++)
            {
                int tipIndex = GetTipIndex(i);
                PipettingInfo tmp = pipettingInfos[i];
                tmp.srcWellID = tipIndex + 1;
                groups[tipIndex].Add(tmp);
            }
            return groups;
        }

        private int GetTipIndex(int i)
        {
            int batchIndex = i / tipCountLiha;
            return i = i - batchIndex * tipCountLiha;
        }


        private string GetAspOrDisp(string sLabware, int wellID, double vol, bool isAsp)
        {

            string str = string.Format("{0};{1};;;{2};;{3};;;",
                         isAsp ? 'A' : 'D',
                        sLabware,
                        wellID,
                        vol);
            return str;
        }

        private string GetAspirate(string sLabware, int srcWellID, double vol)
        {
            return GetAspOrDisp(sLabware, srcWellID, vol, true);
        }

        private string GetDispense(string sLabware, int dstWellID, double vol)
        {
            return GetAspOrDisp(sLabware, dstWellID, vol, false);
        }
    }
}
