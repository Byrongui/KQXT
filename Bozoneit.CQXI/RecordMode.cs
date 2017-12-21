using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bozoneit.CQXI
{
    public class RecordMode
    {
        public string  name { get; set; }
        public string date { get; set; }
        public string  startTime { get; set; }
        public string qdTime { get; set; }

        private string _qtTime;
        //签退时间
        public string qtTime
        {
            get { return _qtTime; }
            set
            {
                _qtTime = value;
                if (string.IsNullOrEmpty(qdTime.Trim())){//签到时间为空时
                    isYC = true;
                    return;
                }

                if(string.IsNullOrEmpty(value)) //签退时间为空
                {
                    isYC = true;
                    return;
                }

                DateTime dtqdTime = Convert.ToDateTime(string.Format("{0} {1}", date, qdTime));
                DateTime dtqtTime = Convert.ToDateTime(string.Format("{0} {1}", date, value));

                TimeSpan ts = dtqtTime - dtqdTime;

                if (ts.TotalMinutes < 525)
                {
                    if (dtqdTime.Hour >= 12)//下午来上班
                    {
                        if (ts.TotalMinutes < 480)
                        {
                            isYC = true;
                        }
                    }
                    else
                    {
                        isYC = true;
                    }
                }
            }
        }

        public string endTime { get; set; }
        public string bz { get; set; }
        public string bm { get; set; }

        /// <summary>
        /// 是否异常
        /// </summary>
        public bool isYC { get; set; }
        public string ycTime { get; set; }

        public string this[int index]
        {
            get
            {
                switch (index)
                {
                    case 0:return name; 
                    case 1: return date;
                    case 2: return startTime;
                    case 3:return endTime;
                    case 4: return qdTime;
                    case 5: return qtTime;
                    case 6: return bm;
                    case 7: return bz;
                    default: return "";
                        
                }
            }
        }
        
    }
}
