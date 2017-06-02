using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LargeDocWriter.Elements
{
    public class ReportFigure_Chart : ReportFigure
    {
        protected Dictionary<string, float> mContent;
        protected string mXLabel = "x";
        protected string mYLabel = "y";

        public ReportFigure_Chart(ReportSection owner)
            : base(owner)
        {
            
        }

        public string XLabel
        {
            get { return mXLabel; }
            set { mXLabel = value; }
        }

        public string YLabel
        {
            get { return mYLabel; }
            set { mYLabel = value; }
        }

        public Dictionary<string, float> Content
        {
            get { return mContent; }
            set { mContent = value; }
        }

        public void Load(float[] data, string[] labels)
        {
            mContent = new Dictionary<string, float>();

            int count = data.Length;
            for (int i = 0; i < count; ++i)
            {
                float val = data[i];
                string label = labels[i];
                mContent[label] = val;
            }
        }

        public float[] Data
        {
            get
            {
                string[] labels = Labels;
                int count = labels.Length;
                float[] data = new float[count];
                for (int i = 0; i < count; ++i)
                {
                    data[i] =  mContent[labels[i]];
                }

                return data;
            }
        }

        public string[] Labels
        {
            get
            {
                return mContent.Keys.ToArray();
            }
        }
    
    }
}
