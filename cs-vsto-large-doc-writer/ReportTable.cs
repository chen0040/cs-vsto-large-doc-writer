using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace LargeDocWriter
{
    public class ReportTable
    {
        protected DataTable mContent;
        protected string mCaption="";
        protected string[] mHeaders;
        protected ReportSection mOwner;
        protected int mID = 0;
        private static int mTableCounter = 0;
        protected string mTag = "";

        public int ID
        {
            get { return mID; }
        }

        public ReportSection Owner
        {
            get { return mOwner; }
        }

        public ReportTable(ReportSection section)
        {
            mID = mTableCounter++;
            mOwner = section;

            mTag = string.Format("Table-{0}", mID);
        }

        public DataTable Content
        {
            get { return mContent; }
            set { mContent = value; }
        }

        public string Caption
        {
            get { return mCaption; }
            set { mCaption = value; }
        }

        public string[] Headers
        {
            get { return mHeaders; }
            set { mHeaders = value; }
        }

        public string Tag
        {
            get { return mTag; }
            set { mTag = value; }
        }
    }
}
