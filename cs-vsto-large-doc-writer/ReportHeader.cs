using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LargeDocWriter
{
    public class ReportHeader
    {
        protected string mTitle = "";
        public string Title
        {
            get { return mTitle; }
            set { mTitle = value; }
        }

        protected ReportSection mOwner=null;

        public ReportSection Owner
        {
            get { return mOwner; }
            set { mOwner = value; }
        }

        public string Tag
        {
            get { return string.Format("Section-{0}", ID);  }
        }

        public int ID
        {
            get { return mOwner.ID; }
        }

        public int Level
        {
            get { return mOwner.Level; }
        }

        public ReportHeader(string title, ReportSection section)
        {
            mTitle = title;
            mOwner = section;
        }
    }
}
