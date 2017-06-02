using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LargeDocWriter.Elements;

namespace LargeDocWriter
{
    public class ReportSection
    {
        protected List<object> mParts = new List<object>();
        protected ReportSection mCurrentSubSection=null;
        protected ReportHeader mHeader = null;
        protected ReportSection mParentSection = null;
        protected int mLevel = 1;
        protected int mID = 0;
        private static int mSectionCounter = 0;

        public int Level
        {
            get { return mLevel; }
        }

        public ReportHeader Header
        {
            get { return mHeader; }
        }

        public int ID
        {
            get { return mID; }
        }

        public ReportSection(string header, int level)
        {
            mHeader = new ReportHeader(header, this);

            mLevel = level;
            mHeader.Title = header;
            mID = mSectionCounter++;
        }

        public List<object> Parts
        {
            get { return mParts; }
        }

        public void AddParagraph(ReportParagraph p)
        {
            mParts.Add(p);
        }

        public void AddFigure(ReportFigure figure)
        {
            mParts.Add(figure);
        }

        public void AddTable(ReportTable table)
        {
            mParts.Add(table);
        }

        public ReportSection ParentSection
        {
            set { mParentSection = value; }
            get { return mParentSection; }
        }

        public ReportSection CurrentSubSection
        {
            get { return mCurrentSubSection; }
        }

        public void AddSubSection(ReportSection section)
        {
            section.ParentSection = this;
            mParts.Add(section);
        }
    }
}
