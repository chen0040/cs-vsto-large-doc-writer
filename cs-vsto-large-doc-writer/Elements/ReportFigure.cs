using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LargeDocWriter.Elements
{
    public class ReportFigure
    {
        protected string mFileName = "";
        protected string mFilePath = "";
        protected string mTag = "";
        protected string mCaption = "";
        protected int mFigureWidth;
        protected int mFigureHeight;
        protected ReportSection mOwner;
        private static int mFigureCounter = 0;
        private int mID;
        private string mTitle = "";

        public int ID
        {
            get { return mID; }
        }

        public ReportFigure(ReportSection owner)
        {
            mID = mFigureCounter++;
            mOwner = owner;
            mTag = string.Format("Figure-{0}", mID);
        }

        public ReportSection Owner
        {
            get { return mOwner; }
        }

        public string FileName
        {
            set { mFileName = value; }
            get { return mFileName; }
        }

        public string Tag
        {
            get { return mTag; }
            set { mTag = value; }
        }

        public string Caption
        {
            get { return mCaption; }
            set { mCaption = value; }
        }

        public string Title
        {
            get { return mTitle; }
            set { mTitle = value; }
        }

        public int Width
        {
            get { return mFigureWidth; }
            set { mFigureWidth = value; }
        }

        public int Height
        {
            get { return mFigureHeight; }
            set { mFigureHeight = value; }
        }

    }
}
