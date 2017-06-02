using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LargeDocWriter
{
    public class ReportParagraph
    {
        protected string mContent=null;
        protected object[] mArgs = null;
        private static int mParagraphCounter = 0;
        protected int mID = 0;

        public int ID
        {
            get { return mID; }
        }

        public string Tag
        {
            get { return string.Format("Paragraph-{0}", mID); }
        }

        public string Content
        {
            get { return mContent; }
        }

        public object[] Args
        {
            get { return mArgs; }
        }

        public ReportParagraph(ReportSection section, string content, params object[] args)
        {
            mID = mParagraphCounter++;
            mContent = content;
            mArgs = args;
        }

        public override string ToString()
        {
            return string.Format(mContent, mArgs);
        }
    }
}
