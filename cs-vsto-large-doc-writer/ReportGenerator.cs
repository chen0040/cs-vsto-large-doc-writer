using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LargeDocWriter
{
    public abstract class ReportGenerator 
    {
        protected ReportModel mModel;
        

        public ReportGenerator(ReportModel model)
        {
            mModel = model;
        }

        public abstract void GenerateReport(string filepath, string content_folder_path);

        public delegate void OnMessageHandler(string message);
        public event OnMessageHandler OnMessage;

        protected void NotifyMessage(string message)
        {
            if (OnMessage != null)
            {
                OnMessage(message);
            }
        }

        public delegate void TaskProgressChangedHandler(string message, int progress_percentage);
        public event TaskProgressChangedHandler TaskProgressChanged;

        protected void NotifyTaskProgressChanged(string message, int progress_percentage)
        {
            if (TaskProgressChanged != null)
            {
                TaskProgressChanged(message, progress_percentage);
            }
        }

        public delegate void OnErrorHandler(string error_message);
        public event OnErrorHandler OnError;

        protected void ReportError(string error_message)
        {
            if (OnError != null)
            {
                OnError(error_message);
            }
        }
    }
}
