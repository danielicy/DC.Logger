using System;
using System.IO;

namespace DC.DataLogger
{
    public abstract class DataLoggerBase
    {
        protected static string LoggersPath { get; private set; }
        public DataLoggerBase(string name,string ext,string path)
        {
            var lpath = string.IsNullOrEmpty(path) ? AppDomain.CurrentDomain.BaseDirectory : path;
            LoggersPath  = Path.Combine(lpath, $"{name}dataLog.{ext}");


            Initialize();
            if (!File.Exists(LoggersPath))
            {
                FileStream fs = File.Create(LoggersPath);
                fs.Close();
                WriteHeaders();
            } 
            
        }

        protected abstract void WriteHeaders();
        protected abstract void Initialize();
    }
}
