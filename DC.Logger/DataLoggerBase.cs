using System;
using System.IO;

namespace DC.DataLogger
{
    public abstract class DataLoggerBase
    {
        protected static string LoggersPath { get; private set; }
        public DataLoggerBase(string name,string ext)
        {
            LoggersPath  = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{name}dataLog.{ext}");


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
