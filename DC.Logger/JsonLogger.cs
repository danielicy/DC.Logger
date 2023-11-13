using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
 

namespace DC.DataLogger
{
    internal class JsonLogger<T> : DataLoggerBase,IDataLogger<T>
    {
        private int _count;
        public JsonLogger(string name, string path="") : base(name,"json",path)
        {
           
        }

        protected override void Initialize()
        {
            File.AppendAllText(LoggersPath, "{\"objects\": [");
        }

        public void Close()
        {
            File.AppendAllText(LoggersPath, $"]}}");

        }

        public List<T1> GetEntryIds<T1>(string v)
        {
            throw new NotImplementedException();
        }

        public void Add(T item)
        {
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(item);
            File.AppendAllText(LoggersPath, $"{ json},");
            _count++;
        }

        public void SaveList(List<T> data)
        {
            throw new NotImplementedException();
        }

        protected override void WriteHeaders()
        {
           
        }

       
    }
}
