using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
 

namespace DC.DataLogger
{
    public class CsvDataLogger<T> : DataLoggerBase,IDataLogger<T>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name">file name</param>
        /// <param name="path">optional parameter
        /// if empty file will be written to base directory
        /// </param>
        public CsvDataLogger(string name,string path="") :base(name,"csv",path)
        {
           
        }

        protected override void Initialize()
        {
            
        }


        public void Close()
        {
           
        }

        private IList<PropertyInfo> Fields
        {
            get
            {
                Type myType = typeof(T);
                IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

                return props;
            }
        }



        public List<T1> GetEntryIds<T1>(string v)
        {
            throw new NotImplementedException();
        }

        public void Add(T item)
        {
            string row="";
            int cnt = 0;
            foreach (PropertyInfo prop in Fields)
            {
                var val = prop.GetValue(item, null);
                try
                {
                    if (val != null)
                        row = row + val;

                    if(cnt< Fields.Count-1)
                        row = row +  ";";
                    cnt++;

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
               
            }
            File.AppendAllText(LoggersPath, $"{row} \n");


        }

        public void SaveList(List<T> data)
        {
            throw new NotImplementedException();
        }

        protected override void WriteHeaders()
        {
            string row = "";
            int cnt = 0;
            foreach (PropertyInfo prop in Fields)
            {
                var val = prop.Name;
                try
                {
                    if (val != null)
                        row = row + val;

                    if (cnt < Fields.Count - 1)
                        row = row + ";";
                    cnt++;

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }

            File.AppendAllText(LoggersPath, $"{row} \n");
        }
    }
}
