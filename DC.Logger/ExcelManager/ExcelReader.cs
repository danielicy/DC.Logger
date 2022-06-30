using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataLogger.ExcelManagers
{
    public class ExcelReader
    {

        public static DataTable GetExcel(string fileName)
        {

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", fileName);
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;


                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
                //// reader.IsFirstRowAsColumnNames
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);


                return dataSet.Tables[0];


            }
        }
    }
}
