using System;

namespace DC.DataLogger
{
    public class DataLoggerFactory<T>
    {

       public static IDataLogger<T> Create(string name,DataLoggerType type,string path="")
        {
            switch(type)
            {
                case DataLoggerType.Csv:
                    return new CsvDataLogger<T>(name,path);
                case DataLoggerType.Excel:
                    return new ExcelDataLogger<T>(name,path);
                case DataLoggerType.Json:
                    return new JsonLogger<T>(name,path);
                    default: throw new NotImplementedException();
            }
        }
    }
    

    
}
