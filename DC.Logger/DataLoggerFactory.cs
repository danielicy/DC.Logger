using System;

namespace DC.DataLogger
{
    public class DataLoggerFactory<T>
    {

       public static IDataLogger<T> Create(string name,DataLoggerType type)
        {
            switch(type)
            {
                case DataLoggerType.Csv:
                    return new CsvDataLogger<T>(name);
                case DataLoggerType.Excel:
                    return new ExcelDataLogger<T>(name);
                case DataLoggerType.Json:
                    return new JsonLogger<T>(name);
                    default: throw new NotImplementedException();
            }
        }
    }
    

    
}
