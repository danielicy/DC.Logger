using System.Collections.Generic;

namespace DC.DataLogger
{
    public interface IDataLogger<T>
    {
        void Add(T item);
        void Close();
        List<T> GetEntryIds<T>(string v);
        void SaveList(List<T> data);
    }

    public enum DataLoggerType
    {
        Csv,Excel,Json
    }
}