using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.IO;
using System.Xml.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace OOSCommon
{
    /// <summary>
    /// 
    /// </summary>
    public class SerializerClass
    {
        public byte[] Serialize_Data(object obj)
        {
            IFormatter formatter = new BinaryFormatter();
            using (MemoryStream ms = new MemoryStream())
            {
                formatter.Serialize(ms, obj);
                ms.Position = 0;
                byte[] b = new byte[ms.Length];
                ms.Read(b, 0, b.Length);
                ms.Close();
                return b;
            }
        }
        public static byte[] Serialize_Data<T>(T obj)
            where T : class, new()
        {
            XmlSerializer xs = new XmlSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream())
            {
                xs.Serialize(ms, obj);
                byte[] result = ms.ToArray();
                ms.Close();
                return result;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static object Deserialize_Data(byte[] data)
        {
            IFormatter formatter = new BinaryFormatter();
            MemoryStream ms = new MemoryStream();
            ms.Write(data, 0, data.Length);
            ms.Position = 0;
            object obj = formatter.Deserialize(ms);
            return obj;
        }
        public static T Deserialize_Data<T>(byte[] data)
            where T : class, new()
        {

            T obj = null;
            XmlSerializer ser = new XmlSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(data, 0, data.Length);
                ms.Position = 0;

                obj = (T)ser.Deserialize(ms);
                ms.Close();
                return obj;
            }
        }
    }
}
