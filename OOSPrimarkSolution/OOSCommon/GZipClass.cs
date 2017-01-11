using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Compression;
using SevenZip;

namespace OOSCommon
{
    public class GZipClass
    {
        /// <summary>
        /// 压缩数据
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static byte[] Compress(byte[] data)
        {
            try
            {
                using (MemoryStream compMS = new MemoryStream())
                {
                    using (GZipStream compStream = new GZipStream(compMS, System.IO.Compression.CompressionMode.Compress, true))
                    {
                        compStream.Write(data, 0, data.Length);
                        compStream.Close();
                        byte[] zipData = compMS.ToArray();
                        compStream.Close();
                        return zipData;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public static byte[] CompressObjBySevenZip<T>(T obj) where T : class, new()
        {
            try
            {
                byte[] serialdata = SerializerClass.Serialize_Data<T>(obj);
                byte[] zipserialdata = SevenZipCompressor.CompressBytes(serialdata);
                return zipserialdata;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static T ExtractorObjBySevenZip<T>(byte[] data) where T : class, new()
        {
            try
            {
                byte[] unzipserialdata = SevenZipExtractor.ExtractBytes(data);
                T obj = SerializerClass.Deserialize_Data<T>(unzipserialdata);
                return obj;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static byte[] CompressStringBySevenZip(string str)
        {
            try
            {
                byte[] data = Encoding.Unicode.GetBytes(str);
                byte[] zipdata = SevenZipCompressor.CompressBytes(data);
                return zipdata;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string ExtractorStringBySevenZip(byte[] data)
        {
            try
            {
                byte[] unzipdata = SevenZipExtractor.ExtractBytes(data);
                string str = Encoding.Unicode.GetString(unzipdata);
                return str;
            }
            catch (Exception)
            {
                return null;
            }
        }


        /// <summary>
        /// 解压缩数据
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static byte[] DeCompress(byte[] data)
        {
            try
            {
                using (MemoryStream input = new MemoryStream())
                {
                    input.Write(data, 0, data.Length);
                    input.Position = 0;
                    using (GZipStream gzip = new GZipStream(input, System.IO.Compression.CompressionMode.Decompress, true))
                    {
                        using (MemoryStream output = new MemoryStream())
                        {
                            byte[] buff = new byte[4096];
                            int read = gzip.Read(buff, 0, buff.Length);
                            while (read > 0)
                            {
                                output.Write(buff, 0, read);
                                read = gzip.Read(buff, 0, buff.Length);
                            }
                            gzip.Close();
                            byte[] rebytes = output.ToArray();
                            output.Close();
                            input.Close();
                            return rebytes;
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
