using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace OOSCommon
{
    /// <summary>
    /// 加密类
    /// </summary>
    public class HashEncrypt
    {
        /// <summary>
        /// </summary>
        /// <param name="strSource">待加密字串</param>
        /// <returns>加密后的字串</returns>
        public static string MD5Encrypt(string strSource)
        {
            return MD5Encrypt(strSource, 32);
        }

        /// <summary>
        /// Verify a hash against a string.
        /// </summary>
        /// <param name="input"></param>
        /// <param name="hash"></param>
        /// <returns></returns>
        public static bool VerifyMd5Hash(string input, string hash)
        {
            // Hash the input.
            string hashOfInput = MSGetMd5Hash(input);
            // Create a StringComparer an comare the hashes.
            StringComparer comparer = StringComparer.OrdinalIgnoreCase;
            if (0 == comparer.Compare(hashOfInput, hash))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="strSource">待加密字串</param>
        /// <param name="length">16或32值之一,其它则采用.net默认MD5加密算法</param>
        /// <returns>加密后的字串</returns>
        static string MD5Encrypt(string strSource, int length)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(strSource);
            byte[] hashValue = ((System.Security.Cryptography.HashAlgorithm)System.Security.Cryptography.CryptoConfig.CreateFromName("MD5")).ComputeHash(bytes);
            StringBuilder sb = new StringBuilder();
            switch (length)
            {
                case 16:
                    for (int i = 4; i < 12; i++)
                        sb.Append(hashValue[i].ToString("x2"));
                    break;
                case 32:
                    for (int i = 0; i < 16; i++)
                    {
                        sb.Append(hashValue[i].ToString("x2"));
                    }
                    break;
                default:
                    for (int i = 0; i < hashValue.Length; i++)
                    {
                        sb.Append(hashValue[i].ToString("x2"));
                    }
                    break;
            }
            return sb.ToString();
        }
        /// <summary>
        /// md5 encrypt
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        static string MSMD5(String str)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] data = System.Text.Encoding.Default.GetBytes(str);
            byte[] result = md5.ComputeHash(data);
            String ret = "";
            for (int i = 0; i < result.Length; i++)
                ret += result[i].ToString("x").PadLeft(2, '0');
            return ret;
        }

        /// <summary>
        /// get md5 hash value
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        static string MSGetMd5Hash(string input)
        {
            // Create a new instance of the MD5CryptoServiceProvider object.
            MD5 md5Hasher = System.Security.Cryptography.MD5.Create();
            // Convert the input string to a byte array and compute the hash.
            byte[] data = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input));
            // Create a new Stringbuilder to collect the bytes
            // and create a string.
            StringBuilder sBuilder = new StringBuilder();
            // Loop through each byte of the hashed data
            // and format each one as a hexadecimal string.
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            // Return the hexadecimal string.
            return sBuilder.ToString();
        }

        /// <summary>
        /// 对文件进行MD5加密
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static string GetFileMD5(string filepath)
        {
            StringBuilder sb = new StringBuilder();
            using (MD5 md5 = MD5.Create())
            {
                using (FileStream fs = File.OpenRead(filepath))
                {
                    byte[] newB = md5.ComputeHash(fs);
                    foreach (byte item in newB)
                    {
                        sb.Append(item.ToString("x2"));
                    }
                }
            }

            return sb.ToString();
        }

        private static string encryptKey = "!QAZ";
        /// <summary> /// 加密字符串   
        /// </summary>  
        /// <param name="str">要加密的字符串</param>  
        /// <returns>加密后的字符串</returns>  
        public static string EncryptForStr(string str)
        {
            try
            {
                DESCryptoServiceProvider descsp = new DESCryptoServiceProvider();   //实例化加/解密类对象
                byte[] key = Encoding.Unicode.GetBytes(encryptKey); //定义字节数组，用来存储密钥  
                byte[] data = Encoding.Unicode.GetBytes(str);//定义字节数组，用来存储要加密的字符串  
                MemoryStream MStream = new MemoryStream(); //实例化内存流对象    
                //使用内存流实例化加密流对象   
                CryptoStream CStream = new CryptoStream(MStream, descsp.CreateEncryptor(key, key), CryptoStreamMode.Write);
                CStream.Write(data, 0, data.Length);  //向加密流中写入数据      
                CStream.FlushFinalBlock();              //释放加密流      
                return Convert.ToBase64String(MStream.ToArray());//返回加密后的字符串  
            }
            catch
            {
                return "";
            }
        }
        /// <summary>  
        /// 解密字符串   
        /// </summary>  
        /// <param name="str">要解密的字符串</param>  
        /// <returns>解密后的字符串</returns>  
        public static string DecryptForStr(string str)
        {
            try
            {
                DESCryptoServiceProvider descsp = new DESCryptoServiceProvider();   //实例化加/解密类对象    
                byte[] key = Encoding.Unicode.GetBytes(encryptKey); //定义字节数组，用来存储密钥    
                byte[] data = Convert.FromBase64String(str);//定义字节数组，用来存储要解密的字符串  
                MemoryStream MStream = new MemoryStream(); //实例化内存流对象      
                //使用内存流实例化解密流对象       
                CryptoStream CStream = new CryptoStream(MStream, descsp.CreateDecryptor(key, key), CryptoStreamMode.Write);
                CStream.Write(data, 0, data.Length);      //向解密流中写入数据     
                CStream.FlushFinalBlock();               //释放解密流      
                return Encoding.Unicode.GetString(MStream.ToArray());       //返回解密后的字符串  
            }
            catch
            {
                return "";
            }
        }
    }
}
