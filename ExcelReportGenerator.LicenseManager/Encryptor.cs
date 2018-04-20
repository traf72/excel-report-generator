using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReportGenerator.LicenseManager
{
    // TODO Дублируется из-за того, что нельзя сделать ссылку на netstandard проект
    // Нужно что-нибудь придумать
    public class Encryptor
    {
        public static string Encrypt(string data, string key)
        {
            byte[] bytes = Encrypt(Encoding.UTF8.GetBytes(data), key);
            return Convert.ToBase64String(bytes);
        }

        public static byte[] Encrypt(byte[] data, string key)
        {
            return Transform(data, key, false);
        }

        public static string Decrypt(string base64Data, string key)
        {
            byte[] data = Convert.FromBase64String(base64Data);
            return Encoding.UTF8.GetString(Decrypt(data, key));
        }

        public static byte[] Decrypt(byte[] data, string key)
        {
            return Transform(data, key, true);
        }

        private static byte[] Transform(byte[] data, string key, bool decrypt)
        {
            using (ICryptoTransform transfrom = CreateCryptoTransform(key, decrypt))
            {
                using (var ms = new MemoryStream())
                {
                    using (var cs = new CryptoStream(ms, transfrom, CryptoStreamMode.Write))
                    {
                        cs.Write(data, 0, data.Length);
                    }
                    return ms.ToArray();
                }
            }
        }

        private static ICryptoTransform CreateCryptoTransform(string key, bool isDecryptor)
        {
            byte[] symmetricKey = new Rfc2898DeriveBytes(key, new byte[16]).GetBytes(16);
            byte[] iv = new byte[16];
            return isDecryptor
                ? Rijndael.Create().CreateDecryptor(symmetricKey, iv)
                : Rijndael.Create().CreateEncryptor(symmetricKey, iv);
        }
    }
}