using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Rendering.Panels;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace ExcelReportGenerator.License
{
    internal class Licensing
    {
        public static string LicenseViolationMessage;

        private static long _licenseExpirationDateTicks;

        private readonly string _encryptionKey;

        private readonly string _licenseFileName;

        private readonly int _licenseExpirationDateByteNumber;

        public Licensing()
        {
            _encryptionKey = GetEncryptionKey();
            _licenseFileName = GetLicenseFileName();
            _licenseExpirationDateByteNumber = GetLicenseExpirationDateByteNumber();
            LicenseViolationMessage = LicenseViolationMessage ?? GetLicenseViolationMessage();
        }

        public static DateTime LicenseExpirationDate => DateTime.FromBinary(_licenseExpirationDateTicks);

        private string GetEncryptionKey()
        {
            Type[] keyPartTypes = typeof(IPanel).Assembly.GetExportedTypes().Where(t => t.IsDefined(typeof(LicenceKeyPartAttribute), false)).ToArray();
            var key = new StringBuilder();
            foreach (Type keyPartType in keyPartTypes)
            {
                var licencePartAttr = keyPartType.GetCustomAttribute<LicenceKeyPartAttribute>();
                int startIndex = (int)(keyPartType.Name.Length * .23);
                int endIndex = (int)(keyPartType.Name.Length * .77);
                string part = keyPartType.Name.Substring(startIndex, endIndex - startIndex);
                if (licencePartAttr.L)
                {
                    part = part.Reverse();
                }
                if (licencePartAttr.U)
                {
                    part = part.ToLower();
                }
                if (licencePartAttr.R)
                {
                    part = part.ToUpper();
                }

                key.Insert((int)(key.Length * .47), part);
            }

            return key.ToString();
        }

        private string GetLicenseFileName()
        {
            return Encryptor.Decrypt("Fy3q6kJ5kZHGRWrQwgwavTCZrkaGgPNkXH7k0hny6rE=", _encryptionKey);
        }

        private int GetLicenseExpirationDateByteNumber()
        {
            return int.Parse(Encryptor.Decrypt("jtwfeAvEGgNNST/cuNJyWA==", _encryptionKey));
        }

        private string GetLicenseViolationMessage()
        {
            return Encryptor.Decrypt("4BlFflydyzyduXfzPQVm9+adf2dNEC9ydZZRieFmkfg=", _encryptionKey);
        }

        public void LoadLicenseInfo()
        {
            Interlocked.Exchange(ref _licenseExpirationDateTicks, GetExpirationDateTicksFromLicenseFile());
        }

        private long GetExpirationDateTicksFromLicenseFile()
        {
            byte[] licenseFileContent = GetLicenseFileContent();
            CheckHashSum(licenseFileContent);
            return BitConverter.ToInt64(licenseFileContent.Skip(_licenseExpirationDateByteNumber).Take(sizeof(long)).ToArray(), 0);
        }

        private byte[] GetLicenseFileContent()
        {
            var licenseFileLocation = new FileInfo(typeof(IPanel).Assembly.Location);
            string filePath = $"{licenseFileLocation.Directory}\\{_licenseFileName}";
            if (!File.Exists(filePath))
            {
                throw new Exception(Encryptor.Decrypt("+j2CNbC4fKTeeHt/ESaW/kP5nCCJn/MaDTmeAytwyu8=", _encryptionKey));
            }

            return Encryptor.Decrypt(File.ReadAllBytes(filePath), _encryptionKey);
        }

        private void CheckHashSum(byte[] data)
        {
            string computedHash;
            int hashSizeInBytes;
            using (MD5 hashAlg = MD5.Create())
            {
                hashSizeInBytes = hashAlg.HashSize / 8;
                byte[] payload = data.Take(data.Length - hashSizeInBytes).ToArray();
                computedHash = Convert.ToBase64String(hashAlg.ComputeHash(payload));
            }

            string originalHash = Convert.ToBase64String(data.Skip(data.Length - hashSizeInBytes).ToArray());
            if (computedHash != originalHash)
            {
                throw new Exception(LicenseViolationMessage);
            }
        }
    }
}