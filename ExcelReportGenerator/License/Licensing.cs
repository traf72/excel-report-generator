using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Rendering.Panels;
using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace ExcelReportGenerator.License
{
    internal class Licensing
    {
        public static string LicenseViolationMessage;

        public static string LicenseExpiredMessage;

        private static long _licenseExpirationDateTicks;

        private readonly string _encryptionKey;

        private readonly string _licenseFileName;

        private readonly int _licenseExpirationDateByteNumber;

        private readonly int _trialLicenseExpirationDaysCount;

        private readonly string _registryPath;

        private readonly string _registryKey;

        public Licensing()
        {
            _encryptionKey = GetEncryptionKey();
            _licenseFileName = GetLicenseFileName();
            _licenseExpirationDateByteNumber = GetLicenseExpirationDateByteNumber();
            _trialLicenseExpirationDaysCount = GetTrialLicenseExpirationDaysCount();
            _registryPath = GetRegistryPath();
            _registryKey = GetRegistryKey();
            LicenseViolationMessage = LicenseViolationMessage ?? GetLicenseViolationMessage();
            LicenseExpiredMessage = LicenseExpiredMessage ?? GetLicenseExpiredMessage();
        }

        public static DateTime LicenseExpirationDate => DateTime.FromBinary(_licenseExpirationDateTicks);

        private string GetEncryptionKey()
        {
            Type[] keyPartTypes = typeof(IPanel).Assembly.GetExportedTypes()
                .Where(t => t.IsDefined(typeof(LicenceKeyPartAttribute), false))
                .OrderBy(t => t.Name[t.Name.Length > 5 ? 5 : t.Name.Length - 1])
                .ToArray();

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
            return Encryptor.Decrypt("QQMwvi3YaowQykjGKE3C4tB4R6V+K8QsvWWLPsw++m0=", _encryptionKey);
        }

        private int GetLicenseExpirationDateByteNumber()
        {
            return int.Parse(Encryptor.Decrypt("4yIXZMhFBqIQi8GVjXwC/w==", _encryptionKey));
        }

        private int GetTrialLicenseExpirationDaysCount()
        {
            return int.Parse(Encryptor.Decrypt("xgJ6Jyafg1PcACSE052gWg==", _encryptionKey));
        }

        private string GetRegistryPath()
        {
            return Encryptor.Decrypt("hc0znFR+v6ivevi7FyTukaLhUrVL1wGgVeSMlg6Hhr/itV8mOIWifcwbX4qGhCUO", _encryptionKey);
        }

        private string GetRegistryKey()
        {
            return Encryptor.Decrypt("r4tm9TgtIw3el1bmx4zpDA==", _encryptionKey);
        }

        private string GetLicenseViolationMessage()
        {
            return Encryptor.Decrypt("ThMkguVsYPwzsRkWlvnT4Pxb141vrk9WQMhfkMKBVAk=", _encryptionKey);
        }

        private string GetLicenseExpiredMessage()
        {
            return Encryptor.Decrypt("IO07IiZ1altX8ijLJI2nAw==", _encryptionKey);
        }

        public void LoadLicenseInfo()
        {
            Interlocked.Exchange(ref _licenseExpirationDateTicks, GetExpirationDateTicks());
        }

        private long GetExpirationDateTicks()
        {
            var assemblyFile = new FileInfo(typeof(IPanel).Assembly.Location);
            string licenseFilePath = $"{assemblyFile.Directory}\\{_licenseFileName}";
            return !File.Exists(licenseFilePath) ? GetTrialExpirationDate(assemblyFile) : GetExpirationDateFromLicenseFile(licenseFilePath);
        }

        private long GetExpirationDateFromLicenseFile(string licenseFilePath)
        {
            byte[] licenseFileContent = Encryptor.Decrypt(File.ReadAllBytes(licenseFilePath), _encryptionKey);
            CheckHashSum(licenseFileContent);
            return BitConverter.ToInt64(licenseFileContent.Skip(_licenseExpirationDateByteNumber).Take(sizeof(long)).ToArray(), 0);
        }

        private long GetTrialExpirationDate(FileInfo assemblyFile)
        {
            if ((DateTime.Now - assemblyFile.CreationTime).TotalDays > _trialLicenseExpirationDaysCount)
            {
                return assemblyFile.CreationTime.Ticks;
            }

            byte[] registryValue = (byte[])Registry.GetValue(_registryPath, _registryKey, null);
            if (registryValue != null)
            {
                byte[] decryptedBytes = Encryptor.Decrypt(registryValue, _encryptionKey);
                return BitConverter.ToInt64(decryptedBytes, 0);
            }

            long trialExpirationDate = DateTime.Now.AddDays(_trialLicenseExpirationDaysCount).Ticks;
            SetTrialExpirationDateInRegistry(trialExpirationDate);
            return trialExpirationDate;
        }

        private void SetTrialExpirationDateInRegistry(long trialExpirationDate)
        {
            byte[] bytes = BitConverter.GetBytes(trialExpirationDate);
            byte[] encryptedBytes = Encryptor.Encrypt(bytes, _encryptionKey);
            Registry.SetValue(_registryPath, _registryKey, encryptedBytes, RegistryValueKind.Binary);
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