using ExcelReportGenerator.License;
using ExcelReportGenerator.Rendering.Panels;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;

namespace ExcelReportGenerator.Tests.License
{
    [TestClass]
    public class LicensingTest
    {
        private const string EncryptionKey = "lColuccimTNERPEULLPARSIstanRTTAtalpmednotfoertyvcnuFecxEelblttempplatrPecnatMethodCtPropertdataittcnuElenaPcERTEMtemPanelEvePLATEimanyDecrFmetemvalueptyValuePallValuePtsnItleProclateproairaVditageralueprundexctllacdEULAVceProvNGSETAVMETIVELEanydlmnAttr";
        private const string LicenseFileName = "ExcelReportGenerator.lic";
        private const string RegistryPath = @"HKEY_CURRENT_USER\SOFTWARE\ProtectedStorage";
        private const string RegistryKey = @"vcim";
        private const int LicenseExpirationDateByteNumber = 217;
        private const int TrialLicenseExpirationDaysCount = 30;
        private const string LicenseViolationMessage = "License violation";
        private const string LicenseExpiredMessage = "License expired";

        [TestMethod]
        public void TestGetEncryptionKey()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetEncryptionKey", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(EncryptionKey, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestLoadLicenseInfo()
        {
            // Trial license from scratch
            File.Delete(LicenseFileName);
            DeleteLicenseInfoFromRegistry();

            DateTime expirationDate = DateTime.Now.AddDays(TrialLicenseExpirationDaysCount);

            var licensing = new Licensing();
            licensing.LoadLicenseInfo();
            Assert.AreEqual(expirationDate.Date, Licensing.LicenseExpirationDate.Date);
            CheckLicenseExpirationDateInRegistry(expirationDate);

            // If trial expiration date already exists in registry
            expirationDate = DateTime.Now.AddDays(1);
            SetLicenseInfoToRegistry(expirationDate);
            licensing.LoadLicenseInfo();
            Assert.AreEqual(expirationDate.Date, Licensing.LicenseExpirationDate.Date);
            CheckLicenseExpirationDateInRegistry(expirationDate);

            DeleteLicenseInfoFromRegistry();

            // License from file
            var rnd = new Random((int)DateTime.Now.Ticks & 0x0000FFFF);
            byte[] fictitiousBytes = new byte[369];
            rnd.NextBytes(fictitiousBytes);

            expirationDate = new DateTime(2100, 12, 20);
            byte[] ticksBytes = BitConverter.GetBytes(expirationDate.Ticks);

            byte[] payload = fictitiousBytes.Take(LicenseExpirationDateByteNumber).Concat(ticksBytes).Concat(fictitiousBytes.Skip(LicenseExpirationDateByteNumber).Take(fictitiousBytes.Length)).ToArray();
            byte[] hash;
            using (MD5 hashAlg = MD5.Create())
            {
                hash = hashAlg.ComputeHash(payload);
            }

            byte[] allBytes = payload.Concat(hash).ToArray();
            byte[] encryptedBytes = Encryptor.Encrypt(allBytes, EncryptionKey);
            File.WriteAllBytes(LicenseFileName, encryptedBytes);

            licensing.LoadLicenseInfo();
            Assert.AreEqual(expirationDate, Licensing.LicenseExpirationDate);
            Assert.IsNull(Registry.GetValue(RegistryPath, RegistryKey, null));

            // License file changed
            using (var fs = File.Open(LicenseFileName, FileMode.Open))
            {
                fs.WriteByte(100);
            }

            ExceptionAssert.Throws<Exception>(() => licensing.LoadLicenseInfo(), LicenseViolationMessage);
            Assert.IsNull(Registry.GetValue(RegistryPath, RegistryKey, null));

            File.Delete(LicenseFileName);
        }

        [TestMethod]
        public void TestGetLicenseFileName()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetLicenseFileName", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual("ExcelReportGenerator.lic", method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestGetLicenseExpirationDateByteNumber()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetLicenseExpirationDateByteNumber", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(LicenseExpirationDateByteNumber, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestGetTrialLicenseExpirationDaysCount()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetTrialLicenseExpirationDaysCount", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(TrialLicenseExpirationDaysCount, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestGetRegistryPath()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetRegistryPath", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(RegistryPath, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestGetRegistryKey()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetRegistryKey", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(RegistryKey, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestGetLicenseViolationMessage()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetLicenseViolationMessage", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(LicenseViolationMessage, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestGetLicenseExpiredMessage()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetLicenseExpiredMessage", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(LicenseExpiredMessage, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestGetTrialExpirationDateFromRegistry()
        {
            DateTime expirationDate = DateTime.Now.AddDays(1);
            SetLicenseInfoToRegistry(expirationDate);

            long result = GetLicenseInfoFromRegistry();
            Assert.AreEqual(expirationDate.Ticks, result);
        }

        [TestMethod]
        public void TestSetTrialExpirationDateToRegistry()
        {
            DeleteLicenseInfoFromRegistry();
            DateTime expectedDate = DateTime.Now.AddDays(TrialLicenseExpirationDaysCount).Date;

            long result = GetLicenseInfoFromRegistry();
            Assert.AreEqual(expectedDate, DateTime.FromBinary(result).Date);

            CheckLicenseExpirationDateInRegistry(expectedDate);
        }

        [TestMethod]
        public void TestCreateLicensing()
        {
            var licensing = new Licensing();
            Assert.AreEqual(EncryptionKey, licensing.GetType().GetField("_encryptionKey", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(LicenseFileName, licensing.GetType().GetField("_licenseFileName", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(LicenseExpirationDateByteNumber, licensing.GetType().GetField("_licenseExpirationDateByteNumber", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(TrialLicenseExpirationDaysCount, licensing.GetType().GetField("_trialLicenseExpirationDaysCount", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(RegistryPath, licensing.GetType().GetField("_registryPath", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(RegistryKey, licensing.GetType().GetField("_registryKey", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(LicenseViolationMessage, Licensing.LicenseViolationMessage);
            Assert.AreEqual(LicenseExpiredMessage, Licensing.LicenseExpiredMessage);
        }

        [TestMethod]
        public void EncryptMessage()
        {
            string encryptedLicenseFileName = Encryptor.Encrypt(LicenseFileName, EncryptionKey);
            string encryptedLicenseExpirationDateByteNumber = Encryptor.Encrypt(LicenseExpirationDateByteNumber.ToString(), EncryptionKey);
            string encryptedTrialLicenseExpirationDaysCount = Encryptor.Encrypt(TrialLicenseExpirationDaysCount.ToString(), EncryptionKey);
            string encryptedRegistryPath = Encryptor.Encrypt(RegistryPath, EncryptionKey);
            string encryptedRegistryKey = Encryptor.Encrypt(RegistryKey, EncryptionKey);
            string encryptedLicenseViolationMessage = Encryptor.Encrypt(LicenseViolationMessage, EncryptionKey);
            string encryptedLicenseExpiredMessage = Encryptor.Encrypt(LicenseExpiredMessage, EncryptionKey);
        }

        private void CheckLicenseExpirationDateInRegistry(DateTime expectedDate)
        {
            byte[] dateTicksFromRegistry = (byte[])Registry.GetValue(RegistryPath, RegistryKey, null);
            byte[] decryptedBytes = Encryptor.Decrypt(dateTicksFromRegistry, EncryptionKey);
            Assert.AreEqual(expectedDate.Date, DateTime.FromBinary(BitConverter.ToInt64(decryptedBytes, 0)).Date);
        }

        private void DeleteLicenseInfoFromRegistry()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\ProtectedStorage", true))
            {
                key?.DeleteValue(RegistryKey, false);
            }
        }

        private void SetLicenseInfoToRegistry(DateTime date)
        {
            var licensing = new Licensing();
            MethodInfo setMethod = licensing.GetType().GetMethod("SetTrialExpirationDateInRegistry", BindingFlags.Instance | BindingFlags.NonPublic);
            setMethod.Invoke(licensing, new object[] { date.Ticks });
        }

        private long GetLicenseInfoFromRegistry()
        {
            var licensing = new Licensing();
            MethodInfo getMethod = licensing.GetType().GetMethod("GetTrialExpirationDate", BindingFlags.Instance | BindingFlags.NonPublic);
            return (long)getMethod.Invoke(licensing, new object[] { new FileInfo(typeof(IPanel).Assembly.Location) });
        }
    }
}