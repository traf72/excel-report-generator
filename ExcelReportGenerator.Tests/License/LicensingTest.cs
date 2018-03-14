using ExcelReportGenerator.License;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using ExcelReportGenerator.Tests.CustomAsserts;

namespace ExcelReportGenerator.Tests.License
{
    [TestClass]
    public class LicensingTest
    {
        private const string EncryptionKey = "ccimLPARStcnulttempplattalpmeERTErPecnatMethodCtProperstancertyvhicaldatRPEULAVMElenaPctemPanTNEcxEelbdnotfocnuFlColuRTTAEULAVmnAttretagerundexcairaVdiVELEelEveimanyDecrETIATADCaitemvaluealuepreProvtyValuePallValuePtsnItlMPLATEtllacdeProclateproFmetINGSETanydl";
        private const string LicenseFileName = "ExcelReportGenerator.lic";
        private const string LicenseViolationMessage = "License violation";
        private const string LicenseExpiredMessage = "License expired";
        private const int LicenseExpirationDateByteNumber = 217;

        [TestMethod]
        public void TestGetEncryptionKey()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetEncryptionKey", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(EncryptionKey, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestLoadLicenseInfo()
        {
            var rnd = new Random((int)DateTime.Now.Ticks & 0x0000FFFF);
            byte[] fictitiousBytes = new byte[369];
            rnd.NextBytes(fictitiousBytes);

            var date = new DateTime(2100, 12, 20);
            byte[] ticksBytes = BitConverter.GetBytes(date.Ticks);

            byte[] payload = fictitiousBytes.Take(LicenseExpirationDateByteNumber).Concat(ticksBytes).Concat(fictitiousBytes.Skip(LicenseExpirationDateByteNumber).Take(fictitiousBytes.Length)).ToArray();
            byte[] hash;
            using (MD5 hashAlg = MD5.Create())
            {
                hash = hashAlg.ComputeHash(payload);
            }

            byte[] allBytes = payload.Concat(hash).ToArray();
            byte[] encryptedBytes = Encryptor.Encrypt(allBytes, EncryptionKey);
            File.WriteAllBytes(LicenseFileName, encryptedBytes);

            var licensing = new Licensing();
            licensing.LoadLicenseInfo();
            Assert.AreEqual(date, Licensing.LicenseExpirationDate);

            using (var fs = File.Open(LicenseFileName, FileMode.Open))
            {
                fs.WriteByte(100);
            }

            ExceptionAssert.Throws<Exception>(() => licensing.LoadLicenseInfo(), LicenseViolationMessage);

            File.Delete(LicenseFileName);
            ExceptionAssert.Throws<Exception>(() => licensing.LoadLicenseInfo(), "License file was not found");
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
            Assert.AreEqual(217, method.Invoke(new Licensing(), null));
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
        public void TestCreateLicensing()
        {
            var licensing = new Licensing();
            Assert.AreEqual(EncryptionKey, licensing.GetType().GetField("_encryptionKey", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(LicenseFileName, licensing.GetType().GetField("_licenseFileName", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(LicenseExpirationDateByteNumber, licensing.GetType().GetField("_licenseExpirationDateByteNumber", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(licensing));
            Assert.AreEqual(LicenseViolationMessage, Licensing.LicenseViolationMessage);
            Assert.AreEqual(LicenseExpiredMessage, Licensing.LicenseExpiredMessage);
        }

        [TestMethod]
        public void EncryptMessage()
        {
            string encryptedMessage = Encryptor.Encrypt("", EncryptionKey);
        }
    }
}