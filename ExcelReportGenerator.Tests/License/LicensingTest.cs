using ExcelReportGenerator.License;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelReportGenerator.Tests.License
{
    [TestClass]
    public class LicensingTest
    {
        private const string LicenseFileName = "ExcelReportGenerator.lic";
        private const string EncryptionKey = "ccimLPARStcnulttempplattalpmeERTErPecnatMethodCtProperstancertyvhicaldatRPEULAVMElenaPctemPanTNEcxEelbdnotfocnuFlColuRTTAEULAVmnAttretagerundexcairaVdiVELEelEveimanyDecrETIATADCaitemvaluealuepreProvtyValuePallValuePtsnItlMPLATEtllacdeProclateproFmetINGSETanydl";

        [TestMethod]
        public void TestGetEncryptionKey()
        {
            MethodInfo method = typeof(Licensing).GetMethod("GetEncryptionKey", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(EncryptionKey, method.Invoke(new Licensing(), null));
        }

        [TestMethod]
        public void TestGetExpirationDateFromLicenseFile()
        {
            var rnd = new Random((int)DateTime.Now.Ticks & 0x0000FFFF);
            byte[] bytes = new byte[369];
            rnd.NextBytes(bytes);

            var date = new DateTime(2100, 12, 20);
            byte[] ticksBytes = BitConverter.GetBytes(date.Ticks);

            int licenseDateByteNumber = 217;
            byte[] allBytes = bytes.Take(licenseDateByteNumber).Concat(ticksBytes).Concat(bytes.Skip(licenseDateByteNumber).Take(bytes.Length)).ToArray();
            byte[] encryptedBytes = Encryptor.Encrypt(allBytes, EncryptionKey);

            File.WriteAllBytes(LicenseFileName, encryptedBytes);

            MethodInfo method = typeof(Licensing).GetMethod("GetExpirationDateFromLicenseFile", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.AreEqual(date, method.Invoke(new Licensing(), null));
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

        //[TestMethod]
        //public void TestThrowLicenseException()
        //{
        //    ExceptionAssert.Throws<Exception>(Licensing.ThrowLicenseException, "License violation");
        //}

        [TestMethod]
        public void EncryptMessage()
        {
            string encryptedMessage = Encryptor.Encrypt("", EncryptionKey);
        }
    }
}