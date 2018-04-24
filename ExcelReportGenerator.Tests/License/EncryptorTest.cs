using System;
using ExcelReportGenerator.License;
using ExcelReportGenerator.Tests.CustomAsserts;
using NUnit.Framework;

namespace ExcelReportGenerator.Tests.License
{
    
    public class EncryptorTest
    {
        [Test]
        public void TestEncryptDecrypt()
        {
            string key = "Key_1";
            string message = "Hello World!";
            string encryptedMessage = Encryptor.Encrypt(message, key);
            Assert.AreEqual(message, Encryptor.Decrypt(encryptedMessage, key));

            key = Guid.NewGuid().ToString();
            message = "nvkadfy84357$*&#($(){}[]<>(&~\"";
            encryptedMessage = Encryptor.Encrypt(message, key);
            Assert.AreEqual(message, Encryptor.Decrypt(encryptedMessage, key));

            message = string.Empty;
            encryptedMessage = Encryptor.Encrypt(message, key);
            Assert.AreEqual(message, Encryptor.Decrypt(encryptedMessage, key));

            key = " ";
            encryptedMessage = Encryptor.Encrypt(message, key);
            Assert.AreEqual(message, Encryptor.Decrypt(encryptedMessage, key));

            ExceptionAssert.Throws<ArgumentNullException>(() => Encryptor.Encrypt((string)null, key));
            ExceptionAssert.Throws<ArgumentNullException>(() => Encryptor.Decrypt((string)null, key));
            ExceptionAssert.Throws<ArgumentException>(() => Encryptor.Encrypt(message, string.Empty));
            ExceptionAssert.Throws<ArgumentException>(() => Encryptor.Decrypt(message, string.Empty));
            ExceptionAssert.Throws<ArgumentException>(() => Encryptor.Encrypt(message, null));
            ExceptionAssert.Throws<ArgumentException>(() => Encryptor.Decrypt(message, null));
        }
    }
}