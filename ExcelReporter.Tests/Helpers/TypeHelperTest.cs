using System.Collections.Generic;
using ExcelReporter.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Helpers
{
    [TestClass]
    public class TypeHelperTest
    {
        [TestMethod]
        public void TestIsKeyValuePair()
        {
            Assert.IsTrue(TypeHelper.IsKeyValuePair(typeof(KeyValuePair<object, object>)));
            Assert.IsTrue(TypeHelper.IsKeyValuePair(typeof(KeyValuePair<string, object>)));
            Assert.IsTrue(TypeHelper.IsKeyValuePair(typeof(KeyValuePair<string, string>)));
            Assert.IsTrue(TypeHelper.IsKeyValuePair(typeof(KeyValuePair<string, int>)));
            Assert.IsTrue(TypeHelper.IsKeyValuePair(typeof(KeyValuePair<int, string>)));
            Assert.IsFalse(TypeHelper.IsKeyValuePair(typeof(IEnumerable<KeyValuePair<object, object>>)));
            Assert.IsFalse(TypeHelper.IsKeyValuePair(typeof(IEnumerable<object>)));
            Assert.IsFalse(TypeHelper.IsKeyValuePair(typeof(string)));
        }

        [TestMethod]
        public void TestIsDictionaryStringObject()
        {
            Assert.IsTrue(TypeHelper.IsDictionaryStringObject(typeof(IDictionary<string, object>)));
            Assert.IsTrue(TypeHelper.IsDictionaryStringObject(typeof(IDictionary<string, int>)));
            Assert.IsTrue(TypeHelper.IsDictionaryStringObject(typeof(IDictionary<string, string>)));
            Assert.IsTrue(TypeHelper.IsDictionaryStringObject(typeof(IDictionary<string, decimal>)));
            Assert.IsFalse(TypeHelper.IsDictionaryStringObject(typeof(IDictionary<object, object>)));
            Assert.IsFalse(TypeHelper.IsDictionaryStringObject(typeof(IDictionary<object, string>)));
            Assert.IsFalse(TypeHelper.IsDictionaryStringObject(typeof(IDictionary<int, string>)));
            Assert.IsFalse(TypeHelper.IsDictionaryStringObject(typeof(IEnumerable<object>)));
            Assert.IsFalse(TypeHelper.IsDictionaryStringObject(typeof(string)));
        }
    }
}