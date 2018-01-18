using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelReportGenerator.Tests.Rendering
{
    [TestClass]
    public class SystemFunctionsTest
    {
        [TestMethod]
        public void TestGetDictVal()
        {
            var dict = new Dictionary<string, int>
            {
                ["Key1"] = 1,
                ["Key2"] = 2,
            };

            Assert.AreEqual(1, SystemFunctions.GetDictVal(dict, "Key1"));
            Assert.AreEqual(2, SystemFunctions.GetDictVal(dict, "Key2"));

            ExceptionAssert.Throws<KeyNotFoundException>(() => SystemFunctions.GetDictVal(dict, "BadKey"), "The given key \"BadKey\" was not present in the dictionary");
            ExceptionAssert.Throws<ArgumentNullException>(() => SystemFunctions.GetDictVal(null, "Key1"));
            ExceptionAssert.Throws<ArgumentNullException>(() => SystemFunctions.GetDictVal(dict, null));
            ExceptionAssert.Throws<ArgumentException>(() => SystemFunctions.GetDictVal(Guid.NewGuid(), "Key1"), $"Parameter \"dictionary\" must implement {nameof(IDictionary)} interface");

            var objKey = new object();
            var objValue = new Random();
            var dict2 = new Hashtable
            {
                [1] = "One",
                ["Two"] = Guid.Empty,
                [objKey] = objValue,
            };

            Assert.AreEqual("One", SystemFunctions.GetDictVal(dict2, 1));
            Assert.AreEqual(Guid.Empty, SystemFunctions.GetDictVal(dict2, "Two"));
            Assert.AreSame(objValue, SystemFunctions.GetDictVal(dict2, objKey));
        }

        [TestMethod]
        public void TestTryGetDictVal()
        {
            var dict = new Dictionary<string, int>
            {
                ["Key1"] = 1,
                ["Key2"] = 2,
            };

            Assert.AreEqual(1, SystemFunctions.TryGetDictVal(dict, "Key1"));
            Assert.AreEqual(2, SystemFunctions.TryGetDictVal(dict, "Key2"));

            Assert.IsNull(SystemFunctions.TryGetDictVal(dict, "BadKey"));
            Assert.IsNull(SystemFunctions.TryGetDictVal(null, "Key1"));
            Assert.IsNull(SystemFunctions.TryGetDictVal(dict, null));
            Assert.IsNull(SystemFunctions.TryGetDictVal(Guid.NewGuid(), "Key1"));

            var objKey = new object();
            var objValue = new Random();
            var dict2 = new Hashtable
            {
                [1] = "One",
                ["Two"] = Guid.Empty,
                [objKey] = objValue,
            };

            Assert.AreEqual("One", SystemFunctions.TryGetDictVal(dict2, 1));
            Assert.AreEqual(Guid.Empty, SystemFunctions.TryGetDictVal(dict2, "Two"));
            Assert.AreSame(objValue, SystemFunctions.TryGetDictVal(dict2, objKey));
        }

        [TestMethod]
        public void TestFormat()
        {
            Assert.AreEqual("31.01.2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy"));
            Assert.AreEqual("31.01.2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "d"));
            Assert.AreEqual("01/31/2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "d", CultureInfo.InvariantCulture));
            Assert.AreEqual(6535.676.ToString("0,0.##"), SystemFunctions.Format(6535.676, "0,0.##"));
            Assert.AreEqual(6535.676.ToString("0,0.##", CultureInfo.InvariantCulture), SystemFunctions.Format(6535.676, "0,0.##", CultureInfo.InvariantCulture));

            Assert.AreEqual(new DateTime(2018, 1, 31).ToString((string) null), SystemFunctions.Format(new DateTime(2018, 1, 31), null));
            Assert.AreEqual(new DateTime(2018, 1, 31).ToString(null, CultureInfo.InvariantCulture), SystemFunctions.Format(new DateTime(2018, 1, 31), null, CultureInfo.InvariantCulture));
            Assert.IsNull(SystemFunctions.Format(null, "dd.MM.yyyy"));
            Assert.IsNull(SystemFunctions.Format(null, "dd.MM.yyyy", CultureInfo.InvariantCulture));

            ExceptionAssert.Throws<ArgumentException>(() => SystemFunctions.Format(new Random(), "0,0.##"), $"Parameter \"input\" must implement {nameof(IFormattable)} interface");
        }
    }
}