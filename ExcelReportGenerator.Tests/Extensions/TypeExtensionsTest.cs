using ExcelReportGenerator.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Tests.Extensions
{
    [TestClass]
    public class TypeExtensionsTest
    {
        [TestMethod]
        public void TestIsNumeric()
        {
            Assert.IsTrue(typeof(byte).IsNumeric());
            Assert.IsTrue(typeof(byte?).IsNumeric());
            Assert.IsTrue(typeof(ushort).IsNumeric());
            Assert.IsTrue(typeof(ushort?).IsNumeric());
            Assert.IsTrue(typeof(uint).IsNumeric());
            Assert.IsTrue(typeof(uint?).IsNumeric());
            Assert.IsTrue(typeof(ulong).IsNumeric());
            Assert.IsTrue(typeof(ulong?).IsNumeric());
            Assert.IsTrue(typeof(sbyte).IsNumeric());
            Assert.IsTrue(typeof(sbyte?).IsNumeric());
            Assert.IsTrue(typeof(short).IsNumeric());
            Assert.IsTrue(typeof(short?).IsNumeric());
            Assert.IsTrue(typeof(int).IsNumeric());
            Assert.IsTrue(typeof(int?).IsNumeric());
            Assert.IsTrue(typeof(long).IsNumeric());
            Assert.IsTrue(typeof(long?).IsNumeric());
            Assert.IsTrue(typeof(float).IsNumeric());
            Assert.IsTrue(typeof(float?).IsNumeric());
            Assert.IsTrue(typeof(double).IsNumeric());
            Assert.IsTrue(typeof(double?).IsNumeric());
            Assert.IsTrue(typeof(decimal).IsNumeric());
            Assert.IsTrue(typeof(decimal?).IsNumeric());

            Assert.IsFalse(typeof(object).IsNumeric());
            Assert.IsFalse(typeof(string).IsNumeric());
            Assert.IsFalse(typeof(char).IsNumeric());
            Assert.IsFalse(typeof(bool).IsNumeric());
            Assert.IsFalse(typeof(Regex).IsNumeric());
        }

        [TestMethod]
        public void TestIsExtendedPrimitive()
        {
            Assert.IsTrue(typeof(byte).IsExtendedPrimitive());
            Assert.IsTrue(typeof(byte?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(ushort).IsExtendedPrimitive());
            Assert.IsTrue(typeof(ushort?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(uint).IsExtendedPrimitive());
            Assert.IsTrue(typeof(uint?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(ulong).IsExtendedPrimitive());
            Assert.IsTrue(typeof(ulong?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(sbyte).IsExtendedPrimitive());
            Assert.IsTrue(typeof(sbyte?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(short).IsExtendedPrimitive());
            Assert.IsTrue(typeof(short?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(int).IsExtendedPrimitive());
            Assert.IsTrue(typeof(int?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(long).IsExtendedPrimitive());
            Assert.IsTrue(typeof(long?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(float).IsExtendedPrimitive());
            Assert.IsTrue(typeof(float?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(double).IsExtendedPrimitive());
            Assert.IsTrue(typeof(double?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(decimal).IsExtendedPrimitive());
            Assert.IsTrue(typeof(decimal?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(char).IsExtendedPrimitive());
            Assert.IsTrue(typeof(char?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(bool).IsExtendedPrimitive());
            Assert.IsTrue(typeof(bool?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(string).IsExtendedPrimitive());
            Assert.IsTrue(typeof(Guid).IsExtendedPrimitive());
            Assert.IsTrue(typeof(Guid?).IsExtendedPrimitive());
            Assert.IsTrue(typeof(DateTime).IsExtendedPrimitive());
            Assert.IsTrue(typeof(DateTime?).IsExtendedPrimitive());

            Assert.IsFalse(typeof(object).IsExtendedPrimitive());
            Assert.IsFalse(typeof(Regex).IsExtendedPrimitive());
        }
    }
}