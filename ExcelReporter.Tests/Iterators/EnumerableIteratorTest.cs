using ExcelReporter.Iterators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace ExcelReporter.Tests.Iterators
{
    [TestClass]
    public class EnumerableIteratorTest
    {
        [TestMethod]
        public void TestIterator()
        {
            int[] array = { 1, 2, 3 };
            IList<int> result = new List<int>(array.Length);
            var iterator = new EnumerableIterator<int>(array);
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next());
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next());
            while (iterator.HaxNext())
            {
                result.Add(iterator.Next());
            }

            MyAssert.Throws<InvalidOperationException>(() => iterator.Next());
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next());
            Assert.IsFalse(iterator.HaxNext());
            Assert.IsFalse(iterator.HaxNext());
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual(1, result[0]);
            Assert.AreEqual(2, result[1]);
            Assert.AreEqual(3, result[2]);

            iterator.Reset();
            Assert.IsTrue(iterator.HaxNext());
            Assert.IsTrue(iterator.HaxNext());
            Assert.AreEqual(2, iterator.Next());

            result.Clear();
            array = new int[0];
            iterator = new EnumerableIterator<int>(array);
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next());
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next());
            while (iterator.HaxNext())
            {
                result.Add(iterator.Next());
            }

            MyAssert.Throws<InvalidOperationException>(() => iterator.Next());
            MyAssert.Throws<InvalidOperationException>(() => iterator.Next());
            Assert.IsFalse(iterator.HaxNext());
            Assert.IsFalse(iterator.HaxNext());
            Assert.AreEqual(0, result.Count);

            MyAssert.Throws<ArgumentNullException>(() => new EnumerableIterator<object>(null));
        }
    }
}