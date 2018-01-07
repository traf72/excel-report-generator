﻿using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using ExcelReportGenerator.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Helpers
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
            Assert.IsFalse(TypeHelper.IsKeyValuePair(null));
        }

        [TestMethod]
        public void TestIsEnumerableOfKeyValuePair()
        {
            Assert.IsTrue(TypeHelper.IsEnumerableOfKeyValuePair(typeof(IEnumerable<KeyValuePair<object, object>>)));
            Assert.IsTrue(TypeHelper.IsEnumerableOfKeyValuePair(typeof(IList<KeyValuePair<string, object>>)));
            Assert.IsTrue(TypeHelper.IsEnumerableOfKeyValuePair(typeof(List<KeyValuePair<string, int>>)));
            Assert.IsTrue(TypeHelper.IsEnumerableOfKeyValuePair(typeof(KeyValuePair<int, string>[])));
            Assert.IsTrue(TypeHelper.IsEnumerableOfKeyValuePair(typeof(IDictionary<int, decimal>)));
            Assert.IsFalse(TypeHelper.IsEnumerableOfKeyValuePair(typeof(ArrayList)));
            Assert.IsFalse(TypeHelper.IsEnumerableOfKeyValuePair(typeof(object)));
            Assert.IsFalse(TypeHelper.IsEnumerableOfKeyValuePair(typeof(string)));
            Assert.IsFalse(TypeHelper.IsEnumerableOfKeyValuePair(null));
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
            Assert.IsFalse(TypeHelper.IsDictionaryStringObject(null));
        }

        [TestMethod]
        public void TestTryGetGenericEnumerableInterface()
        {
            var @interface = TypeHelper.TryGetGenericEnumerableInterface(typeof(IEnumerable<string>));
            Assert.AreEqual(typeof(string),@interface.GetGenericArguments()[0]);

            @interface = TypeHelper.TryGetGenericEnumerableInterface(typeof(IList<int>));
            Assert.AreEqual(typeof(int), @interface.GetGenericArguments()[0]);

            @interface = TypeHelper.TryGetGenericEnumerableInterface(typeof(List<Guid>));
            Assert.AreEqual(typeof(Guid), @interface.GetGenericArguments()[0]);

            @interface = TypeHelper.TryGetGenericEnumerableInterface(typeof(decimal[]));
            Assert.AreEqual(typeof(decimal), @interface.GetGenericArguments()[0]);

            @interface = TypeHelper.TryGetGenericEnumerableInterface(typeof(string));
            Assert.AreEqual(typeof(char), @interface.GetGenericArguments()[0]);

            @interface = TypeHelper.TryGetGenericEnumerableInterface(typeof(IDictionary<string, object>));
            Assert.IsTrue(TypeHelper.IsKeyValuePair(@interface.GetGenericArguments()[0]));

            Assert.IsNull(TypeHelper.TryGetGenericEnumerableInterface(typeof(ArrayList)));
            Assert.IsNull(TypeHelper.TryGetGenericEnumerableInterface(typeof(int)));
            Assert.IsNull(TypeHelper.TryGetGenericEnumerableInterface(null));
        }

        [TestMethod]
        public void TestTryGetGenericDictionaryInterface()
        {
            var @interface = TypeHelper.TryGetGenericDictionaryInterface(typeof(IDictionary<string, int>));
            Assert.AreEqual(typeof(string), @interface.GetGenericArguments()[0]);
            Assert.AreEqual(typeof(int), @interface.GetGenericArguments()[1]);

            @interface = TypeHelper.TryGetGenericDictionaryInterface(typeof(Dictionary<Guid?, decimal>));
            Assert.AreEqual(typeof(Guid?), @interface.GetGenericArguments()[0]);
            Assert.AreEqual(typeof(decimal), @interface.GetGenericArguments()[1]);

            @interface = TypeHelper.TryGetGenericDictionaryInterface(typeof(ConcurrentDictionary<short, object>));
            Assert.AreEqual(typeof(short), @interface.GetGenericArguments()[0]);
            Assert.AreEqual(typeof(object), @interface.GetGenericArguments()[1]);

            Assert.IsNull(TypeHelper.TryGetGenericDictionaryInterface(typeof(IEnumerable<string>)));
            Assert.IsNull(TypeHelper.TryGetGenericDictionaryInterface(typeof(int)));
            Assert.IsNull(TypeHelper.TryGetGenericDictionaryInterface(null));
        }

        [TestMethod]
        public void TestIsGenericEnumerable()
        {
            Assert.IsTrue(TypeHelper.IsGenericEnumerable(typeof(IEnumerable<string>)));
            Assert.IsTrue(TypeHelper.IsGenericEnumerable(typeof(IList<int>)));
            Assert.IsTrue(TypeHelper.IsGenericEnumerable(typeof(List<Guid>)));
            Assert.IsTrue(TypeHelper.IsGenericEnumerable(typeof(decimal[])));
            Assert.IsTrue(TypeHelper.IsGenericEnumerable(typeof(string)));
            Assert.IsTrue(TypeHelper.IsGenericEnumerable(typeof(IDictionary<string, object>)));
            Assert.IsFalse(TypeHelper.IsGenericEnumerable(typeof(ArrayList)));
            Assert.IsFalse(TypeHelper.IsGenericEnumerable(typeof(int)));
            Assert.IsFalse(TypeHelper.IsGenericEnumerable(null));
        }
    }
}