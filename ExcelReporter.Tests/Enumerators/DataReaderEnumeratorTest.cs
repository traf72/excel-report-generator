using System;
using System.Data;
using ExcelReporter.Enumerators;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Enumerators
{
    [TestClass]
    public class DataReaderEnumeratorTest
    {
        [TestMethod]
        public void TestEnumerator()
        {
            MyAssert.Throws<ArgumentNullException>(() => new DataReaderEnumerator(null));

            int counter = 0;
            IDataReader reader = Substitute.For<IDataReader>();
            reader.Read().Returns(x =>
            {
                counter++;
                return counter <= 3;
            });

            var enumerator = new DataReaderEnumerator(reader);
            MyAssert.Throws<InvalidOperationException>(() => { IDataReader current = enumerator.Current; }, "Enumerator has not been started. Call MoveNext() method.");
            MyAssert.Throws<InvalidOperationException>(() => { IDataReader current = enumerator.Current; }, "Enumerator has not been started. Call MoveNext() method.");
            while (enumerator.MoveNext())
            {
            }

            MyAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "Enumerator has been finished");
            MyAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "Enumerator has been finished");
            MyAssert.Throws<InvalidOperationException>(() => { IDataReader current = enumerator.Current; }, "Enumerator has been finished");
            MyAssert.Throws<InvalidOperationException>(() => { IDataReader current = enumerator.Current; }, "Enumerator has been finished");
            reader.Received(4).Read();

            MyAssert.Throws<NotSupportedException>(() => enumerator.Reset(), $"{nameof(DataReaderEnumerator)} does not support reset method");

            reader.DidNotReceive().Close();
            enumerator.Dispose();
            reader.Received(1).Close();

            reader.IsClosed.Returns(true);

            MyAssert.Throws<InvalidOperationException>(() => enumerator.MoveNext(), "DataReader has been closed");
            MyAssert.Throws<InvalidOperationException>(() => { IDataReader current = enumerator.Current; }, "DataReader has been closed");
        }
    }
}