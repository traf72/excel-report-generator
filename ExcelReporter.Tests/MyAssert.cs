using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests
{
    public class MyAssert
    {
        public static void Throws<T>(Action action, string expectedMessage = null) where T : Exception
        {
            try
            {
                action.Invoke();
            }
            catch (T e)
            {
                if (expectedMessage != null)
                {
                    Assert.AreEqual(expectedMessage, e.Message, "Wrong exception message was returned");
                }
                return;
            }
            catch (Exception e)
            {
                Assert.IsInstanceOfType(e, typeof(T), "Wrong type of exception was thrown");
                return;
            }

            Assert.Fail("No exception was thrown");
        }

        public static void ThrowsBaseException<T>(Action action, string expectedMessage = null) where T : Exception
        {
            try
            {
                action.Invoke();
            }
            catch (Exception e)
            {
                Exception baseException = e.GetBaseException();
                Assert.IsInstanceOfType(baseException, typeof(T), "Wrong type of exception was thrown");
                if (expectedMessage != null)
                {
                    Assert.AreEqual(expectedMessage, baseException.Message, "Wrong exception message was returned");
                }
                return;
            }

            Assert.Fail("No exception was thrown");
        }
    }
}