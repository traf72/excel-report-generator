using System;
using NUnit.Framework;

namespace ExcelReportGenerator.Tests.CustomAsserts
{
    public class ExceptionAssert
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
                Assert.IsInstanceOf<T>(e, "Wrong type of exception was thrown");
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
                Assert.IsInstanceOf<T>(baseException, "Wrong type of exception was thrown");
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