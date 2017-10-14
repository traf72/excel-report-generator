using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReporter.Tests.Helpers
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
    }
}