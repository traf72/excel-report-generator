using ExcelReportGenerator.Rendering.Providers;
using ExcelReportGenerator.Tests.CustomAsserts;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering.Providers;

public class DefaultInstanceProviderTest
{
    [Test]
    public void TestGetInstance()
    {
        IInstanceProvider instanceProvider = new DefaultInstanceProvider();
        ExceptionAssert.Throws<InvalidOperationException>(() => instanceProvider.GetInstance(null),
            "Type is not specified but defaultInstance is null");

        var instance1 = (TestType_3) instanceProvider.GetInstance(typeof(TestType_3));
        var instance2 = (TestType_3) instanceProvider.GetInstance(typeof(TestType_3));
        var instance3 = instanceProvider.GetInstance<TestType_3>();
        Assert.AreSame(instance1, instance2);
        Assert.AreSame(instance1, instance3);
        Assert.AreSame(instance2, instance3);

        Assert.IsInstanceOf<TestType_5>(instanceProvider.GetInstance(typeof(TestType_5)));
        Assert.IsInstanceOf<DateTime>(instanceProvider.GetInstance(typeof(DateTime)));

        ExceptionAssert.Throws<MissingMethodException>(() => instanceProvider.GetInstance(typeof(FileInfo)));
        ExceptionAssert.Throws<MissingMethodException>(() => instanceProvider.GetInstance(typeof(Math)));

        var testInstance = new TestType_3();
        instanceProvider = new DefaultInstanceProvider(testInstance);

        instance3 = (TestType_3) instanceProvider.GetInstance(null);
        instance1 = (TestType_3) instanceProvider.GetInstance(typeof(TestType_3));
        instance2 = instanceProvider.GetInstance<TestType_3>();
        Assert.AreSame(instance1, instance2);
        Assert.AreSame(instance1, instance3);
        Assert.AreSame(instance2, instance3);
    }
}