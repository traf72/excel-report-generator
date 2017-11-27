using ExcelReporter.Extensions;
using ExcelReporter.Rendering.TemplateProcessors;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;

namespace ExcelReporter.Tests.Extensions
{
    [TestClass]
    public class ITemplateProcessorExtensionsTest
    {
        [TestMethod]
        public void TestGetTemplateWithoutBorders()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");

            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("{p:Name}"));
            Assert.AreEqual("p:Name>", processor.GetTemplateWithoutBorders("{p:Name>"));
            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("  {p:Name}  "));
            Assert.AreEqual("p:Name ", processor.GetTemplateWithoutBorders("  {p:Name }  "));
            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders(" p:Name "));
            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("p:Name"));

            processor.LeftTemplateBorder.Returns("@<");
            processor.RightTemplateBorder.Returns(")&");

            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("@<p:Name)&"));
            Assert.AreEqual("p:Name)>", processor.GetTemplateWithoutBorders("@<p:Name)>"));
            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("  @<p:Name)&  "));
            Assert.AreEqual(" p:Name", processor.GetTemplateWithoutBorders("  @< p:Name)&  "));
            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders(" p:Name "));
            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("p:Name"));

            processor.LeftTemplateBorder.Returns(string.Empty);
            processor.RightTemplateBorder.Returns(string.Empty);

            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("p:Name"));
            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("p:Name"));
            Assert.AreEqual("{p:Name}", processor.GetTemplateWithoutBorders("{p:Name}"));

            processor.LeftTemplateBorder.Returns((string) null);
            processor.RightTemplateBorder.Returns((string) null);

            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("p:Name"));
            Assert.AreEqual("p:Name", processor.GetTemplateWithoutBorders("  p:Name  "));
            Assert.AreEqual("{p:Name}", processor.GetTemplateWithoutBorders("{p:Name}"));
        }
    }
}