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
        public void TestUnwrapTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");

            Assert.AreEqual("p:Name", processor.UnwrapTemplate("{p:Name}"));
            Assert.AreEqual("p:Name>", processor.UnwrapTemplate("{p:Name>"));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate("  {p:Name}  "));
            Assert.AreEqual("p:Name ", processor.UnwrapTemplate("  {p:Name }  "));
            Assert.AreEqual(string.Empty, processor.UnwrapTemplate("{}"));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate(" p:Name "));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate("p:Name"));

            processor.LeftTemplateBorder.Returns("@<");
            processor.RightTemplateBorder.Returns(")&");

            Assert.AreEqual("p:Name", processor.UnwrapTemplate("@<p:Name)&"));
            Assert.AreEqual("p:Name)>", processor.UnwrapTemplate("@<p:Name)>"));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate("  @<p:Name)&  "));
            Assert.AreEqual(" p:Name", processor.UnwrapTemplate("  @< p:Name)&  "));
            Assert.AreEqual(string.Empty, processor.UnwrapTemplate("@<)&"));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate(" p:Name "));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate("p:Name"));

            processor.LeftTemplateBorder.Returns(string.Empty);
            processor.RightTemplateBorder.Returns(string.Empty);

            Assert.AreEqual("p:Name", processor.UnwrapTemplate("p:Name"));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate("  p:Name  "));
            Assert.AreEqual("{p:Name}", processor.UnwrapTemplate("{p:Name}"));

            processor.LeftTemplateBorder.Returns((string) null);
            processor.RightTemplateBorder.Returns((string) null);

            Assert.AreEqual("p:Name", processor.UnwrapTemplate("p:Name"));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate("  p:Name  "));
            Assert.AreEqual("{p:Name}", processor.UnwrapTemplate("{p:Name}"));
        }

        [TestMethod]
        public void TestWrapTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");

            Assert.AreEqual("{p:Name}", processor.WrapTemplate("p:Name"));
            Assert.AreEqual("{ p:Name }", processor.WrapTemplate(" p:Name "));
            Assert.AreEqual("{}", processor.WrapTemplate(""));

            processor.LeftTemplateBorder.Returns("@<");
            processor.RightTemplateBorder.Returns(")&");

            Assert.AreEqual("@<p:Name)&", processor.WrapTemplate("p:Name"));
            Assert.AreEqual("@<  p:Name  )&", processor.WrapTemplate("  p:Name  "));
            Assert.AreEqual("@<)&", processor.WrapTemplate(""));

            processor.LeftTemplateBorder.Returns(string.Empty);
            processor.RightTemplateBorder.Returns(string.Empty);

            Assert.AreEqual("p:Name", processor.WrapTemplate("p:Name"));
            Assert.AreEqual("  p:Name  ", processor.WrapTemplate("  p:Name  "));
            Assert.AreEqual("{p:Name}", processor.WrapTemplate("{p:Name}"));

            processor.LeftTemplateBorder.Returns((string)null);
            processor.RightTemplateBorder.Returns((string)null);

            Assert.AreEqual("p:Name", processor.WrapTemplate("p:Name"));
            Assert.AreEqual("  p:Name  ", processor.WrapTemplate("  p:Name  "));
            Assert.AreEqual("{p:Name}", processor.WrapTemplate("{p:Name}"));
        }
    }
}