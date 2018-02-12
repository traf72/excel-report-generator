using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Extensions;
using ExcelReportGenerator.Rendering.TemplateProcessors;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelReportGenerator.Tests.Extensions
{
    [TestClass]
    public class TemplateProcessorExtensionsTest
    {
        [TestMethod]
        public void TestUnwrapTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();

            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.UnwrapTemplate("{p:Name}"));

            processor.RightTemplateBorder.Returns("}");
            processor.LeftTemplateBorder.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.UnwrapTemplate("{p:Name}"));

            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");

            ExceptionAssert.Throws<ArgumentNullException>(() => processor.UnwrapTemplate(null));
            Assert.AreEqual(string.Empty, processor.UnwrapTemplate(string.Empty));
            Assert.AreEqual(" ", processor.UnwrapTemplate(" "));

            Assert.AreEqual("p:Name", processor.UnwrapTemplate("{p:Name}"));
            Assert.AreEqual("\\{p:Name", processor.UnwrapTemplate(Regex.Escape("{p:Name}")));
            Assert.AreEqual(Regex.Escape("p**Name"), processor.UnwrapTemplate(Regex.Escape("{p**Name}"), true));
            Assert.AreEqual("p:Name>", processor.UnwrapTemplate("{p:Name>"));
            Assert.AreEqual(" p:Name ", processor.UnwrapTemplate("{ p:Name }"));
            Assert.AreEqual(" {p:Name} ", processor.UnwrapTemplate(" {p:Name} "));
            Assert.AreEqual(string.Empty, processor.UnwrapTemplate("{}"));
            Assert.AreEqual(" ", processor.UnwrapTemplate("{ }"));
            Assert.AreEqual("\\{\\ ", processor.UnwrapTemplate(Regex.Escape("{ }")));
            Assert.AreEqual(Regex.Escape(" "), processor.UnwrapTemplate(Regex.Escape("{ }"), true));
            Assert.AreEqual(" p:Name ", processor.UnwrapTemplate(" p:Name "));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate("p:Name"));

            processor.LeftTemplateBorder.Returns("@<");
            processor.RightTemplateBorder.Returns(")&");

            Assert.AreEqual("p:Name", processor.UnwrapTemplate("@<p:Name)&"));
            Assert.AreEqual("p:Name\\", processor.UnwrapTemplate(Regex.Escape("@<p:Name)&")));
            Assert.AreEqual(Regex.Escape("p[Name"), processor.UnwrapTemplate(Regex.Escape("@<p[Name)&"), true));
            Assert.AreEqual("p:Name)>", processor.UnwrapTemplate("@<p:Name)>"));
            Assert.AreEqual(" p:Name ", processor.UnwrapTemplate("@< p:Name )&"));
            Assert.AreEqual("  @<p:Name)&  ", processor.UnwrapTemplate("  @<p:Name)&  "));
            Assert.AreEqual(string.Empty, processor.UnwrapTemplate("@<)&"));
            Assert.AreEqual("  ", processor.UnwrapTemplate("@<  )&"));
            Assert.AreEqual(" p:Name ", processor.UnwrapTemplate(" p:Name "));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate("p:Name"));

            processor.LeftTemplateBorder.Returns(string.Empty);
            processor.RightTemplateBorder.Returns(string.Empty);

            Assert.AreEqual("p:Name", processor.UnwrapTemplate("p:Name"));
            Assert.AreEqual(" p:Name ", processor.UnwrapTemplate(" p:Name "));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate(Regex.Escape("p:Name")));

            processor.LeftTemplateBorder.Returns(" ");
            processor.RightTemplateBorder.Returns(" ");

            Assert.AreEqual("p:Name", processor.UnwrapTemplate("p:Name"));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate(" p:Name "));
            Assert.AreEqual(" p:Name ", processor.UnwrapTemplate("  p:Name  "));
            Assert.AreEqual("p:Name", processor.UnwrapTemplate(Regex.Escape("p:Name")));
            Assert.AreEqual("\\ p\\*Name\\", processor.UnwrapTemplate(Regex.Escape(" p*Name ")));
            Assert.AreEqual(Regex.Escape("p*Name"), processor.UnwrapTemplate(Regex.Escape(" p*Name "), true));
        }

        [TestMethod]
        public void TestWrapTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");

            Assert.AreEqual("{p:Name}", processor.WrapTemplate("p:Name"));
            Assert.AreEqual("\\{p:Name}", processor.WrapTemplate("p:Name", true));
            Assert.AreEqual("{ p:Name }", processor.WrapTemplate(" p:Name "));
            Assert.AreEqual("{}", processor.WrapTemplate(""));

            processor.LeftTemplateBorder.Returns("@<");
            processor.RightTemplateBorder.Returns(")&");

            Assert.AreEqual("@<p:Name)&", processor.WrapTemplate("p:Name"));
            Assert.AreEqual("@<p:Name\\)&", processor.WrapTemplate("p:Name", true));
            Assert.AreEqual("@<  p:Name  )&", processor.WrapTemplate("  p:Name  "));
            Assert.AreEqual("@<)&", processor.WrapTemplate(""));

            processor.LeftTemplateBorder.Returns(string.Empty);
            processor.RightTemplateBorder.Returns(string.Empty);

            Assert.AreEqual("p:Name", processor.WrapTemplate("p:Name"));
            Assert.AreEqual("p:Name", processor.WrapTemplate("p:Name", true));
            Assert.AreEqual("  p:Name  ", processor.WrapTemplate("  p:Name  "));
            Assert.AreEqual("{p:Name}", processor.WrapTemplate("{p:Name}"));

            processor.LeftTemplateBorder.Returns((string)null);
            processor.RightTemplateBorder.Returns((string)null);

            Assert.AreEqual("p:Name", processor.WrapTemplate("p:Name"));
            Assert.AreEqual(" p:Name ", processor.WrapTemplate(" p:Name "));
            Assert.AreEqual("", processor.WrapTemplate(""));

            ExceptionAssert.Throws<ArgumentNullException>(() => processor.WrapTemplate("p:Name", true));
        }

        [TestMethod]
        public void TestBuildPropertyTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.PropertyMemberLabel.Returns("p");

            Assert.AreEqual("{p:Name}", processor.BuildPropertyTemplate("Name"));
            Assert.AreEqual("{p: Contacts.Phone }", processor.BuildPropertyTemplate(" Contacts.Phone "));
            Assert.AreEqual("{p:}", processor.BuildPropertyTemplate(string.Empty));
            Assert.AreEqual("{p: }", processor.BuildPropertyTemplate(" "));
            Assert.AreEqual("{p:}", processor.BuildPropertyTemplate(null));

            processor.LeftTemplateBorder.Returns((string)null);
            processor.RightTemplateBorder.Returns((string)null);
            processor.MemberLabelSeparator.Returns((string)null);
            processor.PropertyMemberLabel.Returns((string)null);

            Assert.AreEqual("Name", processor.BuildPropertyTemplate("Name"));
            Assert.AreEqual("Contacts.Phone", processor.BuildPropertyTemplate("Contacts.Phone"));
            Assert.AreEqual(string.Empty, processor.BuildPropertyTemplate(string.Empty));
            Assert.AreEqual(" ", processor.BuildPropertyTemplate(" "));
            Assert.AreEqual(string.Empty, processor.BuildPropertyTemplate(null));
        }

        [TestMethod]
        public void TestBuildMethodCallTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.MethodCallMemberLabel.Returns("m");

            Assert.AreEqual("{m: Meth() }", processor.BuildMethodCallTemplate(" Meth() "));
            Assert.AreEqual("{m:Meth(p:Name, di:Value, m:MethInner())}", processor.BuildMethodCallTemplate("Meth(p:Name, di:Value, m:MethInner())"));
            Assert.AreEqual("{m:}", processor.BuildMethodCallTemplate(string.Empty));
            Assert.AreEqual("{m: }", processor.BuildMethodCallTemplate(" "));
            Assert.AreEqual("{m:}", processor.BuildMethodCallTemplate(null));

            processor.LeftTemplateBorder.Returns((string)null);
            processor.RightTemplateBorder.Returns((string)null);
            processor.MemberLabelSeparator.Returns((string)null);
            processor.MethodCallMemberLabel.Returns((string)null);

            Assert.AreEqual("Meth()", processor.BuildMethodCallTemplate("Meth()"));
            Assert.AreEqual("Meth(p:Name, di:Value, MethInner())", processor.BuildMethodCallTemplate("Meth(p:Name, di:Value, MethInner())"));
            Assert.AreEqual(string.Empty, processor.BuildMethodCallTemplate(string.Empty));
            Assert.AreEqual(" ", processor.BuildMethodCallTemplate(" "));
            Assert.AreEqual(string.Empty, processor.BuildMethodCallTemplate(null));
        }

        [TestMethod]
        public void TestBuildDataItemTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.DataItemMemberLabel.Returns("di");

            Assert.AreEqual("{di:Name}", processor.BuildDataItemTemplate("Name"));
            Assert.AreEqual("{di: Contacts.Phone }", processor.BuildDataItemTemplate(" Contacts.Phone "));
            Assert.AreEqual("{di:}", processor.BuildDataItemTemplate(string.Empty));
            Assert.AreEqual("{di: }", processor.BuildDataItemTemplate(" "));
            Assert.AreEqual("{di:}", processor.BuildDataItemTemplate(null));

            processor.LeftTemplateBorder.Returns((string)null);
            processor.RightTemplateBorder.Returns((string)null);
            processor.MemberLabelSeparator.Returns((string)null);
            processor.DataItemMemberLabel.Returns((string)null);

            Assert.AreEqual("Name", processor.BuildDataItemTemplate("Name"));
            Assert.AreEqual("Contacts.Phone", processor.BuildDataItemTemplate("Contacts.Phone"));
            Assert.AreEqual(string.Empty, processor.BuildDataItemTemplate(string.Empty));
            Assert.AreEqual(" ", processor.BuildDataItemTemplate(" "));
            Assert.AreEqual(string.Empty, processor.BuildDataItemTemplate(null));
        }

        [TestMethod]
        public void TestBuildVariableTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.SystemVariableMemberLabel.Returns("sv");

            Assert.AreEqual("{sv:Now}", processor.BuildVariableTemplate("Now"));
            Assert.AreEqual("{sv:}", processor.BuildVariableTemplate(string.Empty));
            Assert.AreEqual("{sv: }", processor.BuildVariableTemplate(" "));
            Assert.AreEqual("{sv:}", processor.BuildVariableTemplate(null));

            processor.LeftTemplateBorder.Returns((string)null);
            processor.RightTemplateBorder.Returns((string)null);
            processor.MemberLabelSeparator.Returns((string)null);
            processor.SystemVariableMemberLabel.Returns((string)null);

            Assert.AreEqual("Now", processor.BuildVariableTemplate("Now"));
            Assert.AreEqual(string.Empty, processor.BuildVariableTemplate(string.Empty));
            Assert.AreEqual(" ", processor.BuildVariableTemplate(" "));
            Assert.AreEqual(string.Empty, processor.BuildVariableTemplate(null));
        }

        [TestMethod]
        public void TestBuildSystemFunctionTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.SystemFunctionMemberLabel.Returns("sf");

            Assert.AreEqual("{sf:Format(p:Prop, 0)}", processor.BuildSystemFunctionTemplate("Format(p:Prop, 0)"));
            Assert.AreEqual("{sf:}", processor.BuildSystemFunctionTemplate(string.Empty));
            Assert.AreEqual("{sf: }", processor.BuildSystemFunctionTemplate(" "));
            Assert.AreEqual("{sf:}", processor.BuildSystemFunctionTemplate(null));

            processor.LeftTemplateBorder.Returns((string)null);
            processor.RightTemplateBorder.Returns((string)null);
            processor.MemberLabelSeparator.Returns((string)null);
            processor.SystemFunctionMemberLabel.Returns((string)null);

            Assert.AreEqual("Format(p:Prop, 0)", processor.BuildSystemFunctionTemplate("Format(p:Prop, 0)"));
            Assert.AreEqual(string.Empty, processor.BuildSystemFunctionTemplate(string.Empty));
            Assert.AreEqual(" ", processor.BuildSystemFunctionTemplate(" "));
            Assert.AreEqual(string.Empty, processor.BuildSystemFunctionTemplate(null));
        }

        [TestMethod]
        public void TestTrimPropertyLabel()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.MemberLabelSeparator.Returns(":");
            processor.PropertyMemberLabel.Returns("p");

            Assert.AreEqual("Name", processor.TrimPropertyLabel("p:Name"));
            Assert.AreEqual("  Name p:Phone", processor.TrimPropertyLabel(" p : Name p:Phone"));
            Assert.AreEqual("di:Name", processor.TrimPropertyLabel("di:Name"));
            Assert.AreEqual("Contacts.Phone", processor.TrimPropertyLabel("Contacts.Phone"));
            Assert.AreEqual(" { Name } ", processor.TrimPropertyLabel(" { p:Name } "));
            Assert.AreEqual(" { Name } ", processor.TrimPropertyLabel(" { P :Name } "));
            Assert.AreEqual("{Name p:Phone}", processor.TrimPropertyLabel("{p:Name p:Phone}"));
            Assert.AreEqual(" { di:Name } ", processor.TrimPropertyLabel(" { di:Name } "));
            Assert.AreEqual("{Contacts.Phone}", processor.TrimPropertyLabel("{Contacts.Phone}"));
            ExceptionAssert.Throws<ArgumentNullException>(() => processor.TrimPropertyLabel(null));

            processor.MemberLabelSeparator.Returns(string.Empty);
            Assert.AreEqual("Name", processor.TrimPropertyLabel("pName"));
            Assert.AreEqual("{Name}", processor.TrimPropertyLabel("{pName}"));

            processor.PropertyMemberLabel.Returns(string.Empty);
            Assert.AreEqual("p:Name", processor.TrimPropertyLabel("p:Name"));
        }

        [TestMethod]
        public void TestTrimDataItemLabel()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.MemberLabelSeparator.Returns(":");
            processor.DataItemMemberLabel.Returns("di");

            Assert.AreEqual("Name", processor.TrimDataItemLabel("DI:Name"));
            Assert.AreEqual("Name di:Phone", processor.TrimDataItemLabel("di:Name di:Phone"));
            Assert.AreEqual("p:Name", processor.TrimDataItemLabel("p:Name"));
            Assert.AreEqual("Contacts.Phone", processor.TrimDataItemLabel("Contacts.Phone"));
            Assert.AreEqual(" {  Name } ", processor.TrimDataItemLabel(" { di : Name } "));
            Assert.AreEqual("{Name di:Phone}", processor.TrimDataItemLabel("{di:Name di:Phone}"));
            Assert.AreEqual(" { p:Name } ", processor.TrimDataItemLabel(" { p:Name } "));
            Assert.AreEqual("{Contacts.Phone}", processor.TrimDataItemLabel("{Contacts.Phone}"));
            ExceptionAssert.Throws<ArgumentNullException>(() => processor.TrimDataItemLabel(null));

            processor.MemberLabelSeparator.Returns(string.Empty);
            Assert.AreEqual("Name", processor.TrimDataItemLabel("diName"));
            Assert.AreEqual("{Name}", processor.TrimDataItemLabel("{diName}"));

            processor.DataItemMemberLabel.Returns(string.Empty);
            Assert.AreEqual("di:Name", processor.TrimDataItemLabel("di:Name"));
        }

        [TestMethod]
        public void TestTrimMethodCallLabel()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.MemberLabelSeparator.Returns(":");
            processor.MethodCallMemberLabel.Returns("m");

            Assert.AreEqual("Meth()", processor.TrimMethodCallLabel("m:Meth()"));
            Assert.AreEqual("Meth1(p:Name, di:Value, m:MethInner()) m:Meth2()", processor.TrimMethodCallLabel("m:Meth1(p:Name, di:Value, m:MethInner()) m:Meth2()"));
            Assert.AreEqual("p:Name", processor.TrimMethodCallLabel("p:Name"));
            Assert.AreEqual(" { Name } ", processor.TrimMethodCallLabel(" { m:Name } "));
            Assert.AreEqual(" {  Name } ", processor.TrimMethodCallLabel(" { m: Name } "));
            Assert.AreEqual("{Meth1(p:Name, di:Value, m:MethInner()) m:Meth2()}", processor.TrimMethodCallLabel("{m:Meth1(p:Name, di:Value, m:MethInner()) m:Meth2()}"));
            Assert.AreEqual(" { p:Name } ", processor.TrimMethodCallLabel(" { p:Name } "));
            ExceptionAssert.Throws<ArgumentNullException>(() => processor.TrimMethodCallLabel(null));

            processor.MemberLabelSeparator.Returns(string.Empty);
            Assert.AreEqual("Meth()", processor.TrimMethodCallLabel("mMeth()"));
            Assert.AreEqual("{Meth()}", processor.TrimMethodCallLabel("{mMeth()}"));

            processor.MethodCallMemberLabel.Returns(string.Empty);
            Assert.AreEqual("m:Meth()", processor.TrimMethodCallLabel("m:Meth()"));
        }

        [TestMethod]
        public void TestTrimVariableLabel()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.MemberLabelSeparator.Returns(":");
            processor.SystemVariableMemberLabel.Returns("sv");

            Assert.AreEqual("Now", processor.TrimVariableLabel("sv:Now"));
            Assert.AreEqual(" { SheetName } ", processor.TrimVariableLabel(" { sv:SheetName } "));
            Assert.AreEqual(" {  SheetName } ", processor.TrimVariableLabel(" { sv : SheetName } "));
            ExceptionAssert.Throws<ArgumentNullException>(() => processor.TrimVariableLabel(null));

            processor.MemberLabelSeparator.Returns(string.Empty);
            Assert.AreEqual("Now", processor.TrimVariableLabel("svNow"));
            Assert.AreEqual("{Now}", processor.TrimVariableLabel("{svNow}"));

            processor.SystemVariableMemberLabel.Returns(string.Empty);
            Assert.AreEqual("sv:Now", processor.TrimVariableLabel("sv:Now"));
        }

        [TestMethod]
        public void TestTrimSystemFunctionLabel()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.MemberLabelSeparator.Returns(":");
            processor.SystemFunctionMemberLabel.Returns("sf");

            Assert.AreEqual("Format(p:Prop, 0)", processor.TrimSystemFunctionLabel("sf:Format(p:Prop, 0)"));
            Assert.AreEqual(" { Format(p:Prop, 0) } ", processor.TrimSystemFunctionLabel(" { sf:Format(p:Prop, 0) } "));
            ExceptionAssert.Throws<ArgumentNullException>(() => processor.TrimSystemFunctionLabel(null));

            processor.MemberLabelSeparator.Returns(string.Empty);
            Assert.AreEqual("Format(p:Prop, 0)", processor.TrimSystemFunctionLabel("sfFormat(p:Prop, 0)"));
            Assert.AreEqual("{Format(p:Prop, 0)}", processor.TrimSystemFunctionLabel("{sfFormat(p:Prop, 0)}"));

            processor.SystemFunctionMemberLabel.Returns(string.Empty);
            Assert.AreEqual("sf:Format(p:Prop, 0)", processor.TrimSystemFunctionLabel("sf:Format(p:Prop, 0)"));
        }

        [TestMethod]
        public void TestGetFullRegexPattern()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();

            // Standard case
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.PropertyMemberLabel.Returns("p");
            processor.DataItemMemberLabel.Returns("di");
            processor.MethodCallMemberLabel.Returns("m");
            processor.SystemVariableMemberLabel.Returns("sv");
            processor.SystemFunctionMemberLabel.Returns("sf");

            string pattern = processor.GetFullRegexPattern();
            Assert.AreEqual("\\{\\s*(p|di|m|sv|sf):.+?\\s*}", pattern);

            MatchCollection matches = Regex.Matches("{p:Name}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{p:Name}", matches[0].Value);

            matches = Regex.Matches("{ p:Employee.Contacts.Phone }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ p:Employee.Contacts.Phone }", matches[0].Value);

            matches = Regex.Matches("{di:Name}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{di:Name}", matches[0].Value);

            matches = Regex.Matches("{di:Employee.Contacts.Phone}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{di:Employee.Contacts.Phone}", matches[0].Value);

            matches = Regex.Matches("{m:Meth()}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{m:Meth()}", matches[0].Value);

            matches = Regex.Matches("{  m:Meth(p:Value, di:Name, 5, \"Hi\", m:Meth2(m:Meth3()))  }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{  m:Meth(p:Value, di:Name, 5, \"Hi\", m:Meth2(m:Meth3()))  }", matches[0].Value);

            matches = Regex.Matches("One {p:Name} {Two} Three{di:Employee.Contacts.Phone} {n:Four} {m:Meth1(hi, 5, p:Value, m:Meth2(m:Meth3(di:Coplex.Field)))} {ms:Method()}", pattern);
            Assert.AreEqual(3, matches.Count);
            Assert.AreEqual("{p:Name}", matches[0].Value);
            Assert.AreEqual("{di:Employee.Contacts.Phone}", matches[1].Value);
            Assert.AreEqual("{m:Meth1(hi, 5, p:Value, m:Meth2(m:Meth3(di:Coplex.Field)))}", matches[2].Value);

            matches = Regex.Matches("{sv:Now}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{sv:Now}", matches[0].Value);

            matches = Regex.Matches("{sf:Format(p:Prop, 0,##)}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{sf:Format(p:Prop, 0,##)}", matches[0].Value);

            matches = Regex.Matches("{ms:Meth()}", pattern);
            Assert.AreEqual(0, matches.Count);

            matches = Regex.Matches("{d:Name}", pattern);
            Assert.AreEqual(0, matches.Count);

            matches = Regex.Matches("{s:Name}", pattern);
            Assert.AreEqual(0, matches.Count);

            matches = Regex.Matches("{Name}", pattern);
            Assert.AreEqual(0, matches.Count);

            matches = Regex.Matches("{p:}", pattern);
            Assert.AreEqual(0, matches.Count);

            matches = Regex.Matches("{{p:Name}}", pattern);
            Assert.AreEqual(1, matches.Count);

            // Overriden props case 1
            processor.LeftTemplateBorder.Returns("[");
            processor.RightTemplateBorder.Returns("]");
            processor.MemberLabelSeparator.Returns("*");
            processor.PropertyMemberLabel.Returns("prop");
            processor.DataItemMemberLabel.Returns("data");
            processor.MethodCallMemberLabel.Returns("meth");
            processor.SystemVariableMemberLabel.Returns("sysv");
            processor.SystemFunctionMemberLabel.Returns("sysf");

            pattern = processor.GetFullRegexPattern();
            Assert.AreEqual("\\[\\s*(prop|data|meth|sysv|sysf)\\*.+?\\s*]", pattern);

            matches = Regex.Matches("[prop*Name]", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("[prop*Name]", matches[0].Value);

            matches = Regex.Matches("[prop*Employee.Contacts.Phone]", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("[prop*Employee.Contacts.Phone]", matches[0].Value);

            matches = Regex.Matches("[data*Name]", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("[data*Name]", matches[0].Value);

            matches = Regex.Matches("[data*Employee.Contacts.Phone]", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("[data*Employee.Contacts.Phone]", matches[0].Value);

            matches = Regex.Matches("[ meth*Meth() ]", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("[ meth*Meth() ]", matches[0].Value);

            matches = Regex.Matches("[meth*Meth(prop*Value, data*Name, 5, \"Hi\", meth*Meth2(meth*Meth3()))]", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("[meth*Meth(prop*Value, data*Name, 5, \"Hi\", meth*Meth2(meth*Meth3()))]", matches[0].Value);

            matches = Regex.Matches("One [prop*Name] [Two] Three [data*Employee.Contacts.Phone] [n*Four] [meth*Meth1(hi, 5, prop*Value, meth*Meth2(meth*Meth3(data:Coplex.Field)))] [ms*Method()]", pattern);
            Assert.AreEqual(3, matches.Count);
            Assert.AreEqual("[prop*Name]", matches[0].Value);
            Assert.AreEqual("[data*Employee.Contacts.Phone]", matches[1].Value);
            Assert.AreEqual("[meth*Meth1(hi, 5, prop*Value, meth*Meth2(meth*Meth3(data:Coplex.Field)))]", matches[2].Value);

            matches = Regex.Matches("[sysv*Now]", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("[sysv*Now]", matches[0].Value);

            matches = Regex.Matches("[sysf*Format(prop*Prop, 0,##)]", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("[sysf*Format(prop*Prop, 0,##)]", matches[0].Value);

            matches = Regex.Matches("[[prop*Name]]", pattern);
            Assert.AreEqual(1, matches.Count);

            // Overriden props case 2
            processor.LeftTemplateBorder.Returns("<<");
            processor.RightTemplateBorder.Returns("@@");
            processor.MemberLabelSeparator.Returns("&&");
            processor.PropertyMemberLabel.Returns("prop");
            processor.DataItemMemberLabel.Returns("data");
            processor.MethodCallMemberLabel.Returns("meth");
            processor.SystemVariableMemberLabel.Returns("sysv");
            processor.SystemFunctionMemberLabel.Returns("sysf");

            pattern = processor.GetFullRegexPattern();
            Assert.AreEqual("<<\\s*(prop|data|meth|sysv|sysf)&&.+?\\s*@@", pattern);

            matches = Regex.Matches("<<prop&&Name@@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<prop&&Name@@", matches[0].Value);

            matches = Regex.Matches("<<prop&&Employee.Contacts.Phone@@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<prop&&Employee.Contacts.Phone@@", matches[0].Value);

            matches = Regex.Matches("<<data&&Name@@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<data&&Name@@", matches[0].Value);

            matches = Regex.Matches("<<  data&&Employee.Contacts.Phone  @@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<  data&&Employee.Contacts.Phone  @@", matches[0].Value);

            matches = Regex.Matches("<<meth&&Meth()@@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<meth&&Meth()@@", matches[0].Value);

            matches = Regex.Matches("<<sysv&&Now@@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<sysv&&Now@@", matches[0].Value);

            matches = Regex.Matches("<<sysf&&Format(prop&&Prop, 0,##)@@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<sysf&&Format(prop&&Prop, 0,##)@@", matches[0].Value);

            matches = Regex.Matches("<<meth&&Meth(prop&&Value, data&&Name, 5, \"Hi\", meth&&Meth2(meth&&Meth3()))@@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<meth&&Meth(prop&&Value, data&&Name, 5, \"Hi\", meth&&Meth2(meth&&Meth3()))@@", matches[0].Value);

            matches = Regex.Matches("One <<prop&&Name@@ <<Two@@ Three <<data&&Employee.Contacts.Phone@@ <<n&&Four@@ <<meth&&Meth1(hi, 5, prop&&Value, meth&&Meth2(meth&&Meth3(data:Complex.Field)))@@ <<ms&&Method()@@ Five <<sysv&&Now@@", pattern);
            Assert.AreEqual(4, matches.Count);
            Assert.AreEqual("<<prop&&Name@@", matches[0].Value);
            Assert.AreEqual("<<data&&Employee.Contacts.Phone@@", matches[1].Value);
            Assert.AreEqual("<<meth&&Meth1(hi, 5, prop&&Value, meth&&Meth2(meth&&Meth3(data:Complex.Field)))@@", matches[2].Value);
            Assert.AreEqual("<<sysv&&Now@@", matches[3].Value);

            matches = Regex.Matches("<<<<prop&&Name@@@@", pattern);
            Assert.AreEqual(1, matches.Count);
        }

        [TestMethod]
        public void TestGetFullRegexPatternWithNullProps()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();

            // Standard case
            processor.LeftTemplateBorder.Returns((string)null);
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.PropertyMemberLabel.Returns("p");
            processor.DataItemMemberLabel.Returns("di");
            processor.MethodCallMemberLabel.Returns("m");
            processor.SystemVariableMemberLabel.Returns("sv");
            processor.SystemFunctionMemberLabel.Returns("sf");

            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.LeftTemplateBorder.Returns(string.Empty);
            Assert.AreEqual("\\s*(p|di|m|sv|sf):.+?\\s*}", processor.GetFullRegexPattern());
            processor.LeftTemplateBorder.Returns(" ");
            Assert.AreEqual("\\ \\s*(p|di|m|sv|sf):.+?\\s*}", processor.GetFullRegexPattern());

            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.RightTemplateBorder.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p|di|m|sv|sf):.+?\\s*", processor.GetFullRegexPattern());
            processor.RightTemplateBorder.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|di|m|sv|sf):.+?\\s*\\ ", processor.GetFullRegexPattern());

            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.MemberLabelSeparator.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p|di|m|sv|sf).+?\\s*}", processor.GetFullRegexPattern());
            processor.MemberLabelSeparator.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|di|m|sv|sf)\\ .+?\\s*}", processor.GetFullRegexPattern());

            processor.MemberLabelSeparator.Returns(":");
            processor.PropertyMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.PropertyMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(|di|m|sv|sf):.+?\\s*}", processor.GetFullRegexPattern());
            processor.PropertyMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*(\\ |di|m|sv|sf):.+?\\s*}", processor.GetFullRegexPattern());

            processor.PropertyMemberLabel.Returns("p");
            processor.DataItemMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.DataItemMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p||m|sv|sf):.+?\\s*}", processor.GetFullRegexPattern());
            processor.DataItemMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|\\ |m|sv|sf):.+?\\s*}", processor.GetFullRegexPattern());

            processor.DataItemMemberLabel.Returns("di");
            processor.MethodCallMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.MethodCallMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p|di||sv|sf):.+?\\s*}", processor.GetFullRegexPattern());
            processor.MethodCallMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|di|\\ |sv|sf):.+?\\s*}", processor.GetFullRegexPattern());

            processor.MethodCallMemberLabel.Returns("m");
            processor.SystemVariableMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.SystemVariableMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p|di|m||sf):.+?\\s*}", processor.GetFullRegexPattern());
            processor.SystemVariableMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|di|m|\\ |sf):.+?\\s*}", processor.GetFullRegexPattern());

            processor.SystemVariableMemberLabel.Returns("sv");
            processor.SystemFunctionMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.SystemFunctionMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p|di|m|sv|):.+?\\s*}", processor.GetFullRegexPattern());
            processor.SystemFunctionMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|di|m|sv|\\ ):.+?\\s*}", processor.GetFullRegexPattern());
        }

        [TestMethod]
        public void TestGetPropertyRegexPattern()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.PropertyMemberLabel.Returns("p");

            string pattern = processor.GetPropertyRegexPattern();
            Assert.AreEqual("\\{\\s*p:.+?\\s*}", pattern);

            MatchCollection matches = Regex.Matches("{p:Name}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{p:Name}", matches[0].Value);

            matches = Regex.Matches("{ p:Contact.Phone }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ p:Contact.Phone }", matches[0].Value);

            matches = Regex.Matches("{pp:Name}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.PropertyMemberLabel.Returns("*");
            pattern = processor.GetPropertyRegexPattern();
            Assert.AreEqual("\\{\\s*\\*:.+?\\s*}", pattern);

            matches = Regex.Matches("{*:Name}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{*:Name}", matches[0].Value);

            matches = Regex.Matches("{*:Contact.Phone}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{*:Contact.Phone}", matches[0].Value);

            matches = Regex.Matches("{**:Name}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.PropertyMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetPropertyRegexPattern());
            processor.PropertyMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*:.+?\\s*}", processor.GetPropertyRegexPattern());
            processor.PropertyMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*\\ :.+?\\s*}", processor.GetPropertyRegexPattern());
        }

        [TestMethod]
        public void TestGetDataItemRegexPattern()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.DataItemMemberLabel.Returns("di");

            string pattern = processor.GetDataItemRegexPattern();
            Assert.AreEqual("\\{\\s*di:.+?\\s*}", pattern);

            MatchCollection matches = Regex.Matches("{ di:Name }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ di:Name }", matches[0].Value);

            matches = Regex.Matches("{di:Contact.Phone}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{di:Contact.Phone}", matches[0].Value);

            matches = Regex.Matches("{d:Name}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.DataItemMemberLabel.Returns("*");
            pattern = processor.GetDataItemRegexPattern();
            Assert.AreEqual("\\{\\s*\\*:.+?\\s*}", pattern);

            matches = Regex.Matches("{ *:Name}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ *:Name}", matches[0].Value);

            matches = Regex.Matches("{*:Contact.Phone}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{*:Contact.Phone}", matches[0].Value);

            matches = Regex.Matches("{**:Name}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.DataItemMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetDataItemRegexPattern());
            processor.DataItemMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*:.+?\\s*}", processor.GetDataItemRegexPattern());
            processor.DataItemMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*\\ :.+?\\s*}", processor.GetDataItemRegexPattern());
        }

        [TestMethod]
        public void TestGetMethodCallRegexPattern()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.MethodCallMemberLabel.Returns("m");

            string pattern = processor.GetMethodCallRegexPattern();
            Assert.AreEqual("\\{\\s*m:.+?\\s*}", pattern);

            MatchCollection matches = Regex.Matches("{m:Meth()}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{m:Meth()}", matches[0].Value);

            matches = Regex.Matches("{ m:Meth(p:Name) }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ m:Meth(p:Name) }", matches[0].Value);

            matches = Regex.Matches("{ms:Meth()}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.MethodCallMemberLabel.Returns("*");
            pattern = processor.GetMethodCallRegexPattern();
            Assert.AreEqual("\\{\\s*\\*:.+?\\s*}", pattern);

            matches = Regex.Matches("{*:Meth()}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{*:Meth()}", matches[0].Value);

            matches = Regex.Matches("{ *:Meth(p:Name)}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ *:Meth(p:Name)}", matches[0].Value);

            matches = Regex.Matches("{**:Meth()}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.MethodCallMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetMethodCallRegexPattern());
            processor.MethodCallMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*:.+?\\s*}", processor.GetMethodCallRegexPattern());
            processor.MethodCallMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*\\ :.+?\\s*}", processor.GetMethodCallRegexPattern());
        }

        [TestMethod]
        public void TestGetVariableRegexPattern()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.SystemVariableMemberLabel.Returns("sv");

            string pattern = processor.GetVariableRegexPattern();
            Assert.AreEqual("\\{\\s*sv:.+?\\s*}", pattern);

            MatchCollection matches = Regex.Matches("{ sv:Now }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ sv:Now }", matches[0].Value);

            matches = Regex.Matches("{d:Now}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.SystemVariableMemberLabel.Returns("*");
            pattern = processor.GetVariableRegexPattern();
            Assert.AreEqual("\\{\\s*\\*:.+?\\s*}", pattern);

            matches = Regex.Matches("{ *:Now}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ *:Now}", matches[0].Value);

            matches = Regex.Matches("{**:Now}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.SystemVariableMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetVariableRegexPattern());
            processor.SystemVariableMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*:.+?\\s*}", processor.GetVariableRegexPattern());
            processor.SystemVariableMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*\\ :.+?\\s*}", processor.GetVariableRegexPattern());
        }

        [TestMethod]
        public void TestGetSystemFunctionRegexPattern()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.SystemFunctionMemberLabel.Returns("sf");

            string pattern = processor.GetSystemFunctionRegexPattern();
            Assert.AreEqual("\\{\\s*sf:.+?\\s*}", pattern);

            MatchCollection matches = Regex.Matches("{ sf:Format(p:Prop, 0) }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ sf:Format(p:Prop, 0) }", matches[0].Value);

            matches = Regex.Matches("{d:Now}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.SystemFunctionMemberLabel.Returns("*");
            pattern = processor.GetSystemFunctionRegexPattern();
            Assert.AreEqual("\\{\\s*\\*:.+?\\s*}", pattern);

            matches = Regex.Matches("{ *:Format(p:Prop, 0)}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ *:Format(p:Prop, 0)}", matches[0].Value);

            matches = Regex.Matches("{**:Format(p:Prop, 0)}", pattern);
            Assert.AreEqual(0, matches.Count);

            processor.SystemFunctionMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetSystemFunctionRegexPattern());
            processor.SystemFunctionMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*:.+?\\s*}", processor.GetSystemFunctionRegexPattern());
            processor.SystemFunctionMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*\\ :.+?\\s*}", processor.GetSystemFunctionRegexPattern());
        }

        [TestMethod]
        public void TestGetAggregationRegexPatterns()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();

            // Standard case
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.DataItemMemberLabel.Returns("di");
            processor.MemberLabelSeparator.Returns(":");

            string[] allAggFuncs = Enum.GetNames(typeof(AggregateFunction)).Where(n => n != AggregateFunction.NoAggregation.ToString()).ToArray();
            Assert.AreEqual(6, allAggFuncs.Length);

            string templatesWithAggregationPattern = processor.GetTemplatesWithAggregationRegexPattern();
            Assert.AreEqual($@"\{{[^}}]*({string.Join("|", allAggFuncs)})\(\s*di\s*:.+?\)[^}}]*}}", templatesWithAggregationPattern);

            string aggregationFuncPattern = processor.GetAggregationFuncRegexPattern();
            Assert.AreEqual($@"({string.Join("|", allAggFuncs)})\((\s*di\s*:.+?)\)", aggregationFuncPattern);

            MatchCollection matches = Regex.Matches("{Sum(di:Amount)}", templatesWithAggregationPattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{Sum(di:Amount)}", matches[0].Value);

            MatchCollection innerMatches = Regex.Matches(matches[0].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Sum(di:Amount)", innerMatches[0].Value);
            Assert.AreEqual("Sum", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Amount", innerMatches[0].Groups[2].Value);

            matches = Regex.Matches("{ Count( di : Amount, CustomAggregation, PostAggregation) }", templatesWithAggregationPattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ Count( di : Amount, CustomAggregation, PostAggregation) }", matches[0].Value);

            innerMatches = Regex.Matches(matches[0].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Count( di : Amount, CustomAggregation, PostAggregation)", innerMatches[0].Value);
            Assert.AreEqual("Count", innerMatches[0].Groups[1].Value);
            Assert.AreEqual(" di : Amount, CustomAggregation, PostAggregation", innerMatches[0].Groups[2].Value);

            matches = Regex.Matches("{  Max(di:Amount,,PostAggregation)  }", templatesWithAggregationPattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{  Max(di:Amount,,PostAggregation)  }", matches[0].Value);

            innerMatches = Regex.Matches(matches[0].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Max(di:Amount,,PostAggregation)", innerMatches[0].Value);
            Assert.AreEqual("Max", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Amount,,PostAggregation", innerMatches[0].Groups[2].Value);

            matches = Regex.Matches("{ Min(di:Result.Amount) }", templatesWithAggregationPattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ Min(di:Result.Amount) }", matches[0].Value);

            innerMatches = Regex.Matches(matches[0].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Min(di:Result.Amount)", innerMatches[0].Value);
            Assert.AreEqual("Min", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Result.Amount", innerMatches[0].Groups[2].Value);

            matches = Regex.Matches("{Avg(di:Amount,  CustomFunc)}", templatesWithAggregationPattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{Avg(di:Amount,  CustomFunc)}", matches[0].Value);

            innerMatches = Regex.Matches(matches[0].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Avg(di:Amount,  CustomFunc)", innerMatches[0].Value);
            Assert.AreEqual("Avg", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Amount,  CustomFunc", innerMatches[0].Groups[2].Value);

            matches = Regex.Matches("{Custom(di:Amount)}", templatesWithAggregationPattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{Custom(di:Amount)}", matches[0].Value);

            innerMatches = Regex.Matches(matches[0].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Custom(di:Amount)", innerMatches[0].Value);
            Assert.AreEqual("Custom", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Amount", innerMatches[0].Groups[2].Value);

            matches = Regex.Matches("Text {Plain Text} {Sum(di:Amount))} {p:Text} {Avg(di:Value)} {sv:RenderDate}", templatesWithAggregationPattern);
            Assert.AreEqual(2, matches.Count);
            Assert.AreEqual("{Sum(di:Amount))}", matches[0].Value);
            Assert.AreEqual("{Avg(di:Value)}", matches[1].Value);

            innerMatches = Regex.Matches(matches[0].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Sum(di:Amount)", innerMatches[0].Value);
            Assert.AreEqual("Sum", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Amount", innerMatches[0].Groups[2].Value);

            innerMatches = Regex.Matches(matches[1].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Avg(di:Value)", innerMatches[0].Value);
            Assert.AreEqual("Avg", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Value", innerMatches[0].Groups[2].Value);

            matches = Regex.Matches("Text {Plain Text} Sum(di:Count) {sf:Format(Sum(di:Amount,,PostAggregation), #,,0.00)} {p:Text} {Max(di:Count)} {m:Meth(1, Avg( di : Value ), Min(di:Amount, CustomAggregation, PostAggregation), \"Str\"} {sv:RenderDate} m:Meth2(Avg(di:Value))", templatesWithAggregationPattern);
            Assert.AreEqual(3, matches.Count);
            Assert.AreEqual("{sf:Format(Sum(di:Amount,,PostAggregation), #,,0.00)}", matches[0].Value);
            Assert.AreEqual("{Max(di:Count)}", matches[1].Value);
            Assert.AreEqual("{m:Meth(1, Avg( di : Value ), Min(di:Amount, CustomAggregation, PostAggregation), \"Str\"}", matches[2].Value);

            innerMatches = Regex.Matches(matches[0].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Sum(di:Amount,,PostAggregation)", innerMatches[0].Value);
            Assert.AreEqual("Sum", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Amount,,PostAggregation", innerMatches[0].Groups[2].Value);

            innerMatches = Regex.Matches(matches[1].Value, aggregationFuncPattern);
            Assert.AreEqual(1, innerMatches.Count);
            Assert.AreEqual("Max(di:Count)", innerMatches[0].Value);
            Assert.AreEqual("Max", innerMatches[0].Groups[1].Value);
            Assert.AreEqual("di:Count", innerMatches[0].Groups[2].Value);

            innerMatches = Regex.Matches(matches[2].Value, aggregationFuncPattern);
            Assert.AreEqual(2, innerMatches.Count);
            Assert.AreEqual("Avg( di : Value )", innerMatches[0].Value);
            Assert.AreEqual("Avg", innerMatches[0].Groups[1].Value);
            Assert.AreEqual(" di : Value ", innerMatches[0].Groups[2].Value);
            Assert.AreEqual("Min(di:Amount, CustomAggregation, PostAggregation)", innerMatches[1].Value);
            Assert.AreEqual("Min", innerMatches[1].Groups[1].Value);
            Assert.AreEqual("di:Amount, CustomAggregation, PostAggregation", innerMatches[1].Groups[2].Value);

            matches = Regex.Matches("{Mix(di:Amount)}", templatesWithAggregationPattern);
            Assert.AreEqual(0, matches.Count);

            matches = Regex.Matches("{Max(Amount)}", templatesWithAggregationPattern);
            Assert.AreEqual(0, matches.Count);

            matches = Regex.Matches("{Max()}", templatesWithAggregationPattern);
            Assert.AreEqual(0, matches.Count);

            // Overriden borders
            processor.LeftTemplateBorder.Returns("<");
            processor.RightTemplateBorder.Returns(">");
            processor.DataItemMemberLabel.Returns("d");
            processor.MemberLabelSeparator.Returns("-");

            templatesWithAggregationPattern = processor.GetTemplatesWithAggregationRegexPattern();
            Assert.AreEqual($@"<[^>]*({string.Join("|", allAggFuncs)})\(\s*d\s*-.+?\)[^>]*>", templatesWithAggregationPattern);

            aggregationFuncPattern = processor.GetAggregationFuncRegexPattern();
            Assert.AreEqual($@"({string.Join("|", allAggFuncs)})\((\s*d\s*-.+?)\)", aggregationFuncPattern);
        }

        [TestMethod]
        public void TestBuildAggregationFuncTemplate()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns(":");
            processor.DataItemMemberLabel.Returns("di");

            Assert.AreEqual("{Sum(di:Amount)}", processor.BuildAggregationFuncTemplate(AggregateFunction.Sum, "Amount"));
            Assert.AreEqual("{Count(di:Value)}", processor.BuildAggregationFuncTemplate(AggregateFunction.Count, "Value"));
            Assert.AreEqual("{Avg(di:Result.Sum)}", processor.BuildAggregationFuncTemplate(AggregateFunction.Avg, "Result.Sum"));
            Assert.AreEqual("{Min(di:)}", processor.BuildAggregationFuncTemplate(AggregateFunction.Min, null));
            Assert.AreEqual("{Max(di:)}", processor.BuildAggregationFuncTemplate(AggregateFunction.Max, string.Empty));
        }

        [TestMethod]
        public void TestIsHorizontalPageBreak()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.HorizontalPageBreakLabel.Returns("Horiz");

            Assert.IsTrue(processor.IsHorizontalPageBreak("{Horiz}"));
            Assert.IsTrue(processor.IsHorizontalPageBreak("{HORIZ}"));
            Assert.IsTrue(processor.IsHorizontalPageBreak("{  Horiz  }"));
            Assert.IsTrue(processor.IsHorizontalPageBreak(" {Horiz} "));
            Assert.IsTrue(processor.IsHorizontalPageBreak("  {  horiz  }  "));

            Assert.IsFalse(processor.IsHorizontalPageBreak("Text {Horiz}"));
            Assert.IsFalse(processor.IsHorizontalPageBreak("{Horiz} Text"));
            Assert.IsFalse(processor.IsHorizontalPageBreak("Horiz"));
            Assert.IsFalse(processor.IsHorizontalPageBreak("{Vert}"));
            Assert.IsFalse(processor.IsHorizontalPageBreak("Bad"));
            Assert.IsFalse(processor.IsHorizontalPageBreak(string.Empty));
            Assert.IsFalse(processor.IsHorizontalPageBreak(null));
        }

        [TestMethod]
        public void TestIsVerticalPageBreak()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");
            processor.VerticalPageBreakLabel.Returns("Vert");

            Assert.IsTrue(processor.IsVerticalPageBreak("{Vert}"));
            Assert.IsTrue(processor.IsVerticalPageBreak(" {  vert  } "));
            Assert.IsTrue(processor.IsVerticalPageBreak("{ VERT }"));
            Assert.IsTrue(processor.IsVerticalPageBreak(" {vErT} "));

            Assert.IsFalse(processor.IsVerticalPageBreak("Text {Vert}"));
            Assert.IsFalse(processor.IsVerticalPageBreak("{Vert} Text"));
            Assert.IsFalse(processor.IsVerticalPageBreak("Vert"));
            Assert.IsFalse(processor.IsVerticalPageBreak("{Horiz}"));
            Assert.IsFalse(processor.IsVerticalPageBreak("Bad"));
            Assert.IsFalse(processor.IsVerticalPageBreak(string.Empty));
            Assert.IsFalse(processor.IsHorizontalPageBreak(null));
        }
    }
}