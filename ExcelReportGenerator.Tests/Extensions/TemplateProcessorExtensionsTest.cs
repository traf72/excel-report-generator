﻿using ExcelReportGenerator.Enums;
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
        public void TestTrimPropertyLabel()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.MemberLabelSeparator.Returns(":");
            processor.PropertyMemberLabel.Returns("p");

            Assert.AreEqual("Name", processor.TrimPropertyLabel("p:Name"));
            Assert.AreEqual("Name p:Phone", processor.TrimPropertyLabel("p:Name p:Phone"));
            Assert.AreEqual("di:Name", processor.TrimPropertyLabel("di:Name"));
            Assert.AreEqual("Contacts.Phone", processor.TrimPropertyLabel("Contacts.Phone"));
            Assert.AreEqual(" { Name } ", processor.TrimPropertyLabel(" { p:Name } "));
            Assert.AreEqual("{Name p:Phone}", processor.TrimPropertyLabel("{p:Name p:Phone}"));
            Assert.AreEqual(" { di:Name } ", processor.TrimPropertyLabel(" { di:Name } "));
            Assert.AreEqual("{Contacts.Phone}", processor.TrimPropertyLabel("{Contacts.Phone}"));
            ExceptionAssert.Throws<ArgumentNullException>(() => processor.TrimPropertyLabel(null));

            processor.MemberLabelSeparator.Returns((string)null);
            Assert.AreEqual("Name", processor.TrimPropertyLabel("pName"));
            Assert.AreEqual("{Name}", processor.TrimPropertyLabel("{pName}"));

            processor.PropertyMemberLabel.Returns((string)null);
            Assert.AreEqual("p:Name", processor.TrimPropertyLabel("p:Name"));
        }

        [TestMethod]
        public void TestTrimDataItemLabel()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();
            processor.MemberLabelSeparator.Returns(":");
            processor.DataItemMemberLabel.Returns("di");

            Assert.AreEqual("Name", processor.TrimDataItemLabel("di:Name"));
            Assert.AreEqual("Name di:Phone", processor.TrimDataItemLabel("di:Name di:Phone"));
            Assert.AreEqual("p:Name", processor.TrimDataItemLabel("p:Name"));
            Assert.AreEqual("Contacts.Phone", processor.TrimDataItemLabel("Contacts.Phone"));
            Assert.AreEqual(" { Name } ", processor.TrimDataItemLabel(" { di:Name } "));
            Assert.AreEqual("{Name di:Phone}", processor.TrimDataItemLabel("{di:Name di:Phone}"));
            Assert.AreEqual(" { p:Name } ", processor.TrimDataItemLabel(" { p:Name } "));
            Assert.AreEqual("{Contacts.Phone}", processor.TrimDataItemLabel("{Contacts.Phone}"));
            ExceptionAssert.Throws<ArgumentNullException>(() => processor.TrimDataItemLabel(null));

            processor.MemberLabelSeparator.Returns((string)null);
            Assert.AreEqual("Name", processor.TrimDataItemLabel("diName"));
            Assert.AreEqual("{Name}", processor.TrimDataItemLabel("{diName}"));

            processor.DataItemMemberLabel.Returns((string)null);
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
            Assert.AreEqual("{Meth1(p:Name, di:Value, m:MethInner()) m:Meth2()}", processor.TrimMethodCallLabel("{m:Meth1(p:Name, di:Value, m:MethInner()) m:Meth2()}"));
            Assert.AreEqual(" { p:Name } ", processor.TrimMethodCallLabel(" { p:Name } "));
            ExceptionAssert.Throws<ArgumentNullException>(() => processor.TrimMethodCallLabel(null));

            processor.MemberLabelSeparator.Returns((string)null);
            Assert.AreEqual("Meth()", processor.TrimMethodCallLabel("mMeth()"));
            Assert.AreEqual("{Meth()}", processor.TrimMethodCallLabel("{mMeth()}"));

            processor.MethodCallMemberLabel.Returns((string)null);
            Assert.AreEqual("m:Meth()", processor.TrimMethodCallLabel("m:Meth()"));
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

            string pattern = processor.GetFullRegexPattern();
            Assert.AreEqual("\\{\\s*(p|di|m):.+?\\s*}", pattern);

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

            pattern = processor.GetFullRegexPattern();
            Assert.AreEqual("\\[\\s*(prop|data|meth)\\*.+?\\s*]", pattern);

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

            matches = Regex.Matches("[[prop*Name]]", pattern);
            Assert.AreEqual(1, matches.Count);

            // Overriden props case 2
            processor.LeftTemplateBorder.Returns("<<");
            processor.RightTemplateBorder.Returns("@@");
            processor.MemberLabelSeparator.Returns("&&");
            processor.PropertyMemberLabel.Returns("prop");
            processor.DataItemMemberLabel.Returns("data");
            processor.MethodCallMemberLabel.Returns("meth");

            pattern = processor.GetFullRegexPattern();
            Assert.AreEqual("<<\\s*(prop|data|meth)&&.+?\\s*@@", pattern);

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

            matches = Regex.Matches("<<meth&&Meth(prop&&Value, data&&Name, 5, \"Hi\", meth&&Meth2(meth&&Meth3()))@@", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("<<meth&&Meth(prop&&Value, data&&Name, 5, \"Hi\", meth&&Meth2(meth&&Meth3()))@@", matches[0].Value);

            matches = Regex.Matches("One <<prop&&Name@@ <<Two@@ Three <<data&&Employee.Contacts.Phone@@ <<n&&Four@@ <<meth&&Meth1(hi, 5, prop&&Value, meth&&Meth2(meth&&Meth3(data:Coplex.Field)))@@ <<ms&&Method()@@", pattern);
            Assert.AreEqual(3, matches.Count);
            Assert.AreEqual("<<prop&&Name@@", matches[0].Value);
            Assert.AreEqual("<<data&&Employee.Contacts.Phone@@", matches[1].Value);
            Assert.AreEqual("<<meth&&Meth1(hi, 5, prop&&Value, meth&&Meth2(meth&&Meth3(data:Coplex.Field)))@@", matches[2].Value);

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

            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.LeftTemplateBorder.Returns(string.Empty);
            Assert.AreEqual("\\s*(p|di|m):.+?\\s*}", processor.GetFullRegexPattern());
            processor.LeftTemplateBorder.Returns(" ");
            Assert.AreEqual("\\ \\s*(p|di|m):.+?\\s*}", processor.GetFullRegexPattern());

            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.RightTemplateBorder.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p|di|m):.+?\\s*", processor.GetFullRegexPattern());
            processor.RightTemplateBorder.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|di|m):.+?\\s*\\ ", processor.GetFullRegexPattern());

            processor.RightTemplateBorder.Returns("}");
            processor.MemberLabelSeparator.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.MemberLabelSeparator.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p|di|m).+?\\s*}", processor.GetFullRegexPattern());
            processor.MemberLabelSeparator.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|di|m)\\ .+?\\s*}", processor.GetFullRegexPattern());

            processor.MemberLabelSeparator.Returns(":");
            processor.PropertyMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.PropertyMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(|di|m):.+?\\s*}", processor.GetFullRegexPattern());
            processor.PropertyMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*(\\ |di|m):.+?\\s*}", processor.GetFullRegexPattern());

            processor.PropertyMemberLabel.Returns("p");
            processor.DataItemMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.DataItemMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p||m):.+?\\s*}", processor.GetFullRegexPattern());
            processor.DataItemMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|\\ |m):.+?\\s*}", processor.GetFullRegexPattern());

            processor.DataItemMemberLabel.Returns("di");
            processor.MethodCallMemberLabel.Returns((string)null);
            ExceptionAssert.Throws<Exception>(() => processor.GetFullRegexPattern());
            processor.MethodCallMemberLabel.Returns(string.Empty);
            Assert.AreEqual("\\{\\s*(p|di|):.+?\\s*}", processor.GetFullRegexPattern());
            processor.MethodCallMemberLabel.Returns(" ");
            Assert.AreEqual("\\{\\s*(p|di|\\ ):.+?\\s*}", processor.GetFullRegexPattern());
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
        public void TestGetFullAggregationRegexPattern()
        {
            ITemplateProcessor processor = Substitute.For<ITemplateProcessor>();

            // Standard case
            processor.LeftTemplateBorder.Returns("{");
            processor.RightTemplateBorder.Returns("}");

            string pattern = processor.GetFullAggregationRegexPattern();
            string[] allAggFuncs = Enum.GetNames(typeof(AggregateFunction)).Where(n => n != AggregateFunction.NoAggregation.ToString()).ToArray();

            Assert.AreEqual(6, allAggFuncs.Length);
            Assert.AreEqual($"\\{{\\s*({string.Join("|", allAggFuncs)})\\((.+?)\\)\\s*}}", pattern);

            MatchCollection matches = Regex.Matches("{Sum(di:Amount)}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{Sum(di:Amount)}", matches[0].Value);

            matches = Regex.Matches("{ Count(di:Amount) }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ Count(di:Amount) }", matches[0].Value);

            matches = Regex.Matches("{  Max(Amount)  }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{  Max(Amount)  }", matches[0].Value);

            matches = Regex.Matches("{ Min(Result.Amount) }", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{ Min(Result.Amount) }", matches[0].Value);

            matches = Regex.Matches("{Avg(di:Amount)}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{Avg(di:Amount)}", matches[0].Value);

            matches = Regex.Matches("{Custom(di:Amount)}", pattern);
            Assert.AreEqual(1, matches.Count);
            Assert.AreEqual("{Custom(di:Amount)}", matches[0].Value);

            matches = Regex.Matches("Text {Sum(di:Amount))} {Text} {Avg(Value)}", pattern);
            Assert.AreEqual(2, matches.Count);
            Assert.AreEqual("{Sum(di:Amount))}", matches[0].Value);
            Assert.AreEqual("{Avg(Value)}", matches[1].Value);

            matches = Regex.Matches("{Mix(di:Amount)}", pattern);
            Assert.AreEqual(0, matches.Count);

            // Overriden borders
            processor.LeftTemplateBorder.Returns("<");
            processor.RightTemplateBorder.Returns(">");

            pattern = processor.GetFullAggregationRegexPattern();
            Assert.AreEqual($"<\\s*({string.Join("|", allAggFuncs)})\\((.+?)\\)\\s*>", pattern);
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
    }
}