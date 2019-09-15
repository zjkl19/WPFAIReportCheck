using Xunit;
using WPFAIReportCheck.Repository;
using WPFAIReportCheck.Models;
using Aspose.Words;
using System.IO;
using Moq;
using System;
using System.Xml.Linq;
using NLog;

namespace WPFAIReportCheckTestProject.Repository
{
    public partial class AsposeAIReportCheckTests : IClassFixture<AsposeAIReportCheckTestsFixture>
    {
        [Fact]
        public void FindHalfWidthCommaWarnning_ReturnsCorrectCountOfReportError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindHalfWidthCommaWarnning.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);

            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindHalfWidthCommaWarnning();
            //Assert
            Assert.Equal(2,ai.reportWarnning.Count);
            Assert.Equal(WarnningNumber.FormatProblem, ai.reportWarnning[0].No);
            Assert.Equal(WarnningNumber.FormatProblem, ai.reportWarnning[1].No);
        }

        [Fact]
        public void FindHalfWidthCommaWarnning_WritesCorrectCommentsInDoc()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindHalfWidthCommaWarnning.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);
            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindHalfWidthCommaWarnning();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment1 = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            Comment docComment2 = (Comment)ai._doc.GetChild(NodeType.Comment, 1, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(2, allComments.Count);
            Assert.Contains("请确认此处是否使用半角逗号", docComment1.GetText());
            Assert.Contains("请确认此处是否使用半角逗号", docComment2.GetText());
        }
    }
}
