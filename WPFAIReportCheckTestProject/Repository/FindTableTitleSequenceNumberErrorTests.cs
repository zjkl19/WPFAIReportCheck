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
        public void FindTableTitleSequenceNumberError_ReturnsCorrectCountOfReportError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindTableTitleSequenceNumberError.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);

            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindTableTitleSequenceNumberError();
            //Assert
            Assert.Equal(2,ai.reportError.Count);
            Assert.Equal(ErrorNumber.Description, ai.reportError[0].No);
        }

        [Fact]
        public void FindTableTitleSequenceNumberError_WritesCorrectCommentsInDoc()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindTableTitleSequenceNumberError.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);
            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindTableTitleSequenceNumberError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment1 = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            Comment docComment2 = (Comment)ai._doc.GetChild(NodeType.Comment, 1, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(2, allComments.Count);
            Assert.Contains("本表格序号应为3-3", docComment1.GetText());
            Assert.Contains("本表格序号应为3-5", docComment2.GetText());
        }
    }
}
