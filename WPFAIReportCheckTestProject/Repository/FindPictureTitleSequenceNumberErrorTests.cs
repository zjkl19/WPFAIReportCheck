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
        public void FindPictureTitleSequenceNumberError_ReturnsCorrectCountOfReportError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindPictureTitleSequenceNumberError.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);

            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindPictureTitleSequenceNumberError();
            //Assert
            Assert.Equal(2,ai.reportError.Count);
            Assert.Equal(ErrorNumber.Description, ai.reportError[0].No);
        }

        [Fact]
        public void FindPictureTitleSequenceNumberError_WritesCorrectCommentsInDoc()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindPictureTitleSequenceNumberError.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);
            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindPictureTitleSequenceNumberError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(2,allComments.Count);
            Assert.Contains("本图片序号应为3-8", docComment.GetText());
        }
    }
}
