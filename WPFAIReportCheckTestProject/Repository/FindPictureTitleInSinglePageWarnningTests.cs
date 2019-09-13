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
        public void FindPictureTitleInSinglePageWarnning_ReturnsCorrectCountOfReportError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindPictureTitleInSinglePageWarnning.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);

            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindPictureTitleInSinglePageWarnning();
            //Assert
            Assert.Single(ai.reportWarnning);
            Assert.Equal(WarnningNumber.FormatProblem, ai.reportWarnning[0].No);
        }

        [Fact]
        public void FindPictureTitleInSinglePageWarnning_WritesCorrectCommentsInDoc()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindPictureTitleInSinglePageWarnning.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);
            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindPictureTitleInSinglePageWarnning();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Single(allComments);
            Assert.Contains("图片标题单独位于某1页",docComment.GetText());
        }
    }
}
