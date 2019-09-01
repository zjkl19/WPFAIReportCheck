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
        /// <summary>
        /// 测试静力荷载试验概述表格及汇总表格附近统计数据错误是否能够识别
        /// </summary>

        [Fact]
        public void FindStrainOrDispContextWarnning_ReturnsCorrectCountOfReportWarnning()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindStrainOrDispContextWarnning.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);

            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindStrainOrDispContextWarnning();
            //Assert
            Assert.Equal(2,ai.reportWarnning.Count);
            Assert.Equal(WarnningNumber.NotFoundInfo, ai.reportWarnning[0].No);
            Assert.Contains("校验系数在0.33～0.36之间",ai.reportWarnning[0].Description);    //概述表格中统计数据错误
            Assert.Equal(WarnningNumber.NotFoundInfo, ai.reportWarnning[1].No);
            Assert.Contains("校验系数在0.35～0.44之间", ai.reportWarnning[1].Description);    //汇总表格附近统计数据错误
        }

        [Fact]
        public void FindStrainOrDispContextWarnning_WritesCorrectCommentsInDoc()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindStrainOrDispContextWarnning.doc";

            var log = _fixture.log;
            string xml = DataInitializer.AIReportCheckXml;
            var config = XDocument.Parse(xml);
            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindStrainOrDispContextWarnning();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment1 = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            Comment docComment2 = (Comment)ai._doc.GetChild(NodeType.Comment, 1, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(2,allComments.Count);
            Assert.Contains("校验系数在0.33～0.36之间", docComment1.GetText());
            Assert.Contains("校验系数在0.35～0.44之间", docComment2.GetText());
        }
    }
}
