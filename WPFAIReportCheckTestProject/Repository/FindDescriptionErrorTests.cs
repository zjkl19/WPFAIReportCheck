using Xunit;
using WPFAIReportCheck.Repository;
using WPFAIReportCheck.Models;
using Aspose.Words;
using System.IO;
using log4net;
using Moq;
using System;

namespace WPFAIReportCheckTestProject.Repository
{
    public partial class AsposeAIReportCheckTests
    {

        [Fact]
        public void FindDescriptionError_ReturnsCorrectCountOfReportError()
        {
            //Arrange
            var log = new Mock<ILog>();
            log.Setup(m => m.Error(It.IsAny<string>(), It.IsAny<Exception>()));   //无实际意义，仅作为1必须参数传入

            var fileName = @"..\..\..\TestFiles\FindDescriptionError.doc";
            var ai = new AsposeAIReportCheck(fileName, log.Object);
            //Act
            ai.FindDescriptionError();
            //Assert
            Assert.Single(ai.reportError);
            Assert.Equal(ErrorNumber.Description, ai.reportError[0].No);
        }

        [Fact]
        public void FindDescriptionError_WritesCorrectCommentsInDoc()
        {
            //Arrange
            var log = new Mock<ILog>();
            log.Setup(m => m.Error(It.IsAny<string>(), It.IsAny<Exception>()));   //无实际意义，仅作为1必须参数传入

            var fileName = @"..\..\..\TestFiles\FindDescriptionError.doc";

            var ai = new AsposeAIReportCheck(fileName,log.Object);
            //Act
            ai.FindDescriptionError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Single(allComments);
            Assert.True(docComment.GetText().IndexOf("应为\"CNAS检验机构\"") >= 0);
        }
    }
}
