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
        public void FindStrainOrDispError_ReturnsCorrectCountOfReportError()
        {
            //Arrange
            var log = new Mock<ILog>();
            log.Setup(m => m.Error(It.IsAny<string>(), It.IsAny<Exception>()));   //无实际意义，仅作为1必须参数传入

            var fileName = @"..\..\..\TestFiles\FindStrainOrDispError.doc";
            var ai = new AsposeAIReportCheck(fileName, log.Object);
            //Act
            ai.FindStrainOrDispError();
            //Assert
            Assert.Equal(6, ai.reportError.Count);
            Assert.Equal(ErrorNumber.Calc, ai.reportError[0].No);
        }

        [Fact]
        public void FindStrainOrDispError_WritesCorrectCommentsInDoc()
        {
            //Arrange
            var log = new Mock<ILog>();
            log.Setup(m => m.Error(It.IsAny<string>(), It.IsAny<Exception>()));   //无实际意义，仅作为1必须参数传入

            var fileName = @"..\..\..\TestFiles\FindStrainOrDispError.doc";

            var ai = new AsposeAIReportCheck(fileName,log.Object);
            //Act
            ai.FindStrainOrDispError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(6,allComments.Count);
            Assert.True(docComment.GetText().IndexOf("计算错误") >= 0);
        }
    }
}
