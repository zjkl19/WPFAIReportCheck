using Xunit;
using WPFAIReportCheck.Repository;
using WPFAIReportCheck.Models;
using Aspose.Words;
using System.IO;
using log4net;
using Moq;
using System;
using System.Xml.Linq;

namespace WPFAIReportCheckTestProject.Repository
{
    public partial class AsposeAIReportCheckTests
    {

        [Fact]
        public void FindStrainOrDispError_ReturnsCorrectCountOfReportError()
        {
            //Arrange
            var fileName = @"C:\VsProjects\WPFAIReportCheck\WPFAIReportCheckTestProject\TestFiles\FindStrainOrDispError.doc";

            var log = new Mock<ILog>();
            log.Setup(m => m.Error(It.IsAny<string>(), It.IsAny<Exception>()));   //无实际意义，仅作为1必须参数传入

            string xml = @"<?xml version=""1.0"" encoding=""utf - 8"" ?>
                            <configuration>
                              <FindStrainOrDispError row1=""0"" col1=""0"" row2 =""1"" col2 =""1""  charactorString =""测点号"" >
                                <Strain charactorString = ""总应变"" />
                                <Disp charactorString = ""总变形"" />
                             </FindStrainOrDispError >
                           </configuration >";
            var config = XDocument.Parse(xml);

            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
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
            var fileName = @"..\..\..\TestFiles\FindStrainOrDispError.doc";

            var log = new Mock<ILog>();
            log.Setup(m => m.Error(It.IsAny<string>(), It.IsAny<Exception>()));   //无实际意义，仅作为1必须参数传入

            string xml = @"<?xml version=""1.0"" encoding=""utf - 8"" ?>
                            <configuration>
                              <FindStrainOrDispError row1=""0"" col1=""0"" row2 =""1"" col2 =""1""  charactorString =""测点号"" >
                                <Strain charactorString = ""总应变"" />
                                <Disp charactorString = ""总变形"" />
                             </FindStrainOrDispError >
                           </configuration >";
            var config = XDocument.Parse(xml);
            var ai = new AsposeAIReportCheck(fileName, log.Object, config);
            //Act
            ai.FindStrainOrDispError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(6, allComments.Count);
            Assert.True(docComment.GetText().IndexOf("计算错误") >= 0);
        }
    }
}
