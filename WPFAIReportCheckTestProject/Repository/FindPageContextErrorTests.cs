using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using WPFAIReportCheck.Models;
using WPFAIReportCheck.Repository;
using Xunit;

namespace WPFAIReportCheckTestProject.Repository
{

    public partial class AsposeAIReportCheckTests : IClassFixture<AsposeAIReportCheckTestsFixture>
    {

        [Fact]
        public void FindPageContextError_ReturnsCorrectCountOfReportError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindPageContextError.doc";

            var log = _fixture.log;
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
            ai.FindPageContextError();
            //Assert
            Assert.Equal(2, ai.reportError.Count);
            Assert.Equal(ErrorNumber.Description, ai.reportError[0].No);
            Assert.Equal(ErrorNumber.Description, ai.reportError[1].No);
        }

        [Fact]
        public void FindPageContextError_WritesCorrectCommentsInDoc()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindPageContextError.doc";

            var log = _fixture.log;
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
            ai.FindPageContextError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment1 = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            Comment docComment2 = (Comment)ai._doc.GetChild(NodeType.Comment, 1, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(2, allComments.Count);
            Assert.Contains("正文第9页未找到2.3 下部结构检查结果", docComment1.GetText());
            Assert.Contains("正文只有18页，但页码索引却有19", docComment2.GetText());
        }
    }
}
