using Xunit;
using WPFAIReportCheck.Repository;
using WPFAIReportCheck.Models;
//using Aspose.Words;
using System.IO;
using Moq;
using System;
using System.Xml.Linq;
using Aspose.Words;

namespace WPFAIReportCheckTestProject.Repository
{
    public partial class AsposeAIReportCheckTests : IClassFixture<AsposeAIReportCheckTestsFixture>
    {

        #region FindSequenceNumberError
        [Fact]
        public void FindSequenceNumberError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindSequenceNumberError.doc";
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
            ai.FindSequenceNumberError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Single(allComments);
            Assert.True(docComment.GetText().IndexOf("序号应连贯，从小到大") >= 0);
        }
        #endregion
    }
}
