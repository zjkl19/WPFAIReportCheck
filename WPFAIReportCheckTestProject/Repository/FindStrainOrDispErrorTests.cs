using Xunit;
using WPFAIReportCheck.Repository;
using WPFAIReportCheck.Models;
using Aspose.Words;
using System.IO;

namespace WPFAIReportCheckTestProject.Repository
{
    public partial class AsposeAIReportCheckTests
    {
        [Fact]
        public void FindStrainOrDispError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\FindStrainOrDispError.doc";
            var ai = new AsposeAIReportCheck(fileName);
            //Act
            ai.FindStrainOrDispError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(3,allComments.Count);    //3处错误
            Assert.True(docComment.GetText().IndexOf("计算错误") >= 0);
        }
    }
}
