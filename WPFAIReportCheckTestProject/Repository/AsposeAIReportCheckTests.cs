using System;
using System.Collections.Generic;
using System.Text;
using Xunit;
using WPFAIReportCheck.Repository;
using WPFAIReportCheck.Models;
using Aspose.Words;
using System.IO;

namespace WPFAIReportCheckTestProject.Repository
{
    public class AsposeAIReportCheckTests
    {
        #region _FindNotExplainComponentNo
        [Fact]
        public void _FindNotExplainComponentNo_ReturnsAReportWarnning_WhileComponentNoNotExplained()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\_FindNotExplainComponentNo.doc";
            var ai = new AsposeAIReportCheck(fileName);
            //Act
            ai._FindNotExplainComponentNo();
            //Assert
            Assert.Equal(WarnningNumber.NotClearInfo, ai.reportWarnning[0].No);     
        }

        [Fact]
        public void _FindNotExplainComponentNo_ReturnsNoneReportWarnning_WhileComponentNoExplained()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\_FindNotExplainComponentNo_Explained.doc";
            var ai = new AsposeAIReportCheck(fileName);
            //Act
            ai._FindNotExplainComponentNo();
            //Assert
            Assert.Empty(ai.reportWarnning);
        }
        #endregion
        #region _FindUnitError
        [Fact]
        public void _FindUnitError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\_FindUnitError.doc";
            var ai = new AsposeAIReportCheck(fileName);
            //Act
            ai._FindUnitError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            Comment docComment1 = (Comment)ai._doc.GetChild(NodeType.Comment, 1, true);

            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Equal(2, allComments.Count);
            Assert.Equal(1, docComment.Count);
            Assert.True(docComment.GetText().IndexOf("应为km/h") >=0);
            Assert.True(docComment1.GetText().IndexOf("应为km/h") >= 0);
            //Assert.Equal("\u0005My comment.\r", docComment.GetText());
        }
        #endregion

        #region _FindSpecificationsError
        [Fact]
        public void _FindSpecificationsError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\_FindSpecificationsError.doc";
            var ai = new AsposeAIReportCheck(fileName);
            //Act
            ai._FindSpecificationsError();

            using (MemoryStream dstStream = new MemoryStream())
            {
                ai._doc.Save(dstStream, SaveFormat.Doc);
            }

            Comment docComment = (Comment)ai._doc.GetChild(NodeType.Comment, 0, true);
            NodeCollection allComments = ai._doc.GetChildNodes(NodeType.Comment, true);

            //Assert
            Assert.Single(allComments);
            Assert.True(docComment.GetText().IndexOf("应为《城市桥梁设计规范》（CJJ 11-2011）") >= 0);
        }
        #endregion

        #region _FindSequenceNumberError
        [Fact]
        public void _FindSequenceNumberError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\_FindSequenceNumberError.doc";
            var ai = new AsposeAIReportCheck(fileName);
            //Act
            ai._FindSequenceNumberError();

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
