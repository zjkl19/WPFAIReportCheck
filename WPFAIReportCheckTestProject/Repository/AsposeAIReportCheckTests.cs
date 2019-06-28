﻿using System;
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
            Assert.Equal(2, allComments.Count);
            Assert.Equal(1, docComment.Count);
            Assert.True(docComment.GetText().IndexOf("单位错误")>=0);
            Assert.True(docComment1.GetText().IndexOf("单位错误") >= 0);
            //Assert.Equal("\u0005My comment.\r", docComment.GetText());
        }
        #endregion
    }
}
