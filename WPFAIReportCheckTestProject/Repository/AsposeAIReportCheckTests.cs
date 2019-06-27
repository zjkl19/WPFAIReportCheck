using System;
using System.Collections.Generic;
using System.Text;
using Xunit;
using WPFAIReportCheck.Repository;
using WPFAIReportCheck.Models;
using Aspose.Words;

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
    }
}
