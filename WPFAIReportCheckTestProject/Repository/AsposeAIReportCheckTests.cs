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
            var ai = new AsposeAIReportCheck(@"..\..\..\TestFiles\DefaultTestFile.doc");
            //Act
            ai._FindNotExplainComponentNo();
            //Assert
            Assert.Equal(WarnningNumber.NotClearInfo, ai.reportWarnning[0].No);
           
        }
        #endregion
    }
}
