﻿using System;
using System.Collections.Generic;
using System.Text;
using Xunit;
using WPFAIReportCheck.Repository;
using WPFAIReportCheck.Models;
//using Aspose.Words;
using System.IO;
//using log4net;
using Moq;
using System.Xml.Linq;
using Aspose.Words;
using NLog;

namespace WPFAIReportCheckTestProject.Repository
{
    public class AsposeAIReportCheckTestsFixture : IDisposable
    {
        public Mock<ILogger> log;

        public AsposeAIReportCheckTestsFixture()
        {
            var _log = new Mock<ILogger>();
            _log.Setup(m => m.Error(It.IsAny<Exception>(), It.IsAny<string>()));   //无实际意义，仅作为1必须参数传入
            log = _log;
        }

        public void Dispose()
        {
        }
    }
    public partial class AsposeAIReportCheckTests : IClassFixture<AsposeAIReportCheckTestsFixture>
    {
        AsposeAIReportCheckTestsFixture _fixture;

        public AsposeAIReportCheckTests(AsposeAIReportCheckTestsFixture fixture)
        {
            _fixture = fixture;
        }

        #region _FindUnitError
        [Fact]
        public void _FindUnitError()
        {
            //Arrange
            var fileName = @"..\..\..\TestFiles\_FindUnitError.doc";

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
            Assert.True(docComment.GetText().IndexOf("应为km/h") >= 0);
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
