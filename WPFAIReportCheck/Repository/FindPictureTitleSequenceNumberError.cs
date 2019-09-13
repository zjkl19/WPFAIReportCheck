using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using System.Linq;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;
using Aspose.Words.Replacing;

namespace WPFAIReportCheck.Repository
{

    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 查找报告中图片序号的错误
        /// </summary>
        /// <remarks>
        /// 要求：
        /// 图片标题仅支持类似：“图 1-2，图3-2”等，最大支持到“图 999-999”
        /// 算法：
        /// 1、在全文用正则表达式匹配图片对应的描述字符@"图[\s]{0,1}[1-9][0-9]?[0-9]?[\s]{0,1}-[\s]{0,1}[1-9][0-9]?[0-9]?[\s]{1}(.*)"查找，将匹配到的第一个结果存到list
        /// 3、将list中的序号从小到大排列，将排序后的list用如下算法进行验证，看是否有错：
        /// 3.1、提取章节号及表序号，第1张图片：一定是x-1
        /// 3.2、第2~n张图片：
        /// 3.3、若章节号同上一张图片，则该图片序号-上一个图片序号==1
        /// 3.4、若章节号不同于一张图片，则图片序号==1
        /// </remarks>
        /// 算法主要参与人员：林迪南
        public void FindPictureTitleSequenceNumberError()
        {
            FindReplaceOptions options;
            MatchCollection matches;

            List<string> list = new List<string>();
            var convertList= new List<string>();    //替换空格后的list
            var regex = new Regex(@"\b图[\s]{0,2}[1-9][0-9]?[0-9]?[\s]{0,2}-[\s]{0,2}[1-9][0-9]?[0-9]?[' ']{1,2}(.[^，。]*?)(?=[\r|\a])");
            int previousChapterNum = 0; int previousPictureSequenceNum = 0; int currChapterNum = 0; int currPictureSequenceNum = 0;
            //var regex = new Regex(@"(?<=\(附[\s]{0,4}页\)\r目录\r)[\s\S]*?(?=\f)"); 
            var chapterNumRegex = new Regex(@"(?<=图)[1-9][0-9]?[0-9]?(?=\-)");
            var pictureSequenceNumRegex = new Regex(@"(?<=\-)[1-9][0-9]?[0-9]?");

            var regex1 = new Regex(@"\b图[\s]{0,2}[1-9][0-9]?[0-9]?[\s]{0,2}-[\s]{0,2}[1-9][0-9]?[0-9]?");

            try
            {
                var t = _originalWholeText.Replace("\u001e", "-");
                matches = regex.Matches(t);

                //TODO：考虑报告一张图片都没有的情况
                if (matches.Count > 0)
                {
                    foreach (Match m in matches)
                    {
                        list.Add(m.Value.ToString());
                        convertList.Add(regex1.Match(m.Value.ToString()).Value.ToString().Replace(" ",""));
                    }
                }

                for (int i = 0; i < list.Count; i++)
                {
                    currChapterNum = Convert.ToInt32(chapterNumRegex.Match(convertList[i]).Value);
                    currPictureSequenceNum = Convert.ToInt32(pictureSequenceNumRegex.Match(convertList[i]).Value);
                    if (i == 0)
                    {
                        previousChapterNum = Convert.ToInt32(chapterNumRegex.Match(convertList[i]).Value);
                        previousPictureSequenceNum = Convert.ToInt32(pictureSequenceNumRegex.Match(convertList[i]).Value);
                        if (currPictureSequenceNum != 1)
                        {
                            reportError.Add(new ReportError(ErrorNumber.Description, $"第1张图", $"第1张图起始序号应为1", true));

                            options = new FindReplaceOptions
                            {
                                ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", "第1张图起始序号应为1"),
                                Direction = FindReplaceDirection.Forward
                            };
                            _doc.Range.Replace(list[i], "", options);
                        }
                    }
                    else
                    {
                        if (currChapterNum == previousChapterNum)
                        {
                            if (currPictureSequenceNum - previousPictureSequenceNum != 1)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Description, $"第{i + 1}张图", $"第{i + 1}张图序号为{currChapterNum}-{currPictureSequenceNum}，实际应为{currChapterNum}-{previousPictureSequenceNum + 1}", true));
                                options = new FindReplaceOptions
                                {
                                    ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", $"本图片序号应为{currChapterNum}-{previousPictureSequenceNum + 1}"),
                                    Direction = FindReplaceDirection.Forward
                                };
                                _doc.Range.Replace(list[i], "", options);
                            }
                        }
                        else   //(currChapterNum>previousChapterNum)
                        {
                            if (currPictureSequenceNum != 1)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Description, $"第{i + 1}张图", $"第{i + 1}张图序号为{currChapterNum}-{currPictureSequenceNum}，实际应为{currChapterNum}-1", true));

                                options = new FindReplaceOptions
                                {
                                    ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", $"本图片序号应为{currChapterNum}-1"),
                                    Direction = FindReplaceDirection.Forward
                                };
                                _doc.Range.Replace(list[i], "", options);
                            }
                        }
                    }
                    previousChapterNum = currChapterNum;
                    previousPictureSequenceNum = currPictureSequenceNum;
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;

#else
                _log.Error(ex, $"FindPictureTitleSequenceNumberError运行出错，错误信息：{ ex.Message.ToString()}");
#endif
            }

        }
    }
}
