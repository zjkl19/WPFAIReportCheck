
using System.Collections.Generic;
using WPFAIReportCheck.Repository;

namespace WPFAIReportCheck.Infrastructure
{
    /// <summary>
    /// 注册校核方法
    /// </summary>
    /// <remarks>
    /// 注意：
    /// Option.config文件也要进行相应补充
    /// 如果有额外的选项，AIReportCheck.config也要进行相应补充
    /// </remarks>
    //public static class RegisterMethod
    //{
    //    private delegate void MethodDelegate();
    //    //举例：
    //    //SelectListInt =[1,0]
    //    //SelectListFunctionName=[FindUnitError,FindNotExplainComponentNo]
    //    //表示算法只算FindUnitError,不算FindNotExplainComponentNo
    //    public ControlClass controlClass = new
    //    {
    //        SelectListInt = new List<int>(),
    //        SelectListFunctionName = new List<MethodDelegate> {
    //                _FindUnitError,_FindNotExplainComponentNo,_FindSpecificationsError,
    //                FindSequenceNumberError,FindStrainOrDispError,FindDescriptionError,
    //                FindOtherBridgesWarnning,FindStrainOrDispContextWarnning,FindPageContextError,
    //                FindTableTitleInSinglePageWarnning,FindPictureTitleInSinglePageWarnning,FindTableTitleSequenceNumberError,
    //                FindPictureTitleSequenceNumberError
    //            },
    //        SelectListContent = new List<string>(),
    //    };
    //}

    //public class ControlClass
    //{
    //    public delegate void MethodDelegate();
    //    public List<int> SelectListInt { get; set; }
    //    public List<MethodDelegate> SelectListFunctionName { get; set; }= new List<MethodDelegate> {
    //                _FindUnitError,_FindNotExplainComponentNo,_FindSpecificationsError,
    //                FindSequenceNumberError,FindStrainOrDispError,FindDescriptionError,
    //                FindOtherBridgesWarnning,FindStrainOrDispContextWarnning,FindPageContextError,
    //                FindTableTitleInSinglePageWarnning,FindPictureTitleInSinglePageWarnning,FindTableTitleSequenceNumberError,
    //                FindPictureTitleSequenceNumberError }

    //    public List<string> SelectListContent{ get; set; }
    //}
}
