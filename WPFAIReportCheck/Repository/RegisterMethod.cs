

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

using System.Collections.Generic;
using WPFAIReportCheck.IRepository;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 需要使用的方法添加到SelectListFunctionName
        /// </summary>
        /// <returns></returns>
        internal ControlClass RegisterMethod()
        {
            return new ControlClass
            {
                SelectListFunctionName = new List<MethodDelegate> {
                    _FindUnitError,_FindNotExplainComponentNo,_FindSpecificationsError,
                    FindSequenceNumberError,FindStrainOrDispError,FindDescriptionError,
                    FindOtherBridgesWarnning,FindStrainOrDispContextWarnning,FindPageContextError,
                    FindTableTitleInSinglePageWarnning,FindPictureTitleInSinglePageWarnning,FindTableTitleSequenceNumberError,
                    FindPictureTitleSequenceNumberError
                },
                SelectListInt = new List<int>(),
                SelectListContent = new List<string>()
            };
        }

        /// <summary>
        /// 控制程序哪些方法需要使用，哪些方法不需要使用
        /// </summary>
        //举例：
        //SelectListInt =[1,0]
        //SelectListFunctionName=[FindUnitError,FindNotExplainComponentNo]
        //表示算法只算FindUnitError,不算FindNotExplainComponentNo
        //SelectListContent会自动匹配
        internal class ControlClass
        {
            public List<int> SelectListInt { get; set; }
            public List<MethodDelegate> SelectListFunctionName { get; set; }
            public List<string> SelectListContent { get; set; }
        }


    }
}