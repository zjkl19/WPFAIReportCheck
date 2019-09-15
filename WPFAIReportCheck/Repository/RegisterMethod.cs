using System.Collections.Generic;
using WPFAIReportCheck.IRepository;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 注册校核的方法，需要使用的方法添加到SelectListFunctionName
        /// </summary>
        internal ControlClass RegisterMethod()
        {
            return new ControlClass
            {
                SelectListFunctionName = new List<MethodDelegate> {
                    FindUnitError,FindNotExplainComponentNoWarnning,FindSpecificationsError,
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