using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Repository;

namespace WPFAIReportCheck.Infrastructure
{
    class NinjectDependencyResolver : Ninject.Modules.NinjectModule
    {
        private string _doc;
        /// <summary>
        /// Ninject依赖注入解析
        /// </summary>
        /// <param name="doc">文件名</param>
        /// <param name="log">interface ILog : ILoggerWrapper</param>
        public NinjectDependencyResolver(string doc)
        {
            _doc = doc;
        }
        public override void Load()
        {
            Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _doc);
            //Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _doc).WithConstructorArgument("log", _log);
        }
    }
}
