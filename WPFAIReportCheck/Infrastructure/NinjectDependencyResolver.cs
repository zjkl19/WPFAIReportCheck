//using log4net;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Repository;

namespace WPFAIReportCheck.Infrastructure
{
    class NinjectDependencyResolver : Ninject.Modules.NinjectModule
    {
        private string _doc;
        //private ILog _log;
        private ILogger _log;
        private XDocument _config;
        /// <summary>
        /// Ninject依赖注入解析
        /// </summary>
        /// <param name="doc">文件名</param>
        /// <param name="log">interface ILog : ILoggerWrapper</param>
        public NinjectDependencyResolver(string doc, ILogger log, XDocument config)
        {
            _doc = doc;
            _log = log;
            _config = config;
        }
        public override void Load()
        {
            Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _doc).WithConstructorArgument("log", _log).WithConstructorArgument("config", _config);
            //Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _doc).WithConstructorArgument("log", _log);
        }
    }
}
