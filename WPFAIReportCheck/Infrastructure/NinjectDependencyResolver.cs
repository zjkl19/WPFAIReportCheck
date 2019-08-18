﻿using log4net;
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
        private ILog _log;
        /// <summary>
        /// Ninject依赖注入解析
        /// </summary>
        /// <param name="doc">文件名</param>
        /// <param name="log">interface ILog : ILoggerWrapper</param>
        public NinjectDependencyResolver(string doc,ILog log)
        {
            _doc = doc;
            _log = log;
        }
        public override void Load()
        {
            Bind<ILog>().ToSelf().InSingletonScope();
            Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _doc).WithConstructorArgument("log", _log);
            //Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _doc).WithConstructorArgument("log", _log);
        }
    }
}
