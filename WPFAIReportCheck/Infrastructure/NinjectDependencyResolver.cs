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
        public NinjectDependencyResolver(string doc)
        {
            _doc = doc;
        }
        public override void Load()
        {
            Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _doc);
        }
    }
}
