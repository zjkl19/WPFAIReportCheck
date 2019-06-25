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
        private Aspose.Words.Document _d;
        public NinjectDependencyResolver(Aspose.Words.Document d)
        {
            _d = d;
        }
        public override void Load()
        {
            Bind<IAIReportCheck>().To<AsposeAIReportCheck>().WithConstructorArgument("doc", _d);
        }
    }
}
