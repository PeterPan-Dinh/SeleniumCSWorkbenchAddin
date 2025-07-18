using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumGendKS.Core.BaseClass
{
    public class BasePageValidator<S, M, V>
        where S : BasePage<S, M, V>
        where M : BasePageElementMap, new()
        where V : BasePageValidator<S, M, V>, new()
    {
        protected S pageInstance;

        public BasePageValidator(S currentInstance)
        {
            this.pageInstance = currentInstance;
        }

        public BasePageValidator()
        { }

        protected M Map 
        {
            get { return new M(); }
        }
    }
}
