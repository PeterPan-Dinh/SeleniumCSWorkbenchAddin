using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumGendKS.Core.BaseClass
{
    public abstract class BaseSingletonThreadSafeNestedConstructors<T>
    {
        public static T? Instance
        {
            get 
            {
                return SingleonFactory.Instance; 
            }
        }

        public static T? Reset
        {
            get
            {
                return default(T);
            }
        }

        internal static class SingleonFactory
        {
            internal static T Instance;
            static SingleonFactory()
            {
                CreateInstance(typeof(T));
            }

            public static T CreateInstance(Type type)
            {
                ConstructorInfo[] ctorsPublic = type.GetConstructors(BindingFlags.Instance | BindingFlags.Public);
                if(ctorsPublic.Length > 0)
                {
                    throw new Exception(string.Concat(type.FullName, "has one or more public constructors so the property cannot be enforced."));
                }

                ConstructorInfo nonPublicConstructor = type.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new Type[0], new ParameterModifier[0]);

                if (nonPublicConstructor == null)
                {
                    throw new Exception(string.Concat(type.FullName, " does not have a private/protected constructor so the property cannot be enforced."));
                }

                try
                {
                    return Instance = (T)nonPublicConstructor.Invoke(new object[0]);
                }
                catch (Exception e)
                {
                    throw new Exception(
                        string.Concat("The Singleton could not be constructed. Check if ", type.FullName, " has a default constructor."), e);
                }
            }
        }
    }
}
