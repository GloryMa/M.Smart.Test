using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MST.Engine
{
    class UniversalMethod
    {
        public static void ClassMethodCaller(string className, string methodName,object[] o)
        {            
            Type t = Type.GetType(className);//FullName
            var instance = Activator.CreateInstance(t);
            t.InvokeMember(methodName, BindingFlags.Public | BindingFlags.Instance
                | BindingFlags.InvokeMethod, null, instance, o);           
        }
    }
}
