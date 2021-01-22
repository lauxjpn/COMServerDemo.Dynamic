using System;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;

namespace COMClient
{
    class Program
    {
        static void Main(string[] args)
        {
            var serverType = Type.GetTypeFromCLSID(new Guid(ContractGuids.ServerClass));
            var serverObject = Activator.CreateInstance(serverType);

            // ITypeInfo is available and returns expected data.
            var typeInfo = (ITypeInfo)serverObject;
            
            string[] names = new string[256];
            typeInfo.GetNames(1, names, 256, out var namesCount);
            
            Trace.Assert(namesCount == 1);
            Trace.Assert(names[0] == "ComputePi");

            // This works:
            // var server = (IServer) serverObject;

            // This does not work:
            dynamic server = serverObject;

            var pi = server.ComputePi();
            Console.WriteLine($"\u03C0 = {pi}");
        }
    }
}
