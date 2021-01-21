using System;

namespace COMClient
{
    class Program
    {
        static void Main(string[] args)
        {
            var serverType = Type.GetTypeFromCLSID(new Guid(ContractGuids.ServerClass));
            var serverObject = Activator.CreateInstance(serverType);

            // This works:
            // var server = (IServer) serverObject;

            // This does not work:
            dynamic server = serverObject;

            var pi = server.ComputePi();
            Console.WriteLine($"\u03C0 = {pi}");
        }
    }
}
