using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace COMServer
{
    [ComVisible(true)]
    [Guid(ContractGuids.ServerClass)]
    [ProgId("ComServerTlb.ServerTlb")]
    [ComDefaultInterface(typeof(IServer))]
    public class Server : IServer, ITypeInfo
    {
        private ITypeInfo _typeInfo;

        public Server()
        {
            var libid = new Guid(ContractGuids.TypeLibrary);
            Marshal.ThrowExceptionForHR(
                TypeLib.OleAut32.LoadRegTypeLib(ref libid, 1, 0, 0, out var typeLib));
            
            var iidIServer = new Guid(ContractGuids.ServerInterface);
            typeLib.GetTypeInfoOfGuid(ref iidIServer, out _typeInfo);
        }

        ~Server()
        {
            _typeInfo = null;
        }
        
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            try
            { 
                TypeLib.Register(Path.GetFullPath("COMServer.tlb"));
                Console.Beep();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error registering 'COMServer.tlb':\r\n\r\n" + ex);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                TypeLib.Unregister(Path.GetFullPath("COMServer.tlb"));
                Console.Beep();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error unregistering 'COMServer.tlb':\r\n\r\n" + ex);
            }
        }
        
        public double ComputePi()
        {
            double sum = 0.0;
            int sign = 1;
            for (int i = 0; i < 1024; ++i)
            {
                sum += sign / (2.0 * i + 1.0);
                sign *= -1;
            }

            return 4.0 * sum;
        }
        
        double IServer.ComputePi()
        {
            //return ComputePi();
            double sum = 0.0;
            int sign = 1;
            for (int i = 0; i < 1024; ++i)
            {
                sum += sign / (2.0 * i + 1.0);
                sign *= -1;
            }

            return 4.0 * sum;
        }
        
        #region ITypeInfo

        void ITypeInfo.AddressOfMember(int memid, INVOKEKIND invKind, out IntPtr ppv)
            => _typeInfo.AddressOfMember(memid, invKind, out ppv);

        void ITypeInfo.CreateInstance(object? pUnkOuter, ref Guid riid, out object ppvObj)
            => _typeInfo.CreateInstance(pUnkOuter, ref riid, out ppvObj);

        void ITypeInfo.GetContainingTypeLib(out ITypeLib ppTLB, out int pIndex)
            => _typeInfo.GetContainingTypeLib(out ppTLB, out pIndex);

        void ITypeInfo.GetDllEntry(int memid, INVOKEKIND invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal)
            => _typeInfo.GetDllEntry(memid, invKind, pBstrDllName, pBstrName, pwOrdinal);

        void ITypeInfo.GetDocumentation(int index, out string strName, out string strDocString, out int dwHelpContext, out string strHelpFile)
            => _typeInfo.GetDocumentation(index, out strName, out strDocString, out dwHelpContext, out strHelpFile);

        void ITypeInfo.GetFuncDesc(int index, out IntPtr ppFuncDesc)
            => _typeInfo.GetFuncDesc(index, out ppFuncDesc);

        void ITypeInfo.GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId)
            => _typeInfo.GetIDsOfNames(rgszNames, cNames, pMemId);

        void ITypeInfo.GetImplTypeFlags(int index, out IMPLTYPEFLAGS pImplTypeFlags)
            => _typeInfo.GetImplTypeFlags(index, out pImplTypeFlags);

        void ITypeInfo.GetMops(int memid, out string? pBstrMops)
            => _typeInfo.GetMops(memid, out pBstrMops);

        void ITypeInfo.GetNames(int memid, string[] rgBstrNames, int cMaxNames, out int pcNames)
            => _typeInfo.GetNames(memid, rgBstrNames, cMaxNames, out pcNames);

        void ITypeInfo.GetRefTypeInfo(int hRef, out ITypeInfo ppTI)
            => _typeInfo.GetRefTypeInfo(hRef, out ppTI);

        void ITypeInfo.GetRefTypeOfImplType(int index, out int href)
            => _typeInfo.GetRefTypeOfImplType(index, out href);

        void ITypeInfo.GetTypeAttr(out IntPtr ppTypeAttr)
            => _typeInfo.GetTypeAttr(out ppTypeAttr);

        void ITypeInfo.GetTypeComp(out ITypeComp ppTComp)
            => _typeInfo.GetTypeComp(out ppTComp);

        void ITypeInfo.GetVarDesc(int index, out IntPtr ppVarDesc)
            => _typeInfo.GetVarDesc(index, out ppVarDesc);

        void ITypeInfo.Invoke(object pvInstance, int memid, short wFlags, ref DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, out int puArgErr)
            => _typeInfo.Invoke(pvInstance, memid, wFlags, ref pDispParams, pVarResult, pExcepInfo, out puArgErr);

        void ITypeInfo.ReleaseFuncDesc(IntPtr pFuncDesc)
            => _typeInfo.ReleaseFuncDesc(pFuncDesc);

        void ITypeInfo.ReleaseTypeAttr(IntPtr pTypeAttr)
            => _typeInfo.ReleaseTypeAttr(pTypeAttr);

        void ITypeInfo.ReleaseVarDesc(IntPtr pVarDesc)
            => _typeInfo.ReleaseVarDesc(pVarDesc);

        #endregion
    }
}
