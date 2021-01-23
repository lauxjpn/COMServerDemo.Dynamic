using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

using ComTypes = System.Runtime.InteropServices.ComTypes;

[assembly: Guid(ContractGuids.TypeLibrary)]
[assembly: ImportedFromTypeLib("COMServer.tlb")]

namespace COMServer
{
    internal static class TypeLib
    {
        public static void Register(string tlbPath)
        {
            var ctl = new ConsoleTraceListener();
            Trace.Listeners.Add(ctl);

            try
            {
                Trace.WriteLine($"Registering type library:");
                Trace.Indent();
                Trace.WriteLine(tlbPath);
                Trace.Unindent();

                Marshal.ThrowExceptionForHR(
                    OleAut32.LoadTypeLibEx(tlbPath, OleAut32.REGKIND.REGKIND_REGISTER, out ComTypes.ITypeLib _));
            }
            finally
            {
                Trace.Listeners.Remove(ctl);
            }
        }

        public static void Unregister(string tlbPath)
        {
            var ctl = new ConsoleTraceListener();
            Trace.Listeners.Add(ctl);

            try
            {
                Trace.WriteLine($"Unregistering type library:");
                Trace.Indent();
                Trace.WriteLine(tlbPath);
                Trace.Unindent();

                int hr = OleAut32.LoadTypeLibEx(tlbPath, OleAut32.REGKIND.REGKIND_NONE, out var typeLib);
                if (hr < 0)
                {
                    Trace.WriteLine($"Unregistering type library failed: 0x{hr:x}");
                    return;
                }

                IntPtr attrPtr = IntPtr.Zero;
                try
                {
                    typeLib.GetLibAttr(out attrPtr);
                    if (attrPtr != IntPtr.Zero)
                    {
                        ComTypes.TYPELIBATTR attr = Marshal.PtrToStructure<ComTypes.TYPELIBATTR>(attrPtr);
                        hr = OleAut32.UnRegisterTypeLib(ref attr.guid, attr.wMajorVerNum, attr.wMinorVerNum, attr.lcid, attr.syskind);
                        if (hr < 0)
                        {
                            Trace.WriteLine($"Unregistering type library failed: 0x{hr:x}");
                        }
                    }
                }
                finally
                {
                    if (attrPtr != IntPtr.Zero)
                    {
                        typeLib.ReleaseTLibAttr(attrPtr);
                    }
                }
            }
            finally
            {
                Trace.Listeners.Remove(ctl);
            }
        }

        public class OleAut32
        {
            // https://docs.microsoft.com/windows/api/oleauto/ne-oleauto-regkind
            public enum REGKIND
            {
                REGKIND_DEFAULT = 0,
                REGKIND_REGISTER = 1,
                REGKIND_NONE = 2
            }

            // https://docs.microsoft.com/windows/api/oleauto/nf-oleauto-loadtypelibex
            [DllImport(nameof(OleAut32), CharSet = CharSet.Unicode, ExactSpelling = true)]
            public static extern int LoadTypeLibEx(
                [In, MarshalAs(UnmanagedType.LPWStr)] string fileName,
                REGKIND regKind,
                out ComTypes.ITypeLib typeLib);

            // https://docs.microsoft.com/windows/api/oleauto/nf-oleauto-unregistertypelib
            [DllImport(nameof(OleAut32))]
            public static extern int UnRegisterTypeLib(
                ref Guid id,
                short majorVersion,
                short minorVersion,
                int lcid,
                ComTypes.SYSKIND sysKind);
            
            // https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-loadregtypelib
            [DllImport(nameof(OleAut32))]
            public static extern int LoadRegTypeLib(
                ref Guid rguid,
                short majorVersion,
                short minorVersion,
                int lcid,
                out ComTypes.ITypeLib pptlib);
        }
    }
}
