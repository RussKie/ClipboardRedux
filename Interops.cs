using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

// Clean up warnings
#pragma warning disable CS0649
#pragma warning disable CA1416

namespace ClipboardRedux
{
    internal static class Interops
    {
        internal static class Ole32
        {
            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern HRESULT OleInitialize(IntPtr reserved);

            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern HRESULT OleGetClipboard(out IntPtr dataObject);

            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern HRESULT OleSetClipboard(IntPtr dataObject);

            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern HRESULT OleFlushClipboard();

            public enum AgileReferenceOptions
            {
                Default = 0,
                DelayedMarshal = 1,
            };

            // The RoGetAgileReference API is supported on Windows 8.1+.
            //   See: https://docs.microsoft.com/windows/win32/api/combaseapi/nf-combaseapi-rogetagilereference
            // For prior OS versions use the Global Interface Table (GIT).
            //   See: https://docs.microsoft.com/windows/win32/com/creating-the-global-interface-table
            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern HRESULT RoGetAgileReference(
                AgileReferenceOptions opts,
                ref Guid riid,
                IntPtr instance,
                out IntPtr agileReference);
        }

        internal struct STGMEDIUM
        {
            public TYMED tymed;
            public IntPtr unionmember;
            public IntPtr pUnkForRelease;
        }

    }

}
