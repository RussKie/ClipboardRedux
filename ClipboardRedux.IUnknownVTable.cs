using System;

namespace WindowsFormsClipboardRedux
{
    partial class ClipboardRedux
    {
        internal unsafe struct IUnknownVTable
        {
            public static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

            // IUnknown
            public IntPtr QueryInterface;
            public IntPtr AddRef;
            public IntPtr Release;

            public unsafe delegate int QueryInterfaceDelegate(IntPtr _this, Guid* iid, IntPtr* obj);
            public delegate uint AddRefDelegate(IntPtr _this);
            public delegate uint ReleaseDelegate(IntPtr _this);
        }
    }
}
