using System;

namespace ClipboardRedux
{
    internal unsafe struct IUnknownVTable
    {
        public static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        // IUnknown
#if NET50_OR_GREATER
        public delegate* unmanaged[Stdcall]<IntPtr, Guid*, IntPtr*, int> QueryInterface;
        public delegate* unmanaged[Stdcall]<IntPtr, int> AddRef;
        public delegate* unmanaged[Stdcall]<IntPtr, int> Release;
#else
        public IntPtr QueryInterface;
        public IntPtr AddRef;
        public IntPtr Release;

        public unsafe delegate int QueryInterfaceDelegate(IntPtr _this, Guid* iid, IntPtr* obj);
        public delegate uint AddRefDelegate(IntPtr _this);
        public delegate uint ReleaseDelegate(IntPtr _this);
#endif
    }
}
