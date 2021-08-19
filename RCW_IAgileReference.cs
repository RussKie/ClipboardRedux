using System;
using System.Runtime.InteropServices;

namespace ClipboardRedux
{
    internal class RCW_IAgileReference
    {
        private unsafe struct AgileReferenceVTable
        {
#pragma warning disable CS0649
            // Due to a bug in .NET Core 3.1 (3.1.18 as of now) we must inline IUnknown interface
            // members explicitly, and can't do the following:
            //
            //      public IUnknownVTable UnknownVTable;
            public delegate* unmanaged[Stdcall]<IntPtr, Guid*, IntPtr*, HRESULT> QueryInterface;
            public delegate* unmanaged[Stdcall]<IntPtr, HRESULT> AddRef;
            public delegate* unmanaged[Stdcall]<IntPtr, HRESULT> Release;

            // IAgileReference
            public delegate* unmanaged[Stdcall]<IntPtr, ref Guid, out IntPtr, HRESULT> Resolve;
#pragma warning restore CS0649
        }

        private readonly IntPtr _instance;
        private readonly unsafe AgileReferenceVTable* _vtable;

        public unsafe RCW_IAgileReference(IntPtr instance)
        {
            _instance = instance;
            _vtable = *(AgileReferenceVTable**)instance;
        }

        // An IAgileReference instance handles release on the correct context.
        ~RCW_IAgileReference()
        {
            if (_instance != IntPtr.Zero)
            {
                Marshal.Release(_instance);
            }
        }

        public unsafe IntPtr Resolve(Guid iid)
        {
            // Marshal and dispatch
            _vtable->Resolve(_instance, ref iid, out IntPtr resolvedInstance)
                .ThrowIfFailed();

            // Unmarshal
            return resolvedInstance;
        }
    }
}
