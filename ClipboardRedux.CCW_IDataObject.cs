using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Forms_IDataObject = System.Windows.Forms.IDataObject;

namespace WindowsFormsClipboardRedux
{
    partial class ClipboardRedux
    {
        internal static class CCW_IDataObject
        {
            private static readonly List<Delegate> s_delegates = new List<Delegate>();
            private static readonly Lazy<IntPtr> s_ccwVTable = new Lazy<IntPtr>(AllocateVTable, isThreadSafe: true);

            private unsafe struct Lifetime
            {
                public DataObjectVTable* VTable;
                public IntPtr Handle;
                public int RefCount;
            }

            public unsafe static IntPtr CreateInstance(Forms_IDataObject dataObject)
            {
                var wrapper = (Lifetime*)Marshal.AllocCoTaskMem(sizeof(Lifetime));

                // Create the wrapper instance.
                wrapper->VTable = (DataObjectVTable*)s_ccwVTable.Value;
                wrapper->Handle = GCHandle.ToIntPtr(GCHandle.Alloc(dataObject));
                wrapper->RefCount = 1;

                return (IntPtr)wrapper;
            }

            public unsafe static bool TryGetInstance(IntPtr instanceMaybe, [NotNullWhen(true)] out Forms_IDataObject? forms_dataObject)
            {
                forms_dataObject = null;

                // This is a dangerous cast since it relies on strictly
                // following the COM ABI. If the VTable is ours the rest of
                // the structure is good, otherwise it is unknown.
                var lifetime = (Lifetime*)instanceMaybe;
                if (lifetime->VTable != (DataObjectVTable*)s_ccwVTable.Value)
                {
                    return false;
                }

                forms_dataObject = GetInstance(instanceMaybe);
                return true;
            }

            private unsafe static IntPtr AllocateVTable()
            {
                // Allocate and create a singular VTable for this type projection.
                var vtable = (DataObjectVTable*)Marshal.AllocCoTaskMem(sizeof(DataObjectVTable));

                var queryInterfaceDelegate = new IUnknownVTable.QueryInterfaceDelegate(QueryInterface);
                var addRefDelegate = new IUnknownVTable.AddRefDelegate(AddRef);
                var releaseDelegate = new IUnknownVTable.ReleaseDelegate(Release);
                var getDataDelegate = new DataObjectVTable.GetDataDelegate(GetData);
                var getDataHereDelegate = new DataObjectVTable.GetDataHereDelegate(GetDataHere);
                var queryGetDataDelegate = new DataObjectVTable.QueryGetDataDelegate(QueryGetData);
                var getCanonicalFormatEtcDelegate = new DataObjectVTable.GetCanonicalFormatEtcDelegate(GetCanonicalFormatEtc);
                var setDataDelegate = new DataObjectVTable.SetDataDelegate(SetData);
                var enumFormatEtcDelegate = new DataObjectVTable.EnumFormatEtcDelegate(EnumFormatEtc);
                var dAdviseDelegate = new DataObjectVTable.DAdviseDelegate(DAdvise);
                var dUnadviseDelegate = new DataObjectVTable.DUnadviseDelegate(DUnadvise);
                var enumDAdviseDelegate = new DataObjectVTable.EnumDAdviseDelegate(EnumDAdvise);

                s_delegates.Add(queryInterfaceDelegate);
                s_delegates.Add(addRefDelegate);
                s_delegates.Add(releaseDelegate);
                s_delegates.Add(getDataDelegate);
                s_delegates.Add(getDataHereDelegate);
                s_delegates.Add(queryGetDataDelegate);
                s_delegates.Add(getCanonicalFormatEtcDelegate);
                s_delegates.Add(setDataDelegate);
                s_delegates.Add(enumFormatEtcDelegate);
                s_delegates.Add(dAdviseDelegate);
                s_delegates.Add(dUnadviseDelegate);
                s_delegates.Add(enumDAdviseDelegate);

                // IUnknown
                vtable->UnknownVTable.QueryInterface = Marshal.GetFunctionPointerForDelegate(queryInterfaceDelegate);
                vtable->UnknownVTable.AddRef = Marshal.GetFunctionPointerForDelegate(addRefDelegate);
                vtable->UnknownVTable.Release = Marshal.GetFunctionPointerForDelegate(releaseDelegate);

                // IDataObject
                vtable->GetData = Marshal.GetFunctionPointerForDelegate(getDataDelegate);
                vtable->GetDataHere = Marshal.GetFunctionPointerForDelegate(getDataDelegate);
                vtable->GetData = Marshal.GetFunctionPointerForDelegate(getDataDelegate);
                vtable->GetDataHere = Marshal.GetFunctionPointerForDelegate(getDataHereDelegate);
                vtable->QueryGetData = Marshal.GetFunctionPointerForDelegate(queryGetDataDelegate);
                vtable->GetCanonicalFormatEtc = Marshal.GetFunctionPointerForDelegate(getCanonicalFormatEtcDelegate);
                vtable->SetData = Marshal.GetFunctionPointerForDelegate(setDataDelegate);
                vtable->EnumFormatEtc = Marshal.GetFunctionPointerForDelegate(enumFormatEtcDelegate);
                vtable->DAdvise = Marshal.GetFunctionPointerForDelegate(dAdviseDelegate);
                vtable->DUnadvise = Marshal.GetFunctionPointerForDelegate(dUnadviseDelegate);
                vtable->EnumDAdvise = Marshal.GetFunctionPointerForDelegate(enumDAdviseDelegate);

                return (IntPtr)vtable;
            }

            private unsafe static Forms_IDataObject GetInstance(IntPtr wrapper)
            {
                var lifetime = (Lifetime*)wrapper;

                Debug.Assert(lifetime->Handle != IntPtr.Zero);
                Debug.Assert(lifetime->RefCount > 0);
                return (Forms_IDataObject)GCHandle.FromIntPtr(lifetime->Handle).Target!;
            }

            // IUnknown
            private unsafe static int QueryInterface(IntPtr _this, Guid* iid, IntPtr* ppObject)
            {
                if (*iid != IUnknownVTable.IID_IUnknown && *iid != DataObjectVTable.IID_IDataObject)
                {
                    *ppObject = IntPtr.Zero;
                    const int E_NOINTERFACE = unchecked((int)0x80004002L);
                    return E_NOINTERFACE;
                }

                *ppObject = _this;
                AddRef(_this);
                return 0;
            }

            private unsafe static uint AddRef(IntPtr _this)
            {
                var lifetime = (Lifetime*)_this;
                Debug.Assert(lifetime->Handle != IntPtr.Zero);
                Debug.Assert(lifetime->RefCount > 0);
                return (uint)Interlocked.Increment(ref lifetime->RefCount);
            }

            private unsafe static uint Release(IntPtr _this)
            {
                var lifetime = (Lifetime*)_this;
                Debug.Assert(lifetime->Handle != IntPtr.Zero);
                Debug.Assert(lifetime->RefCount > 0);
                uint count = (uint)Interlocked.Decrement(ref lifetime->RefCount);
                if (count == 0)
                {
                    GCHandle.FromIntPtr(lifetime->Handle).Free();
                    lifetime->Handle = IntPtr.Zero;
                }

                return count;
            }


            // IDataObject
            private unsafe static HRESULT GetData(IntPtr _this, FORMATETC* format, Interops.STGMEDIUM* medium)
            {
                *medium = default;
                return HRESULT.E_NOTIMPL;
            }

            private unsafe static HRESULT GetDataHere(IntPtr _this, FORMATETC* format, Interops.STGMEDIUM* medium)
            {
                return HRESULT.E_NOTIMPL;
            }

            private unsafe static HRESULT QueryGetData(IntPtr _this, /*optional*/ FORMATETC* format)
            {
                return HRESULT.E_NOTIMPL;
            }

            private unsafe static HRESULT GetCanonicalFormatEtc(IntPtr _this, /*optional*/ FORMATETC* formatIn, FORMATETC* formatOut)
            {
                *formatOut = default;
                return HRESULT.E_NOTIMPL;
            }

            private unsafe static HRESULT SetData(IntPtr _this, FORMATETC* format, Interops.STGMEDIUM* medium, int shouldRelease)
            {
                return HRESULT.E_NOTIMPL;
            }

            private unsafe static HRESULT EnumFormatEtc(IntPtr _this, int direction, IntPtr* enumFORMATETC)
            {
                *enumFORMATETC = default;
                return HRESULT.E_NOTIMPL;
            }

            private unsafe static HRESULT DAdvise(IntPtr _this, FORMATETC* format, int advf, IntPtr adviseSink, int* connection)
            {
                *connection = default;
                return HRESULT.E_NOTIMPL;
            }

            private static HRESULT DUnadvise(IntPtr _this, int connection)
            {
                return HRESULT.E_NOTIMPL;
            }

            private unsafe static HRESULT EnumDAdvise(IntPtr _this, IntPtr* enumSTATDATA)
            {
                *enumSTATDATA = default;
                return HRESULT.E_NOTIMPL;
            }
        }
    }
}
