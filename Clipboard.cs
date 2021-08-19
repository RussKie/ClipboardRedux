using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using Forms_IDataObject = System.Windows.Forms.IDataObject;
using Forms_DataFormats = System.Windows.Forms.DataFormats;

// Clean up warnings
#pragma warning disable CS0649
#pragma warning disable CA1416

namespace ClipboardRedux
{
    internal static class ClipboardImpl
    {
        public static void Initialize()
        {
            Interops.Ole32
                .OleInitialize(IntPtr.Zero)
                .ThrowIfFailed();
        }

        public static object Get()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                throw new InvalidOperationException();
            }

            Interops.Ole32.OleGetClipboard(out IntPtr instance).ThrowIfFailed();

            // Check if the returned value is actually a wrapped managed instance.
            if (CCW_IDataObject.TryGetInstance(instance, out Forms_IDataObject? forms_dataObject))
            {
                return forms_dataObject;
            }

            var iid = typeof(IDataObject).GUID;
            HRESULT hr = Interops.Ole32.RoGetAgileReference(
                Interops.Ole32.AgileReferenceOptions.Default,
                ref iid,
                instance,
                out IntPtr agileReference);
            if (hr != HRESULT.S_OK)
            {
                // Release the clipboard object if agile
                // reference creation failed.
                Marshal.Release(instance);
                Marshal.ThrowExceptionForHR((int)hr);
            }

            // Wrap the agile reference.
            var agileRef = new RCW_IAgileReference(agileReference);

            // Release the current instance as it is now controlled
            // by the agile reference RCW.
            Marshal.Release(instance);

            return new RCW_IDataObject(agileRef);
        }

        public static void Set(object obj)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                throw new InvalidOperationException();
            }

            HRESULT hr;
            if (obj is null)
            {
                Interops.Ole32
                    .OleSetClipboard(IntPtr.Zero)
                    .ThrowIfFailed();
                return;
            }

            if (obj is RCW_IDataObject com_dataobject)
            {
                hr = Interops.Ole32.OleSetClipboard(com_dataobject.GetInstanceForSta());
            }
            else if (obj is Forms_IDataObject forms_dataObject)
            {
                // This approach is less than ideal since a new wrapper is always
                // created. Having an efficient cache would be more effective.
                var ccw = CCW_IDataObject.CreateInstance(forms_dataObject);
                hr = Interops.Ole32.OleSetClipboard(ccw);
            }
            else
            {
                // This requires implementing a universal CCW or alternatively
                // leveraging the built-in system. It isn't obvious which one is
                // the best option - both are possible.
                throw new NotImplementedException();
            }

            hr.ThrowIfFailed();
        }

        public static HRESULT Flush()
        {
            return Interops.Ole32.OleFlushClipboard();
        }
    }

    internal class RCW_IAgileReference
    {
        private unsafe struct AgileReferenceVTable
        {
            // Due to a bug in .NET Core 3.1 (3.1.18 as of now) we must inline IUnknown interface
            // members explicitly, and can't do the following:
            //
            //      public IUnknownVTable UnknownVTable;

            public delegate* unmanaged[Stdcall]<IntPtr, Guid*, IntPtr*, HRESULT> QueryInterface;
            public delegate* unmanaged[Stdcall]<IntPtr, HRESULT> AddRef;
            public delegate* unmanaged[Stdcall]<IntPtr, HRESULT> Release;

            // IAgileReference
            public delegate* unmanaged[Stdcall]<IntPtr, ref Guid, out IntPtr, HRESULT> Resolve;
        }

        private readonly IntPtr instance;
        private readonly unsafe AgileReferenceVTable* vtable;

        public RCW_IAgileReference(IntPtr instance)
        {
            this.instance = instance;
            unsafe
            {
                this.vtable = *(AgileReferenceVTable**)instance;
            }
        }

        // An IAgileReference instance handles release on the correct context.
        ~RCW_IAgileReference()
        {
            if (this.instance != IntPtr.Zero)
            {
                Marshal.Release(this.instance);
            }
        }

        public IntPtr Resolve(Guid iid)
        {
            unsafe
            {
                // Marshal
                IntPtr resolvedInstance;

                // Dispatch
                this.vtable->Resolve(this.instance, ref iid, out resolvedInstance)
                    .ThrowIfFailed();

                // Unmarshal
                return resolvedInstance;
            }
        }
    }

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

    internal unsafe struct DataObjectVTable
    {
        public static readonly Guid IID_IDataObject = new Guid("0000010e-0000-0000-C000-000000000046");

        public IUnknownVTable UnknownVTable;

        // IDataObject
#if NET50_OR_GREATER
        public delegate* unmanaged[Stdcall]<IntPtr, FORMATETC*, Interops.STGMEDIUM*, int> GetData;
        public delegate* unmanaged[Stdcall]<IntPtr, FORMATETC*, Interops.STGMEDIUM*, int> GetDataHere;
        public delegate* unmanaged[Stdcall]<IntPtr, /*optional*/ FORMATETC*, int> QueryGetData;
        public delegate* unmanaged[Stdcall]<IntPtr, /*optional*/ FORMATETC*, FORMATETC*, int> GetCanonicalFormatEtc;
        public delegate* unmanaged[Stdcall]<IntPtr, FORMATETC*, Interops.STGMEDIUM*, int, int> SetData;
        public delegate* unmanaged[Stdcall]<IntPtr, int, IntPtr*, int> EnumFormatEtc;
        public delegate* unmanaged[Stdcall]<IntPtr, FORMATETC*, int, IntPtr, int*, int> DAdvise;
        public delegate* unmanaged[Stdcall]<IntPtr, int, int> DUnadvise;
        public delegate* unmanaged[Stdcall]<IntPtr, IntPtr*, int> EnumDAdvise;
#else
        public IntPtr GetData;
        public IntPtr GetDataHere;
        public IntPtr QueryGetData;
        public IntPtr GetCanonicalFormatEtc;
        public IntPtr SetData;
        public IntPtr EnumFormatEtc;
        public IntPtr DAdvise;
        public IntPtr DUnadvise;
        public IntPtr EnumDAdvise;

        public unsafe delegate HRESULT GetDataDelegate(IntPtr _this, FORMATETC* format, Interops.STGMEDIUM* medium);
        public unsafe delegate HRESULT GetDataHereDelegate(IntPtr _this, FORMATETC* format, Interops.STGMEDIUM* medium);
        public unsafe delegate HRESULT QueryGetDataDelegate(IntPtr _this, /*optional*/ FORMATETC* format);
        public unsafe delegate HRESULT GetCanonicalFormatEtcDelegate(IntPtr _this, /*optional*/ FORMATETC* formatIn, FORMATETC* formatOut);
        public unsafe delegate HRESULT SetDataDelegate(IntPtr _this, FORMATETC* format, Interops.STGMEDIUM* medium, int shouldRelease);
        public unsafe delegate HRESULT EnumFormatEtcDelegate(IntPtr _this, int direction, IntPtr* enumFORMATETC);
        public unsafe delegate HRESULT DAdviseDelegate(IntPtr _this, FORMATETC* format, int advf, IntPtr adviseSink, int* connection);
        public delegate HRESULT DUnadviseDelegate(IntPtr _this, int connection);
        public unsafe delegate HRESULT EnumDAdviseDelegate(IntPtr _this, IntPtr* enumSTATDATA);
#endif
    }

    internal class RCW_IDataObject : IDataObject
    {
        private readonly RCW_IAgileReference agileInstance;
        private readonly IntPtr instanceInSta;
        private readonly unsafe DataObjectVTable* vtableInSta;

        public RCW_IDataObject(RCW_IAgileReference agileReference)
        {
            // Use IAgileReference instance to always be in context.
            this.agileInstance = agileReference;

            Debug.Assert(Thread.CurrentThread.GetApartmentState() == ApartmentState.STA);

            // Assuming this class is always in context getting it once is possible.
            // See Finalizer for lifetime detail concerns. If the Clipboard instance
            // is considered a process singleton, then it could be leaked.
            (IntPtr instance, IntPtr vtable) = GetContextSafeRef(this.agileInstance);
            this.instanceInSta = instance;
            unsafe
            {
                this.vtableInSta = (DataObjectVTable*)vtable;
            }
        }

        // This Finalizer only works if the IDataObject is free threaded or if code
        // is added to ensure the Release takes place in the correct context.
        ~RCW_IDataObject()
        {
            // This should likely be some other mechanism, but the concept is correct.
            // For WinForms we need any STA Control since all of them possess a
            // BeginInvoke call. Alternatively this could be pass over to a
            // cleanup thread that asks the main STA to clean up instances.
            var formMaybe = System.Windows.Forms.Form.ActiveForm;
            if (formMaybe != null)
            {
                IntPtr instanceLocal = this.instanceInSta;
                if (instanceLocal != IntPtr.Zero)
                {
                    // Clean up on the main thread
                    formMaybe.BeginInvoke(new Action(() =>
                    {
                        Marshal.Release(instanceLocal);
                    }));
                }
            }
        }

        private unsafe static (IntPtr inst, IntPtr vtable) GetContextSafeRef(RCW_IAgileReference agileRef)
        {
            IntPtr instSafe = agileRef.Resolve(typeof(IDataObject).GUID);

            // Retain the instance's vtable when in context.
            unsafe
            {
                var vtableSafe = (IntPtr)(*(DataObjectVTable**)instSafe);
                return (instSafe, vtableSafe);
            }
        }

        public IntPtr GetInstanceForSta()
        {
            Debug.Assert(Thread.CurrentThread.GetApartmentState() == ApartmentState.STA);
            return this.instanceInSta;
        }

        public void GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            unsafe
            {
                // Marshal
                var stgmed = default(Interops.STGMEDIUM);
                medium = default;

                // Dispatch
                HRESULT hr;
                (IntPtr instance, IntPtr vtable) = GetContextSafeRef(this.agileInstance);
                fixed (FORMATETC* pFormat = &format)
                {
                    var getDataDelegate = Marshal.GetDelegateForFunctionPointer<DataObjectVTable.GetDataDelegate>(((DataObjectVTable*)vtable)->GetData);
                    hr = getDataDelegate.Invoke(instance, pFormat, &stgmed);
                }
                Marshal.Release(instance);

                hr.ThrowIfFailed();

                // Unmarshal
                medium.tymed = stgmed.tymed;
                medium.unionmember = stgmed.unionmember;
                if (stgmed.pUnkForRelease != IntPtr.Zero)
                {
                    medium.pUnkForRelease = Marshal.GetObjectForIUnknown(stgmed.pUnkForRelease);
                }
            }
        }

        public void GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            throw new NotImplementedException();
        }

        public unsafe int QueryGetData(ref FORMATETC format)
        {
            throw new NotImplementedException();
        }

        public int GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            throw new NotImplementedException();
        }

        public void SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            unsafe
            {
                // Marshal
                var pUnk = default(IntPtr);
                if (medium.pUnkForRelease is not null)
                {
                    pUnk = Marshal.GetIUnknownForObject(medium.pUnkForRelease);
                }

                var stgmed = new Interops.STGMEDIUM()
                {
                    unionmember = medium.unionmember,
                    tymed = medium.tymed,
                    pUnkForRelease = pUnk
                };

                int isRelease = release ? 1 : 0;

                // Dispatch
                HRESULT hr;
                (IntPtr instance, IntPtr vtable) = GetContextSafeRef(this.agileInstance);
                fixed (FORMATETC* pFormat = &formatIn)
                {
                    var setDataDelegate = Marshal.GetDelegateForFunctionPointer<DataObjectVTable.SetDataDelegate>(((DataObjectVTable*)vtable)->SetData);
                    hr = setDataDelegate.Invoke(instance, pFormat, &stgmed, isRelease);
                }

                Marshal.Release(instance);

                hr.ThrowIfFailed();
            }
        }

        public IEnumFORMATETC EnumFormatEtc(DATADIR direction)
        {
            throw new NotImplementedException();
        }

        public int DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
        {
            throw new NotImplementedException();
        }

        public void DUnadvise(int connection)
        {
            throw new NotImplementedException();
        }

        public int EnumDAdvise(out IEnumSTATDATA enumAdvise)
        {
            throw new NotImplementedException();
        }
    }

    internal static class CCW_IDataObject
    {
        private static List<Delegate> delegates = new List<Delegate>();
        private static Lazy<IntPtr> CCWVTable = new Lazy<IntPtr>(AllocateVTable, isThreadSafe: true);

        private unsafe struct Lifetime
        {
            public DataObjectVTable* VTable;
            public IntPtr Handle;
            public int RefCount;
        }

        public static IntPtr CreateInstance(Forms_IDataObject dataObject)
        {
            unsafe
            {
                var wrapper = (Lifetime*)Marshal.AllocCoTaskMem(sizeof(Lifetime));

                // Create the wrapper instance.
                wrapper->VTable = (DataObjectVTable*)CCWVTable.Value;
                wrapper->Handle = GCHandle.ToIntPtr(GCHandle.Alloc(dataObject));
                wrapper->RefCount = 1;

                return (IntPtr)wrapper;
            }
        }

        public static bool TryGetInstance(IntPtr instanceMaybe, [NotNullWhen(true)] out Forms_IDataObject? forms_dataObject)
        {
            forms_dataObject = null;

            unsafe
            {
                // This is a dangerous cast since it relies on strictly
                // following the COM ABI. If the VTable is ours the rest of
                // the structure is good, otherwise it is unknown.
                var lifetime = (Lifetime*)instanceMaybe;
                if (lifetime->VTable != (DataObjectVTable*)CCWVTable.Value)
                {
                    return false;
                }

                forms_dataObject = GetInstance(instanceMaybe);
                return true;
            }
        }

        private static IntPtr AllocateVTable()
        {
            unsafe
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

                delegates.Add(queryInterfaceDelegate);
                delegates.Add(addRefDelegate);
                delegates.Add(releaseDelegate);
                delegates.Add(getDataDelegate);
                delegates.Add(getDataHereDelegate);
                delegates.Add(queryGetDataDelegate);
                delegates.Add(getCanonicalFormatEtcDelegate);
                delegates.Add(setDataDelegate);
                delegates.Add(enumFormatEtcDelegate);
                delegates.Add(dAdviseDelegate);
                delegates.Add(dUnadviseDelegate);
                delegates.Add(enumDAdviseDelegate);

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
        }

        private static Forms_IDataObject GetInstance(IntPtr wrapper)
        {
            unsafe
            {
                var lifetime = (Lifetime*)wrapper;

                Debug.Assert(lifetime->Handle != IntPtr.Zero);
                Debug.Assert(lifetime->RefCount > 0);
                return (Forms_IDataObject)GCHandle.FromIntPtr(lifetime->Handle).Target!;
            }
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

        private static uint AddRef(IntPtr _this)
        {
            unsafe
            {
                var lifetime = (Lifetime*)_this;
                Debug.Assert(lifetime->Handle != IntPtr.Zero);
                Debug.Assert(lifetime->RefCount > 0);
                return (uint)Interlocked.Increment(ref lifetime->RefCount);
            }
        }

        private static uint Release(IntPtr _this)
        {
            unsafe
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
            formatOut = default;
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
