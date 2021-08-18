using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using Forms_IDataObject = System.Windows.Forms.IDataObject;

// Clean up warnings
#pragma warning disable CS0649
#pragma warning disable CA1416

namespace ClipboardRedux
{
    internal static class ClipboardImpl
    {
        public static void Initialize()
        {
            int hr = Ole32.OleInitialize(IntPtr.Zero);
            // S_FALSE is a valid HRESULT
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }
        }

        public static object Get()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                throw new InvalidOperationException();
            }

            IntPtr instance;
            int hr = Ole32.OleGetClipboard(out instance);
            if (hr != 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }

            // Check if the returned value is actually a wrapped managed instance.
            if (CCW_IDataObject.TryGetInstance(instance, out Forms_IDataObject? forms_dataObject))
            {
                return forms_dataObject;
            }

            IntPtr agileReference;
            var iid = typeof(IDataObject).GUID;
            hr = Ole32.RoGetAgileReference(
                Ole32.AgileReferenceOptions.Default,
                ref iid,
                instance,
                out agileReference);
            if (hr != 0)
            {
                // Release the clipboard object if agile
                // reference creation failed.
                Marshal.Release(instance);
                Marshal.ThrowExceptionForHR(hr);
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

            int hr;
            if (obj is null)
            {
                hr = Ole32.OleSetClipboard(IntPtr.Zero);
                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                return;
            }

            if (obj is RCW_IDataObject com_dataobject)
            {
                hr = Ole32.OleSetClipboard(com_dataobject.GetInstanceForSta());
            }
            else if (obj is Forms_IDataObject forms_dataObject)
            {
                // This approach is less than ideal since a new wrapper is always
                // created. Having an efficient cache would be more effective.
                var ccw = CCW_IDataObject.CreateInstance(forms_dataObject);
                hr = Ole32.OleSetClipboard(ccw);
            }
            else
            {
                // This requires implementing a universal CCW or alternatively
                // leveraging the built-in system. It isn't obvious which one is
                // the best option - both are possible.
                throw new NotImplementedException();
            }

            if (hr != 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }
        }

        public static int Flush()
        {
            return Ole32.OleFlushClipboard();
        }

        private static class Ole32
        {
            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern int OleInitialize(IntPtr reserved);

            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern int OleGetClipboard(out IntPtr dataObject);

            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern int OleSetClipboard(IntPtr dataObject);

            [DllImport(nameof(Ole32), ExactSpelling = true)]
            public static extern int OleFlushClipboard();

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
            public static extern int RoGetAgileReference(
                AgileReferenceOptions opts,
                ref Guid riid,
                IntPtr instance,
                out IntPtr agileReference);
        }
    }

    internal class RCW_IAgileReference
    {
        private unsafe struct AgileReferenceVTable
        {
            public IUnknownVTable UnknownVTable;

            // IAgileReference
            public delegate* unmanaged[Stdcall]<IntPtr, ref Guid, out IntPtr, int> Resolve;
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
                int hr = this.vtable->Resolve(this.instance, ref iid, out resolvedInstance);
                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                // Unmarshal
                return resolvedInstance;
            }
        }
    }

    internal struct STGMEDIUM_Blittable
    {
        public TYMED tymed;
        public IntPtr unionmember;
        public IntPtr pUnkForRelease;
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
#endif
    }

    internal unsafe struct DataObjectVTable
    {
        public static readonly Guid IID_IDataObject = new Guid("0000010e-0000-0000-C000-000000000046");

        public IUnknownVTable UnknownVTable;

        // IDataObject
#if NET50_OR_GREATER
        public delegate* unmanaged[Stdcall]<IntPtr, FORMATETC*, STGMEDIUM_Blittable*, int> GetData;
        public delegate* unmanaged[Stdcall]<IntPtr, FORMATETC*, STGMEDIUM_Blittable*, int> GetDataHere;
        public delegate* unmanaged[Stdcall]<IntPtr, /*optional*/ FORMATETC*, int> QueryGetData;
        public delegate* unmanaged[Stdcall]<IntPtr, /*optional*/ FORMATETC*, FORMATETC*, int> GetCanonicalFormatEtc;
        public delegate* unmanaged[Stdcall]<IntPtr, FORMATETC*, STGMEDIUM_Blittable*, int, int> SetData;
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
                var stgmed = default(STGMEDIUM_Blittable);
                medium = default;

                // Dispatch
                int hr;
                (IntPtr instance, IntPtr vtable) = GetContextSafeRef(this.agileInstance);
                fixed (FORMATETC* formatFixed = &format)
                {
                    hr = ((DataObjectVTable*)vtable)->GetData(instance, formatFixed, &stgmed);
                }
                Marshal.Release(instance);

                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

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

        public int QueryGetData(ref FORMATETC format)
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

                var stgmed = new STGMEDIUM_Blittable()
                {
                    unionmember = medium.unionmember,
                    tymed = medium.tymed,
                    pUnkForRelease = pUnk
                };

                int isRelease = release ? 1 : 0;

                // Dispatch
                int hr;
                (IntPtr instance, IntPtr vtable) = GetContextSafeRef(this.agileInstance);
                fixed (FORMATETC* formatFixed = &formatIn)
                {
                    hr = ((DataObjectVTable*)vtable)->SetData(instance, formatFixed, &stgmed, isRelease);
                }

                Marshal.Release(instance);

                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }
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
        private const int E_NOTIMPL = unchecked((int)0x80004001);

        // IUnknown
        private unsafe delegate int QueryInterfaceDelegate(IntPtr _this, Guid* iid, IntPtr* obj);
        private delegate uint AddRefDelegate(IntPtr _this);
        private delegate uint ReleaseDelegate(IntPtr _this);
        // IDataObject
        private unsafe delegate int GetDataDelegate(IntPtr _this, FORMATETC* format, STGMEDIUM_Blittable* medium);
        private unsafe delegate int GetDataHereDelegate(IntPtr _this, FORMATETC* format, STGMEDIUM_Blittable* medium);
        private unsafe delegate int QueryGetDataDelegate(IntPtr _this, /*optional*/ FORMATETC* format);
        private unsafe delegate int GetCanonicalFormatEtcDelegate(IntPtr _this, /*optional*/ FORMATETC* formatIn, FORMATETC* formatOut);
        private unsafe delegate int SetDataDelegate(IntPtr _this, FORMATETC* format, STGMEDIUM_Blittable* medium, int shouldRelease);
        private unsafe delegate int EnumFormatEtcDelegate(IntPtr _this, int direction, IntPtr* enumFORMATETC);
        private unsafe delegate int DAdviseDelegate(IntPtr _this, FORMATETC* format, int advf, IntPtr adviseSink, int* connection);
        private delegate int DUnadviseDelegate(IntPtr _this, int connection);
        private unsafe delegate int EnumDAdviseDelegate(IntPtr _this, IntPtr* enumSTATDATA);

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

                var queryInterfaceDelegate = new QueryInterfaceDelegate(QueryInterface);
                var addRefDelegate = new AddRefDelegate(AddRef);
                var releaseDelegate = new ReleaseDelegate(Release);
                var getDataDelegate = new GetDataDelegate(GetData);
                var getDataHereDelegate = new GetDataHereDelegate(GetDataHere);
                var queryGetDataDelegate = new QueryGetDataDelegate(QueryGetData);
                var getCanonicalFormatEtcDelegate = new GetCanonicalFormatEtcDelegate(GetCanonicalFormatEtc);
                var setDataDelegate = new SetDataDelegate(SetData);
                var enumFormatEtcDelegate = new EnumFormatEtcDelegate(EnumFormatEtc);
                var dAdviseDelegate = new DAdviseDelegate(DAdvise);
                var dUnadviseDelegate = new DUnadviseDelegate(DUnadvise);
                var enumDAdviseDelegate = new EnumDAdviseDelegate(EnumDAdvise);

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
                return (Forms_IDataObject)GCHandle.FromIntPtr(lifetime->Handle).Target;
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
        private unsafe static int GetData(IntPtr _this, FORMATETC* format, STGMEDIUM_Blittable* medium)
        {
            *medium = default;
            return E_NOTIMPL;
        }

        private unsafe static int GetDataHere(IntPtr _this, FORMATETC* format, STGMEDIUM_Blittable* medium)
        {
            return E_NOTIMPL;
        }

        private unsafe static int QueryGetData(IntPtr _this, /*optional*/ FORMATETC* format)
        {
            return E_NOTIMPL;
        }

        private unsafe static int GetCanonicalFormatEtc(IntPtr _this, /*optional*/ FORMATETC* formatIn, FORMATETC* formatOut)
        {
            formatOut = default;
            return E_NOTIMPL;
        }

        private unsafe static int SetData(IntPtr _this, FORMATETC* format, STGMEDIUM_Blittable* medium, int shouldRelease)
        {
            return E_NOTIMPL;
        }

        private unsafe static int EnumFormatEtc(IntPtr _this, int direction, IntPtr* enumFORMATETC)
        {
            *enumFORMATETC = default;
            return E_NOTIMPL;
        }

        private unsafe static int DAdvise(IntPtr _this, FORMATETC* format, int advf, IntPtr adviseSink, int* connection)
        {
            *connection = default;
            return E_NOTIMPL;
        }

        private static int DUnadvise(IntPtr _this, int connection)
        {
            return E_NOTIMPL;
        }

        private unsafe static int EnumDAdvise(IntPtr _this, IntPtr* enumSTATDATA)
        {
            *enumSTATDATA = default;
            return E_NOTIMPL;
        }
    }
}
