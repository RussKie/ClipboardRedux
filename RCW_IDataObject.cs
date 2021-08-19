using System;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Forms_DataFormats = System.Windows.Forms.DataFormats;

namespace ClipboardRedux
{
    internal class RCW_IDataObject : IDataObject
    {
        private static readonly TYMED[] ALLOWED_TYMEDS = new[]
        {
            TYMED.TYMED_HGLOBAL,
            TYMED.TYMED_ISTREAM,
            TYMED.TYMED_GDI
        };

        private readonly RCW_IAgileReference _agileInstance;
        private readonly IntPtr _instanceInSta;
        private readonly unsafe DataObjectVTable* _vtableInSta;

        public RCW_IDataObject(RCW_IAgileReference agileReference)
        {
            // Use IAgileReference instance to always be in context.
            _agileInstance = agileReference;

            Debug.Assert(Thread.CurrentThread.GetApartmentState() == ApartmentState.STA);

            // Assuming this class is always in context getting it once is possible.
            // See Finalizer for lifetime detail concerns. If the Clipboard instance
            // is considered a process singleton, then it could be leaked.
            (IntPtr instance, IntPtr vtable) = GetContextSafeRef(_agileInstance);
            _instanceInSta = instance;
            unsafe
            {
                _vtableInSta = (DataObjectVTable*)vtable;
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
            if (formMaybe is not null)
            {
                IntPtr instanceLocal = _instanceInSta;
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
            var vtableSafe = (IntPtr)(*(DataObjectVTable**)instSafe);
            return (instSafe, vtableSafe);
        }

        public IntPtr GetInstanceForSta()
        {
            Debug.Assert(Thread.CurrentThread.GetApartmentState() == ApartmentState.STA);
            return _instanceInSta;
        }

        public unsafe void GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            // Marshal
            var stgmed = default(Interops.STGMEDIUM);
            medium = default;

            // Dispatch
            HRESULT hr;
            (IntPtr instance, IntPtr vtable) = GetContextSafeRef(_agileInstance);
            fixed (FORMATETC* pFormat = &format)
            {
                var @delegate = Marshal.GetDelegateForFunctionPointer<DataObjectVTable.GetDataDelegate>(((DataObjectVTable*)vtable)->GetData);
                hr = @delegate.Invoke(instance, pFormat, &stgmed);
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

        public void GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            throw new NotImplementedException();
        }

        public unsafe int QueryGetData(ref FORMATETC format)
        {
            // Dispatch
            HRESULT hr;
            (IntPtr instance, IntPtr vtable) = GetContextSafeRef(_agileInstance);
            fixed (FORMATETC* pFormat = &format)
            {
                var @delegate = Marshal.GetDelegateForFunctionPointer<DataObjectVTable.QueryGetDataDelegate>(((DataObjectVTable*)vtable)->QueryGetData);
                hr = @delegate.Invoke(instance, pFormat);
            }
            Marshal.Release(instance);
            if (hr.Failed())
            {
                return (int)HRESULT.S_FALSE;
            }

            if (format.dwAspect != DVASPECT.DVASPECT_CONTENT)
            {
                return (int)HRESULT.DV_E_DVASPECT;
            }

            if (!GetTymedUseable(format.tymed))
            {
                return (int)HRESULT.DV_E_TYMED;
            }

            if (format.cfFormat == 0)
            {
                return (int)HRESULT.S_FALSE;
            }

            Forms_DataFormats.Format dataFormat = Forms_DataFormats.GetFormat(format.cfFormat);
            if (dataFormat.Id != format.cfFormat)
            {
                format.cfFormat = unchecked((short)(ushort)dataFormat.Id);
                if (QueryGetData(ref format) != (int)HRESULT.S_OK)
                {
                    return (int)HRESULT.DV_E_FORMATETC;
                }
            }

            return (int)HRESULT.S_OK;
        }

        public int GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        ///  Returns true if the tymed is useable.
        /// </summary>
        private bool GetTymedUseable(TYMED tymed)
        {
            for (int i = 0; i < ALLOWED_TYMEDS.Length; i++)
            {
                if ((tymed & ALLOWED_TYMEDS[i]) != 0)
                {
                    return true;
                }
            }
            return false;
        }

        public unsafe void SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
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
            (IntPtr instance, IntPtr vtable) = GetContextSafeRef(_agileInstance);
            fixed (FORMATETC* pFormat = &formatIn)
            {
                var setDataDelegate = Marshal.GetDelegateForFunctionPointer<DataObjectVTable.SetDataDelegate>(((DataObjectVTable*)vtable)->SetData);
                hr = setDataDelegate.Invoke(instance, pFormat, &stgmed, isRelease);
            }

            Marshal.Release(instance);

            hr.ThrowIfFailed();
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
}
