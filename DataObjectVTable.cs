using System;
using System.Runtime.InteropServices.ComTypes;

namespace ClipboardRedux
{
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
}
