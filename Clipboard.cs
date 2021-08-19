using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Forms_IDataObject = System.Windows.Forms.IDataObject;

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
}
