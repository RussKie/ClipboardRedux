using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using FormsDataObject = System.Windows.Forms.DataObject;
using IComDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using IFormsDataObject = System.Windows.Forms.IDataObject;

namespace WindowsFormsClipboardRedux
{
    public static partial class ClipboardRedux
    {
        public static IFormsDataObject? GetDataObject()
        {
            if (Application.OleRequired() != ApartmentState.STA)
            {
                // Only throw if a message loop was started. This makes the case of trying
                // to query the clipboard from your finalizer or non-ui MTA thread
                // silently fail, instead of making your app die.
                //
                // however, if you are trying to write a normal windows forms app and
                // forget to set the STAThread attribute, we will correctly report
                // an error to aid in debugging.
                if (Application.MessageLoop)
                {
                    throw new ThreadStateException("SR.ThreadMustBeSTA");
                }

                return null;
            }

            // We need to retry the GetDataObject() since the clipBoard is busy sometimes and
            // hence the GetDataObject would fail with ClipBoardException.
            object dataObject = GetDataObject(retryTimes: 10, retryDelay: 100);
            if (dataObject is not null)
            {
                if (dataObject is IFormsDataObject ido)
                {
                    return ido;
                }

                return new FormsDataObject(dataObject);
            }

            return null;
        }

        private static object GetDataObject(int retryTimes, int retryDelay)
        {
            IntPtr instance;
            HRESULT hr;
            int retry = retryTimes;
            do
            {
                hr = Interops.Ole32.OleGetClipboard(out instance);
                if (hr != HRESULT.S_OK)
                {
                    if (retry == 0)
                    {
                        hr.ThrowIfFailed();
                    }

                    retry--;
                    Thread.Sleep(millisecondsTimeout: retryDelay);
                }
            }
            while (hr != HRESULT.S_OK);

            // Check if the returned value is actually a wrapped managed instance.
            if (CCW_IDataObject.TryGetInstance(instance, out IFormsDataObject? forms_dataObject))
            {
                return forms_dataObject;
            }

            var iid = typeof(IComDataObject).GUID;
            hr = Interops.Ole32.RoGetAgileReference(
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
    }
}
