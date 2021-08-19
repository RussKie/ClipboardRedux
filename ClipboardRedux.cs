using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using IFormsDataObject = System.Windows.Forms.IDataObject;
using FormsDataObject = System.Windows.Forms.DataObject;

namespace WindowsFormsClipboardRedux
{
    public static partial class ClipboardRedux
    {
        public static IFormsDataObject? GetDataObject()
        {
            ClipboardImpl.Initialize();

            object dataObject = ClipboardImpl.Get();
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
    }
}
