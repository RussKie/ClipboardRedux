using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using WindowsFormsClipboardRedux;

namespace app15
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var format = DataFormats.UnicodeText;

            //Clipboard.SetDataObject("Hello world!");
            // If we don't do this - it can fail with CLIPBRD_E_CANT_OPEN
            //var dataObject = Clipboard.GetDataObject();



            // Fails with: 'System.InvalidCastException: Specified cast is not valid.'
            var dataObject1 = GetDataObject();
            if (dataObject1 is not null)
            {
                if (dataObject1.GetDataPresent(DataFormats.Bitmap))
                {
                    var data = dataObject1.GetData(DataFormats.Bitmap);
                    if (data is Bitmap bitmap)
                    {
                        bitmap.Save(@"C:\Users\igveliko\Desktop\b.bmp");
                        bitmap.Dispose();
                    }
                    else
                    {
                        Debug.WriteLine("Not a bitmap");
                    }
                }
                else if (dataObject1.GetDataPresent(DataFormats.Text))
                {
                    Debug.WriteLine(dataObject1.GetData(DataFormats.Text));
                }
            }
            else
            {
                Debug.WriteLine("opps");
            }
        }

        private static IDataObject? GetDataObject()
        {
            object dataObject = ClipboardImpl.Get();
            if (dataObject is not null)
            {
                if (dataObject is IDataObject ido)
                {
                    return ido;
                }

                return new DataObject(dataObject);
            }

            return null;
        }
    }
}

