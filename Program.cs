using System;
using System.Windows.Forms;
using ClipboardRedux;

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

            Clipboard.SetDataObject("Hello world!");
            // If we don't do this - it can fail with CLIPBRD_E_CANT_OPEN
            //var dataObject = Clipboard.GetDataObject();


            ClipboardImpl.Initialize();

            // Fails with: 'System.InvalidCastException: Specified cast is not valid.'
            var dataObject1 = GetDataObject();
            if (dataObject1 is not null)
            {
                Console.WriteLine(dataObject1.GetData(format));
            }

            Console.WriteLine("opps");
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

