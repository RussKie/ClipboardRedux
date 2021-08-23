using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
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
            //Clipboard.SetDataObject("Hello world!");
            // If we don't do this - it can fail with CLIPBRD_E_CANT_OPEN
            //var dataObject = Clipboard.GetDataObject();

            // Fails with: 'System.InvalidCastException: Specified cast is not valid.'

            Console.WriteLine("Press ESC to stop");
            do
            {
                while (!Console.KeyAvailable)
                {
                    //var dataObject1 = Clipboard.GetDataObject();
                    var dataObject1 = ClipboardRedux.GetDataObject();
                    if (dataObject1 is not null)
                    {
                        if (dataObject1.GetDataPresent(DataFormats.Bitmap))
                        {
                            var data = dataObject1.GetData(DataFormats.Bitmap);
                            if (data is Bitmap bitmap)
                            {
                                //string path = Path.GetFullPath("clipboard.bmp");
                                //bitmap.Save(path);
                                //bitmap.Dispose();
                                //Debug.WriteLine($"Bitmap saved to {path}");
                                Console.WriteLine("Is a bitmap");
                            }
                            else
                            {
                                Debug.WriteLine("Not a bitmap");
                            }
                        }
                        else if (dataObject1.GetDataPresent(DataFormats.Text))
                        {
                            Console.WriteLine($"Text: {dataObject1.GetData(DataFormats.Text)}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("opps");
                    }
                }
            } while (Console.ReadKey(true).Key != ConsoleKey.Escape);
        }
    }
}

