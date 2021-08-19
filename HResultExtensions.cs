using System;
using System.Runtime.InteropServices;

namespace WindowsFormsClipboardRedux
{
    internal static class HResultExtensions
    {
        public static bool Succeeded(this HRESULT hr) => hr >= 0;

        public static bool Failed(this HRESULT hr) => hr < 0;

        public static void ThrowIfFailed(this HRESULT hr)
        {
            if (hr.Failed())
            {
                Marshal.ThrowExceptionForHR((int)hr);
            }
        }

        public static string AsString(this HRESULT hr)
            => Enum.IsDefined(typeof(HRESULT), hr)
                ? $"HRESULT {hr} [0x{(int)hr:X} ({(int)hr:D})]"
                : $"HRESULT [0x{(int)hr:X} ({(int)hr:D})]";
    }

}
