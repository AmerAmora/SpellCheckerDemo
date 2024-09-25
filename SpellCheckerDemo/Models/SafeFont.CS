using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpellCheckerDemo.Models
{
    public class SafeFont : IDisposable
    {
        private readonly IntPtr _hdc;
        private readonly IntPtr _hFont;
        private readonly IntPtr _oldFont;

        public SafeFont(IntPtr handle)
        {
            _hdc = NativeMethods.GetDC(handle);
            _hFont = NativeMethods.SendMessage(handle , NativeMethods.WM_GETFONT , IntPtr.Zero , IntPtr.Zero);
            _oldFont = NativeMethods.SelectObject(_hdc , _hFont);
        }

        public void Dispose()
        {
            NativeMethods.SelectObject(_hdc , _oldFont);
            NativeMethods.ReleaseDC(IntPtr.Zero , _hdc);
        }
    }
}
