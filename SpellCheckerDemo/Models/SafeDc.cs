using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpellCheckerDemo.Models
{
    public class SafeDC : IDisposable
    {
        public IntPtr HDC { get; }
        private readonly IntPtr _handle;

        public SafeDC(IntPtr handle)
        {
            _handle = handle;
            HDC = NativeMethods.GetDC(handle);
        }

        public void Dispose() => NativeMethods.ReleaseDC(_handle , HDC);
    }
}
