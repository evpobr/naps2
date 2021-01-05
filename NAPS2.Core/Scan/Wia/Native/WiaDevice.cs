using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace NAPS2.Scan.Wia.Native
{
    public class WiaDevice : WiaItemBase, IWiaDeviceProps
    {
        protected internal WiaDevice(WiaVersion version, IntPtr handle) : base(version, handle)
        {
        }
        
        public WiaItem PromptToConfigure(IntPtr parentWindowHandle)
        {
            if (Version == WiaVersion.Wia20)
            {
                throw new InvalidOperationException("WIA 2.0 does not support PromptToConfigure. Use WiaDeviceManager.PromptForImage if you want to use the native WIA 2.0 UI.");
            }

            var device = (IWiaItem)Marshal.GetObjectForIUnknown(Handle);
            var hr = NativeWiaMethods.ConfigureDevice1(device, parentWindowHandle, 0, 0, out var itemCount, out var items);
            if (hr == 1)
            {
                return null;
            }
            WiaException.Check(hr);
            return new WiaItem(Version, Marshal.GetIUnknownForObject(items[0]));
        }
    }
}
