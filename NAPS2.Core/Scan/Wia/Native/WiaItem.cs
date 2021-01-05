using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace NAPS2.Scan.Wia.Native
{
    public class WiaItem : WiaItemBase
    {
        protected internal WiaItem(WiaVersion version, IntPtr handle) : base(version, handle)
        {
        }
        
        public WiaTransfer StartTransfer()
        {
            IntPtr transferHandle = IntPtr.Zero;
            if (Version == WiaVersion.Wia10)
            {
                WiaException.Check(NativeWiaMethods.StartTransfer1((IWiaItem)Marshal.GetObjectForIUnknown(Handle), out var transfer));
                transferHandle = Marshal.GetIUnknownForObject(transfer);
            }
            else
            {
                WiaException.Check(NativeWiaMethods.StartTransfer2((IWiaItem2)Marshal.GetObjectForIUnknown(Handle), out var transfer));
                transferHandle = Marshal.GetIUnknownForObject(transfer);
            }
            return new WiaTransfer(Version, transferHandle);
        }
    }
}