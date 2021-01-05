using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace NAPS2.Scan.Wia.Native
{
    public class WiaItemBase : NativeWiaObject
    {
        private WiaPropertyCollection properties;

        protected internal WiaItemBase(WiaVersion version, IntPtr handle) : base(version, handle)
        {
        }

        public WiaPropertyCollection Properties
        {
            get
            {
                if (properties == null)
                {
                    WiaException.Check(NativeWiaMethods.GetItemPropertyStorage(Marshal.GetObjectForIUnknown(Handle), out var propStorage));
                    properties = new WiaPropertyCollection(Version, propStorage);
                }
                return properties;
            }
        }

        public List<WiaItem> GetSubItems()
        {
            var items = new List<WiaItem>();
            WiaException.Check(Version == WiaVersion.Wia10
                ? NativeWiaMethods.EnumerateItems1((IWiaItem)Marshal.GetObjectForIUnknown(Handle), itemHandle => items.Add(new WiaItem(Version, Marshal.GetIUnknownForObject(itemHandle))))
                : NativeWiaMethods.EnumerateItems2((IWiaItem2)Marshal.GetObjectForIUnknown(Handle), itemHandle => items.Add(new WiaItem(Version, Marshal.GetIUnknownForObject(itemHandle)))));
            return items;
        }

        public WiaItem FindSubItem(string name)
        {
            return GetSubItems().FirstOrDefault(x => x.Name() == name);
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                properties?.Dispose();
            }
        }
    }
}