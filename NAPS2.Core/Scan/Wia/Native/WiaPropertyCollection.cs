using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NAPS2.Util;

namespace NAPS2.Scan.Wia.Native
{
    public class WiaPropertyCollection : NativeWiaObject, IEnumerable<WiaProperty>
    {
        private readonly Dictionary<int, WiaProperty> propertyDict;

        public WiaPropertyCollection(WiaVersion version, IWiaPropertyStorage propertyStorageHandle) : base(version, Marshal.GetIUnknownForObject(propertyStorageHandle))
        {
            propertyDict = new Dictionary<int, WiaProperty>();
            WiaException.Check(NativeWiaMethods.EnumerateProperties((IWiaPropertyStorage)Marshal.GetObjectForIUnknown(Handle),
                (id, name, type) => propertyDict.Add(id, new WiaProperty((IWiaPropertyStorage)Marshal.GetObjectForIUnknown(Handle), id, name, type))));
        }
        
        public WiaProperty this[int propId] => propertyDict.Get(propId);

        public IEnumerator<WiaProperty> GetEnumerator() => propertyDict.Values.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
