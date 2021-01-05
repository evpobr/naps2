using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace NAPS2.Scan.Wia.Native
{
    public class WiaDeviceManager : NativeWiaObject
    {
        private const int SCANNER_DEVICE_TYPE = 1;
        private const int SELECT_DEVICE_NODEFAULT = 1;

        public WiaDeviceManager() : base(WiaVersion.Default)
        {
        }

        public WiaDeviceManager(WiaVersion version) : base(version)
        {
            var handle = IntPtr.Zero;
            if (Version == WiaVersion.Wia10)
            {
                WiaException.Check(NativeWiaMethods.GetDeviceManager1(out var deviceManager));
                handle = Marshal.GetIUnknownForObject(deviceManager);
            }
            else
            {
                WiaException.Check(NativeWiaMethods.GetDeviceManager2(out var deviceManager));
                handle = Marshal.GetIUnknownForObject(deviceManager);
            }
            Handle = handle;
        }

        protected internal WiaDeviceManager(WiaVersion version, IntPtr handle) : base(version, handle)
        {
        }

        public IEnumerable<WiaDeviceInfo> GetDeviceInfos()
        {
            List<WiaDeviceInfo> result = new List<WiaDeviceInfo>();
            if (Version == WiaVersion.Wia10)
            {
                var deviceManager = (IWiaDevMgr)Marshal.GetObjectForIUnknown(Handle);
                WiaException.Check(NativeWiaMethods.EnumerateDevices1(deviceManager, x => result.Add(new WiaDeviceInfo(Version, x))));
            }
            else
            {
                var deviceManager = (IWiaDevMgr2)Marshal.GetObjectForIUnknown(Handle);
                WiaException.Check(NativeWiaMethods.EnumerateDevices2(deviceManager, x => result.Add(new WiaDeviceInfo(Version, x))));
            }
            return result;
        }

        public WiaDevice FindDevice(string deviceID)
        {
            IntPtr deviceHandle;
            if (Version == WiaVersion.Wia10)
            {
                var deviceManager = (IWiaDevMgr)Marshal.GetObjectForIUnknown(Handle);
                WiaException.Check(NativeWiaMethods.GetDevice1(deviceManager, deviceID, out var device));
                deviceHandle = Marshal.GetIUnknownForObject(device);
            }
            else
            {
                var deviceManager = (IWiaDevMgr2)Marshal.GetObjectForIUnknown(Handle);
                WiaException.Check(NativeWiaMethods.GetDevice2(deviceManager, deviceID, out var device));
                deviceHandle = Marshal.GetIUnknownForObject(device);
            }
            return new WiaDevice(Version, deviceHandle);
        }

        public WiaDevice PromptForDevice(IntPtr parentWindowHandle)
        {
            uint hr;
            IntPtr deviceHandle;
            if (Version == WiaVersion.Wia10)
            {
                var deviceManager = (IWiaDevMgr)Marshal.GetObjectForIUnknown(Handle);
                hr = NativeWiaMethods.SelectDevice1(deviceManager, parentWindowHandle, SCANNER_DEVICE_TYPE, SELECT_DEVICE_NODEFAULT, out _, out var device);
                deviceHandle = Marshal.GetIUnknownForObject(device);
            }
            else
            {
                var deviceManager = (IWiaDevMgr2)Marshal.GetObjectForIUnknown(Handle);
                hr = NativeWiaMethods.SelectDevice2(deviceManager, parentWindowHandle, SCANNER_DEVICE_TYPE, SELECT_DEVICE_NODEFAULT, out _, out var device);
                deviceHandle = Marshal.GetIUnknownForObject(device);
            }
            if (hr == 1)
            {
                return null;
            }
            WiaException.Check(hr);
            return new WiaDevice(Version, deviceHandle); ;
        }

        public string[] PromptForImage(IntPtr parentWindowHandle, WiaDevice device)
        {
            var fileName = Path.GetRandomFileName();
            IntPtr itemHandle = IntPtr.Zero;
            int fileCount = 0;
            string[] filePaths = new string[10];
            uint hr = 0;
            if (Version == WiaVersion.Wia10)
            {
                var deviceManager = (IWiaDevMgr)Marshal.GetObjectForIUnknown(Handle);
                hr = NativeWiaMethods.GetImage1(deviceManager, parentWindowHandle, SCANNER_DEVICE_TYPE, 0, 0, Path.Combine(Paths.Temp, fileName), null);
            }
            else
            {
                var deviceManager = (IWiaDevMgr2)Marshal.GetObjectForIUnknown(Handle);
                IWiaItem2 item = null;
                hr = NativeWiaMethods.GetImage2(deviceManager, 0, device.Id(), parentWindowHandle, Paths.Temp, fileName, ref fileCount, ref filePaths, ref item);
            }
            if (hr == 1)
            {
                return null;
            }
            WiaException.Check(hr);
            return filePaths ?? new[] { Path.Combine(Paths.Temp, fileName) };
        }
    }
}
