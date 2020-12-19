using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NAPS2.Platform;
using NAPS2.Util;

namespace NAPS2.Scan.Wia.Native
{
    [ComImport, Guid("6CBA0075-1287-407d-9B77-CF0E030435CC")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaItem2
    {
        IWiaItem2 CreateChildItem(int lItemFlags, int lCreationFlags, in string bstrItemName);
        void DeleteItem(int lFlags);
        IEnumWiaItem2 EnumChildItems(in Guid pCategoryGUID);
        IWiaItem2 FindItemByName(int lFlags, in string bstrFullItemName);
        Guid GetItemCategory();
        int GetItemType();

        [return: MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.Interface)]
        IWiaItem2[] DeviceDlg(int lFlags, IntPtr hwndParent, in string bstrFolderName, in string bstrFilename, out int plNumFiles, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)] out string[] ppbstrFilePaths);
        void DeviceCommand(int lFlags, in Guid pCmdGUID, ref IWiaItem2 ppIWiaItem2);
        IEnumWIA_DEV_CAPS EnumDeviceCapabilities(int lFlags);
        [return: MarshalAs(UnmanagedType.Bool)]
        bool CheckExtension(int lFlags, in string bstrName, in Guid riidExtensionInterface);
        [PreserveSig]
        int GetExtension(int lFlags, in string bstrName, in Guid riidExtensionInterface, [MarshalAs(UnmanagedType.Interface, IidParameterIndex = 3)] IntPtr ppOut);
        IWiaItem2 GetParentItem();
        IWiaItem2 GetRootItem();
        IWiaPreview GetPreviewComponent(int lFlags);
        IEnumWIA_DEV_CAPS EnumRegisterEventInfo(int lFlags, in Guid pEventGUID);
        [PreserveSig]
        int Diagnostic(int ulSize, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] out byte[] pBuffer);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WIA_DEV_CAP
    {
        public Guid guid;
        public int ulFlags;
        [MarshalAs(UnmanagedType.BStr)] public string bstrName;
        [MarshalAs(UnmanagedType.BStr)] public string bstrDescription;
        [MarshalAs(UnmanagedType.BStr)] public string bstrIcon;
        [MarshalAs(UnmanagedType.BStr)] public string bstrCommandline;
    }

    [ComImport, Guid("1fcc4287-aca6-11d2-a093-00c04f72dc3c")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumWIA_DEV_CAPS
    {
        [PreserveSig]
        int Next(int celt, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] out WIA_DEV_CAP[] rgelt, out int pceltFetched);
        int Skip();
        void Reset();
        IEnumWIA_DEV_CAPS Clone();
        int GetCount();
    }

    [ComImport, Guid("95C2B4FD-33F2-4d86-AD40-9431F0DF08F7")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaPreview
    {
        void GetNewPreview(int lFlags, IWiaItem2 pWiaItem2, IWiaTransferCallback pWiaTransferCallback);
        void UpdatePreview(int lFlags, IWiaItem2 pChildWiaItem2, IWiaTransferCallback pWiaTransferCallback);
        void DetectRegions(int lFlags);
        void Clear();
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WiaTransferParams
    {
        public int lMessage;
        public int lPercentComplete;
        public long ulTransferredBytes;
        [MarshalAs(UnmanagedType.Error)]
        public int hrErrorStatus;
    }
    
    [ComImport, Guid("27d4eaaf-28a6-4ca5-9aab-e678168b9527")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaTransferCallback
    {
        void TransferCallback(int lFlags, in WiaTransferParams pWiaTransferParams);
        IStream GetNextStream(int lFlags, string bstrItemName, string bstrFullItemName);
    };

    [ComImport, Guid("59970AF4-CD0D-44d9-AB24-52295630E582")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumWiaItem2
    {
        [PreserveSig]
        int Next(int cElt, out IWiaItem2 ppIWiaItem2, out int pcEltFetched);
        void Skip(int cElt);
        void Reset();
        IEnumWiaItem2 Clone();
        int GetCount();
    }

    [ComImport, Guid("ae6287b0-0084-11d2-973b-00a0c9068f2e")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaEventCallback
    {
        void ImageEventCallback(
            in Guid pEventGUID,
            string bstrEventDescription,
            string bstrDeviceID,
            string bstrDeviceDescription,
            int dwDeviceType,
            string bstrFullItemName,
            ref int pulEventType,
            int ulReserved
        );
    }

    [ComImport, Guid("5e38b83c-8cf1-11d1-bf92-0060081ed811")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumWIA_DEV_INFO
    {
        [PreserveSig]
        int Next(int celt, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] out IWiaPropertyStorage[] rgelt, out int pceltFetched);
        int Skip();
        void Reset();
        IEnumWIA_DEV_INFO Clone();
        int GetCount();
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPSPEC
    {
        int propid;
        IntPtr unionmember;
    }

    [ComImport, Guid("98B5E8A0-29CC-491a-AAC0-E6DB4FDCCEB6")]
    public interface IWiaPropertyStorage
    {
        [PreserveSig]
        int ReadMultiple(
            int cpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] in PROPSPEC[] rgpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] out object[] rgpropvar
        );
        
        [PreserveSig]
        int WriteMultiple(
            int cpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] in PROPSPEC[] rgpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] object[] rgpropvar,
            int propidNameFirst
        );
        
        void DeleteMultiple(
            int cpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] in PROPSPEC[] rgpspec
        );
        
        void ReadPropertyNames(
            int cpropid,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] in int[] rgpropid,
            [MarshalAs(UnmanagedType.LPArray | UnmanagedType.LPWStr, SizeParamIndex = 0)] out string[] rglpwstrName
        );
        
        void WritePropertyNames(
            int cpropid,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] in int[] rgpropid,
            [MarshalAs(UnmanagedType.LPArray | UnmanagedType.LPWStr, SizeParamIndex = 0)] in string[] rglpwstrName
        );
        
        void DeletePropertyNames(
            int cpropid,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] in int[] rgpropid
        );
        
        void Commit(int grfCommitFlags);
        
        void Revert();
        
        void Enum([MarshalAs(UnmanagedType.Interface)] out IntPtr ppenum);
        
        void SetTimes(
             [MarshalAs(UnmanagedType.I8)] IntPtr pctime,
             [MarshalAs(UnmanagedType.I8)] IntPtr patime,
             [MarshalAs(UnmanagedType.I8)] IntPtr pmtime);
        
        void SetClass(in Guid clsid);
        
        IntPtr Stat();
        
        [PreserveSig]
        int GetPropertyAttributes(
            int cpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] out int[] rgflags,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] out object[] rgpropvar
        );
        
        int GetCount();
        
        IStream GetPropertyStream(out Guid pCompatibilityId);
        
        void SetPropertyStream(in Guid pCompatibilityId, in IStream pIStream);
    }

    [ComImport, Guid("79C07CF1-CBDD-41ee-8EC3-F00080CADA7A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaDevMgr2
    {
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumWIA_DEV_INFO EnumDeviceInfo(int lFlags);
        IWiaItem2 CreateDevice(int lFlags, string bstrDeviceID);
        IWiaItem2 SelectDeviceDlg(IntPtr hwndParent, int lDeviceType, int lFlags, ref string pbstrDeviceID);
        string SelectDeviceDlgID(IntPtr hwndParent, int lDeviceType, int lFlags);
        [PreserveSig]
        int RegisterEventCallbackInterface(int lFlags, string bstrDeviceID, in Guid pEventGUID, IWiaEventCallback pIWiaEventCallback, [MarshalAs(UnmanagedType.Interface, IidParameterIndex = 2)] out IntPtr pEventObject);
        void RegisterEventCallbackProgram(int lFlags, string bstrDeviceID, in Guid pEventGUID, string bstrFullAppName, string bstrCommandLineArg, string bstrName, string bstrDescription, string bstrIcon);
        void RegisterEventCallbackCLSID(int lFlags, string bstrDeviceID, in Guid pEventGUID, in Guid pClsID, string bstrName, string bstrDescription, string bstrIcon);
        IWiaItem2 GetImageDlg(int lFlags, string bstrDeviceID, IntPtr hwndParent, string bstrFolderName, string bstrFilename, out int plNumFiles, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 5)] out string[] ppbstrFilePaths);
    };

    [ComImport, Guid("B6C292BC-7C88-41ee-8B54-8EC92617E599")]
    public class WiaDevMgr2
    {
    }

    class CWiaTransferCallback2 : IWiaTransferCallback
    {
        const int WIA_TRANSFER_MSG_STATUS = 0x00001;
        const int WIA_TRANSFER_MSG_END_OF_STREAM = 0x00002;
        const int WIA_TRANSFER_MSG_END_OF_TRANSFER = 0x00003;
        const int WIA_TRANSFER_MSG_DEVICE_STATUS = 0x00005;
        const int WIA_TRANSFER_MSG_NEW_PAGE = 0x00006;

        Queue<IStream> m_streams;
        NativeWiaMethods.TransferStatusCallback m_statusCallback = null;

        public CWiaTransferCallback2(NativeWiaMethods.TransferStatusCallback statusCallback)
        {
            m_statusCallback = statusCallback;
        }

        public void TransferCallback(int lFlags, in WiaTransferParams pWiaTransferParams)
        {
            IStream stream = null;
            int hr = 0;

            switch (pWiaTransferParams.lMessage)
            {
                case WIA_TRANSFER_MSG_STATUS:
                    break;
                case WIA_TRANSFER_MSG_END_OF_STREAM:
                    stream = m_streams.Dequeue();
                    break;
                case WIA_TRANSFER_MSG_END_OF_TRANSFER:
                    break;
                default:
                    break;
            }

            if (m_statusCallback != null)
            {
                m_statusCallback.Invoke(
                    pWiaTransferParams.lMessage,
                    pWiaTransferParams.lPercentComplete,
                    (ulong)pWiaTransferParams.ulTransferredBytes,
                    (uint)pWiaTransferParams.hrErrorStatus,
                    stream
                );
            }
            {
                hr = 1;
            }

            if (stream != null)
            {
                Marshal.ReleaseComObject(stream);
            }
        }

        public IStream GetNextStream(int lFlags, string bstrItemName, string bstrFullItemName)
        {
            throw new NotImplementedException();
        }
    }

    internal static class NativeWiaMethods
    {
        public const int WIA_DEVINFO_ENUM_LOCAL = 0x00000010;

        static NativeWiaMethods()
        {
            const string lib64Dir = "64";
            if (Environment.Is64BitProcess && PlatformCompat.System.CanUseWin32)
            {
                var location = Assembly.GetExecutingAssembly().Location;
                var coreDllDir = System.IO.Path.GetDirectoryName(location);
                if (coreDllDir != null)
                {
                    Win32.SetDllDirectory(System.IO.Path.Combine(coreDllDir, lib64Dir));
                }
            }
        }

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint GetDeviceManager1([Out] out IntPtr deviceManager);

        public static uint GetDeviceManager2([Out] out IntPtr deviceManager)
        {
            var temp = Activator.CreateInstance(typeof(WiaDevMgr2));
            deviceManager = Marshal.GetComInterfaceForObject(temp, typeof(IWiaDevMgr2));
            return 0;
        }

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint GetDevice1(IntPtr deviceManager, [MarshalAs(UnmanagedType.BStr)] string deviceId, [Out] out IntPtr device);

        public static uint GetDevice2(IntPtr deviceManager, [MarshalAs(UnmanagedType.BStr)] string deviceId, [Out] out IntPtr device)
        {
            var temp = (IWiaDevMgr2) Marshal.GetObjectForIUnknown(deviceManager);
            device = Marshal.GetComInterfaceForObject(temp.CreateDevice(0, deviceId), typeof(IWiaItem2));
            return 0;
        }

        public delegate void EnumDeviceCallback(IntPtr devicePropStorage);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint EnumerateDevices1(IntPtr deviceManager, [MarshalAs(UnmanagedType.FunctionPtr)] EnumDeviceCallback func);

        public static uint EnumerateDevices2(IntPtr deviceManager, [MarshalAs(UnmanagedType.FunctionPtr)] EnumDeviceCallback func)
        {
            var temp = (IWiaDevMgr2)Marshal.GetObjectForIUnknown(deviceManager);
            var deviceInfos = temp.EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL);
            int hr = 0;

            while (hr == 0)
            {
                IWiaPropertyStorage[] propStorage;
                hr = deviceInfos.Next(1, out propStorage, out _);
                if (hr == 0)
                {
                    func(Marshal.GetIUnknownForObject(propStorage[0]));
                }
            }
            if (hr == 1)
            {
                hr = 0;
            }

            return (uint) hr;
        }

        public delegate void EnumItemCallback(IntPtr item);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint EnumerateItems1(IntPtr item, [MarshalAs(UnmanagedType.FunctionPtr)] EnumItemCallback func);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint EnumerateItems2(IntPtr item, [MarshalAs(UnmanagedType.FunctionPtr)] EnumItemCallback func);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint GetItemPropertyStorage(IntPtr item, out IntPtr propStorage);

        public delegate void EnumPropertyCallback(int propId, [MarshalAs(UnmanagedType.LPWStr)] string propName, ushort propType);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint EnumerateProperties(IntPtr propStorage, [MarshalAs(UnmanagedType.FunctionPtr)] EnumPropertyCallback func);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint GetPropertyBstr(IntPtr propStorage, int propId, [MarshalAs(UnmanagedType.BStr), Out] out string value);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint GetPropertyInt(IntPtr propStorage, int propId, [Out] out int value);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint SetPropertyInt(IntPtr propStorage, int propId, int value);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint GetPropertyAttributes(
            IntPtr propStorage,
            int propId,
            [Out] out int flags,
            [Out] out int min,
            [Out] out int nom,
            [Out] out int max,
            [Out] out int step,
            [Out] out int numElems,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 7), Out] out int[] elems);
        
        [DllImport("NAPS2.WIA.dll")]
        public static extern uint StartTransfer1(IntPtr item, [Out] out IntPtr transfer);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint StartTransfer2(IntPtr item, [Out] out IntPtr transfer);

        public delegate bool TransferStatusCallback(int msgType, int percent, ulong bytesTransferred, uint hresult, [MarshalAs(UnmanagedType.Interface)] IStream stream);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint Download1(IntPtr transfer, [MarshalAs(UnmanagedType.FunctionPtr)] TransferStatusCallback func);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint Download2(IntPtr transfer, [MarshalAs(UnmanagedType.FunctionPtr)] TransferStatusCallback func);

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint SelectDevice1(IntPtr deviceManager, IntPtr hwnd, int deviceType, int flags, [MarshalAs(UnmanagedType.BStr), Out] out string deviceId, [Out] out IntPtr device);

        public static uint SelectDevice2(IntPtr deviceManager, IntPtr hwnd, int deviceType, int flags, [MarshalAs(UnmanagedType.BStr), Out] out string deviceId, [Out] out IntPtr device)
        {
            var temp = (IWiaDevMgr2) Marshal.GetObjectForIUnknown(deviceManager);
            string _deviceId = String.Empty;
            device = Marshal.GetIUnknownForObject(temp.SelectDeviceDlg(hwnd, deviceType, flags, ref _deviceId));
            deviceId = _deviceId;
            return 0;
        }

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint GetImage1(IntPtr deviceManager, IntPtr hwnd, int deviceType, int flags, int intent, [MarshalAs(UnmanagedType.BStr)] string filePath, IntPtr item);

        public static uint GetImage2(
            IntPtr deviceManager,
            int flags,
            [MarshalAs(UnmanagedType.BStr)] string deviceId,
            IntPtr hwnd,
            [MarshalAs(UnmanagedType.BStr)] string folder,
            [MarshalAs(UnmanagedType.BStr)] string fileName, [In, Out] ref int numFiles,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 6, ArraySubType = UnmanagedType.BStr), In, Out] ref string[] filePaths,
            [In, Out] ref IntPtr item)
        {
            var temp = (IWiaDevMgr2)Marshal.GetObjectForIUnknown(deviceManager);
            item = Marshal.GetIUnknownForObject(temp.GetImageDlg(flags, deviceId, hwnd, folder, fileName, out numFiles, out filePaths));
            return 0;
        }

        [DllImport("NAPS2.WIA.dll")]
        public static extern uint ConfigureDevice1(IntPtr device, IntPtr hwnd, int flags, int intent, [In, Out] ref int itemCount, [In, Out] ref IntPtr[] items);

        public static uint ConfigureDevice2(
            IntPtr device,
            int flags,
            IntPtr hwnd,
            [MarshalAs(UnmanagedType.BStr)] string folder,
            [MarshalAs(UnmanagedType.BStr)] string fileName,
            [In, Out] ref int numFiles,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 5, ArraySubType = UnmanagedType.BStr), In, Out] ref string[] filePaths,
            IntPtr[] items)
        {
            var temp = (IWiaItem2) Marshal.GetObjectForIUnknown(device);
            var _items = temp.DeviceDlg(flags, hwnd, folder, fileName, out numFiles, out filePaths);
            items = new IntPtr[_items.Length];
            for (int i = 0; i < _items.Length; i++)
            {
                items[i] = Marshal.GetIUnknownForObject(_items[i]);
            }
            return 0;
        }
    }
}
