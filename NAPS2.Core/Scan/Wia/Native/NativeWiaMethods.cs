using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NAPS2.Platform;
using NAPS2.Util;

namespace NAPS2.Scan.Wia.Native
{
    [StructLayout(LayoutKind.Sequential)]
    public struct PROPSPEC
    {
        public uint ulKind;
        public IntPtr unionmember;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct STATPROPSTG
    {
        [MarshalAs(UnmanagedType.LPWStr)] public string lpwstrName;
        public uint propid;
        public ushort vt;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct STATPROPSETSTG
    {
        public Guid fmtid;
        public Guid clsid;
        public uint grfFlags;
        public System.Runtime.InteropServices.ComTypes.FILETIME mtime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ctime;
        public System.Runtime.InteropServices.ComTypes.FILETIME atime;
        public uint dwOSVersion;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WIA_FORMAT_INFO
    {
        public Guid guidFormatID;
        public int lTymed;
    }

    [ComImport, Guid("00000139-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumSTATPROPSTG
    {
        [PreserveSig]
        int Next(
            uint celt,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] STATPROPSTG[] rgelt,
            out uint pceltFetched
        );

        void Skip(uint celt);

        void Reset();

        IEnumSTATPROPSTG Clone();
    }

    [ComImport, Guid("98B5E8A0-29CC-491a-AAC0-E6DB4FDCCEB6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaPropertyStorage
    {
        void ReadMultiple(
            uint cpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPVARIANT[] rgpropvar
        );
        void WriteMultiple(
            uint cpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPVARIANT[] rgpropvar,
            uint propidNameFirst
        );
        void DeleteMultiple(
            uint cpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec
        );
        void ReadPropertyNames(
            uint cpropid,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0, ArraySubType = UnmanagedType.LPWStr)] string[] rglpwstrName
        );
        void WritePropertyNames(
            uint cpropid,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0, ArraySubType = UnmanagedType.LPWStr)] string[] rglpwstrName
        );
        void DeletePropertyNames(
            uint cpropid,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid
        );
        void Commit(uint grfCommitFlags);
        void Revert();
        IEnumSTATPROPSTG Enum();
        void SetTimes(
            in System.Runtime.InteropServices.ComTypes.FILETIME pctime,
            in System.Runtime.InteropServices.ComTypes.FILETIME patime,
            in System.Runtime.InteropServices.ComTypes.FILETIME pmtime
        );
        void SetClass(in Guid clsid);
        void Stat(out STATPROPSETSTG pstatpsstg);
        void GetPropertyAttributes(
            uint cpspec,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPSPEC[] rgpspec,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgflags,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] PROPVARIANT[] rgpropvar
        );
        uint GetCount();
        void GetPropertyStream(out Guid pCompatibilityId, out IStream ppIStream);
        void SetPropertyStream(in Guid pCompatibilityId, IStream pIStream);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WiaTransferParams
    {
        public int lMessage;
        public int lPercentComplete;
        public ulong ulTransferredBytes;
        public int hrErrorStatus;
    }

    [ComImport, Guid("27d4eaaf-28a6-4ca5-9aab-e678168b9527")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaTransferCallback
    {
        [PreserveSig]
        int TransferCallback(int lFlags, in WiaTransferParams pWiaTransferParams);
        void GetNextStream(int lFlags, string bstrItemName, string bstrFullItemName, out IStream ppDestination);

    }

    [ComImport, Guid("81BEFC5B-656D-44f1-B24C-D41D51B4DC81")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumWIA_FORMAT_INFO
    {
        [PreserveSig]
        int Next(
            uint celt,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] WIA_FORMAT_INFO[] rgelt,
            out uint pceltFetched
        );
        void Skip(uint celt);
        void Reset();
        IEnumWIA_FORMAT_INFO Clone();
        uint GetCount();
    }

    [ComImport, Guid("a558a866-a5b0-11d2-a08f-00c04f72dc3c")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaDataCallback
    {
        [PreserveSig]
        int BandedDataCallback(
            int lMessage,
            int lStatus,
            int lPercentComplete,
            int lOffset,
            int lLength,
            int lReserved,
            int lResLength,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)] byte[] pbBuffer
        );
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WIA_DATA_TRANSFER_INFO
    {
        public uint ulSize;
        public uint ulSection;
        public uint ulBufferSize;
        [MarshalAs(UnmanagedType.Bool)]
        public bool bDoubleBuffer;
        public uint ulReserved1;
        public uint ulReserved2;
        public uint ulReserved3;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WIA_EXTENDED_TRANSFER_INFO
    {
        public uint ulSize;
        public uint ulMinBufferSize;
        public uint ulOptimalBufferSize;
        public uint ulMaxBufferSize;
        public uint ulNumBuffers;
    }

    [ComImport, Guid("a6cef998-a5b0-11d2-a08f-00c04f72dc3c")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaDataTransfer
    {
        [PreserveSig]
        int idtGetData(ref STGMEDIUM pMedium, IWiaDataCallback pIWiaDataCallback);
        [PreserveSig]
        int idtGetBandedData(in WIA_DATA_TRANSFER_INFO pWiaDataTransInfo, IWiaDataCallback pIWiaDataCallback);
        void idtQueryGetData(in WIA_FORMAT_INFO pfe);
        IEnumWIA_FORMAT_INFO idtEnumWIA_FORMAT_INFO();
        void idtGetExtendedTransferInfo(out WIA_EXTENDED_TRANSFER_INFO pExtendedTransferInfo);

    };

    [ComImport, Guid("c39d6942-2f4e-4d04-92fe-4ef4d3a1de5a")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaTransfer
    {
        void Download(int lFlags, IWiaTransferCallback pIWiaTransferCallback);
        void Upload(int lFlags, IStream pSource, IWiaTransferCallback pIWiaTransferCallback);
        void Cancel();
        IEnumWIA_FORMAT_INFO EnumWIA_FORMAT_INFO();
    }

    [ComImport, Guid("59970AF4-CD0D-44d9-AB24-52295630E582")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumWiaItem2
    {
        [PreserveSig]
        int Next(
            uint cElt,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IWiaItem2[] ppIWiaItem2,
            out uint pcEltFetched
        );
        void Skip(uint cElt);
        void Reset();
        IEnumWiaItem2 Clone();
        uint GetCount();
    }

    [ComImport, Guid("95C2B4FD-33F2-4d86-AD40-9431F0DF08F7")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaPreview
    {
        void GetNewPreview(
            int lFlags,
            IWiaItem2 pWiaItem2,
            IWiaTransferCallback pWiaTransferCallback
        );
        void UpdatePreview(
            int lFlags,
            IWiaItem2 pChildWiaItem2,
            IWiaTransferCallback pWiaTransferCallback
        );
        void DetectRegions(int lFlags);
        void Clear();
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WIA_DEV_CAP
    {
        public Guid guid;
        public uint ulFlags;
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
        int Next(
            uint celt,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] WIA_DEV_CAP[] rgelt,
            out uint pceltFetched
        );
        void Skip(uint celt);
        void Reset();
        IEnumWIA_DEV_CAPS Clone();
        uint GetCount();
    }

    [ComImport, Guid("5e8383fc-3391-11d2-9a33-00c04fa36145")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumWiaItem
    {
        [PreserveSig]
        int Next(
            uint celt,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IWiaItem[] ppIWiaItem,
            out uint pceltFetched
        );
        void Skip(uint celt);
        void Reset();
        IEnumWiaItem Clone();
        uint GetCount();
    }

    [ComImport, Guid("4db1ad10-3391-11d2-9a33-00c04fa36145")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaItem
    {
        int GetItemType();
        void AnalyzeItem(int lFlags);
        IEnumWiaItem EnumChildItems();
        void DeleteItem(int lFlags);
        IWiaItem CreateChildItem(int lFlags, string bstrItemName, string bstrFullItemName);
        IEnumWIA_DEV_CAPS EnumRegisterEventInfo(int lFlags, in Guid pEventGUID);
        IWiaItem FindItemByName(int lFlags, string bstrFullItemName);
        
        [PreserveSig]
        int DeviceDlg(IntPtr hwndParent, int lFlags, int lIntent, out int plItemCount, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 3)] out IWiaItem[] ppIWiaItem);
        IWiaItem DeviceCommand(int lFlags, in Guid pCmdGUID);
        IWiaItem GetRootItem();
        IEnumWIA_DEV_CAPS EnumDeviceCapabilities(int lFlags);
        string DumpItemData();
        string DumpDrvItemData();
        string DumpTreeItemData();
        void Diagnostic(
            uint ulSize,
            [Out, MarshalAs(UnmanagedType.LPArray)] byte[] pBuffer
        );
    }

    [ComImport, Guid("6CBA0075-1287-407d-9B77-CF0E030435CC")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaItem2
    {
        IWiaItem2 CreateChildItem(int lItemFlags, int lCreationFlags, string bstrItemName);

        void DeleteItem(int lFlags);
        IEnumWiaItem2 EnumChildItems(IntPtr pCategoryGUID);
        IWiaItem2 FindItemByName(int lFlags, string bstrFullItemName);
        void GetItemCategory(out Guid pItemCategoryGUID);
        int GetItemType();
        void DeviceDlg(
            int lFlags,
            IntPtr hwndParent,
            string bstrFolderName,
            string bstrFilename,
            out int plNumFiles,
            [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 4)] string[] ppbstrFilePaths,
            ref IWiaItem2 ppItem
        );
        void DeviceCommand(int lFlags, in Guid pCmdGUID, out IWiaItem2 ppIWiaItem2);
        IEnumWIA_DEV_CAPS EnumDeviceCapabilities(int lFlags);
        bool CheckExtension(int lFlags, string bstrName, in Guid riidExtensionInterface);
        void GetExtension(int lFlags, string bstrName, in Guid riidExtensionInterface, [MarshalAs(UnmanagedType.IUnknown)] out object ppOut);
        IWiaItem2 GetParentItem();
        IWiaItem2 GetRootItem();
        IWiaPreview GetPreviewComponent(int lFlags);
        IEnumWIA_DEV_CAPS EnumRegisterEventInfo(int lFlags, in Guid pEventGUID);
        void Diagnostic(uint ulSize, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] byte[] pBuffer);
    }

    [ComImport, Guid("5e38b83c-8cf1-11d1-bf92-0060081ed811")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumWIA_DEV_INFO
    {
        [PreserveSig]
        int Next(
            uint celt,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IWiaPropertyStorage[] rgelt,
            out uint pceltFetched
        );
        void Skip(uint celt);
        void Reset();
        IEnumWIA_DEV_INFO Clone();
        uint GetCount();
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
            uint dwDeviceType,
            string bstrFullItemName,
            ref uint pulEventType,
            uint ulReserved
        );
    }

    [ComImport, Guid("5eb2502a-8cf1-11d1-bf92-0060081ed811")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaDevMgr
    {
        IEnumWIA_DEV_INFO EnumDeviceInfo(int lFlag);
        IWiaItem CreateDevice(string bstrDeviceID);
        IWiaItem SelectDeviceDlg(
            IntPtr hwndParent,
            int lDeviceType,
            int lFlags,
            ref string pbstrDeviceID
        );
        string SelectDeviceDlgID(
            IntPtr hwndParent,
            int lDeviceType,
            int lFlags
        );

        void GetImageDlg(
            IntPtr hwndParent,
            int lDeviceType,
            int lFlags,
            int lIntent,
            IWiaItem pItemRoot,
            string bstrFilename,
            ref Guid pguidFormat
        );
        void RegisterEventCallbackProgram(
            int lFlags,
            string bstrDeviceID,
            in Guid pEventGUID,
            string bstrCommandline,
            string bstrName,
            string bstrDescription,
            string bstrIcon
        );
        void RegisterEventCallbackInterface(
            int lFlags,
            string bstrDeviceID,
            in Guid pEventGUID,
            IWiaEventCallback pIWiaEventCallback,
            [MarshalAs(UnmanagedType.IUnknown)] out object pEventObject
        );

        void RegisterEventCallbackCLSID(
            int lFlags,
            string bstrDeviceID,
            in Guid pEventGUID,
            in Guid pClsID,
            string bstrName,
            string bstrDescription,
            string bstrIcon
        );
        void AddDeviceDlg(IntPtr hwndParent, int lFlags);
    }

    [ComImport, Guid("a1f4e726-8cf1-11d1-bf92-0060081ed811")]
    public class WiaDevMgr
    {
    }

    [ComImport, Guid("79C07CF1-CBDD-41ee-8EC3-F00080CADA7A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWiaDevMgr2
    {
        IEnumWIA_DEV_INFO EnumDeviceInfo(int lFlags);
        IWiaItem2 CreateDevice(int lFlags, string bstrDeviceID);
        IWiaItem2 SelectDeviceDlg(IntPtr hwndParent, int lDeviceType, int lFlags, ref string pbstrDeviceID);
        string SelectDeviceDlgID(IntPtr hwndParent, int lDeviceType, int lFlags);
        void RegisterEventCallbackInterface(
            int lFlags,
            string bstrDeviceID,
            in Guid pEventGUID,
            IWiaEventCallback pIWiaEventCallback,
            [MarshalAs(UnmanagedType.IUnknown)] out object pEventObject
        );
        void RegisterEventCallbackProgram(
            int lFlags,
            string bstrDeviceID,
            in Guid pEventGUID,
            string bstrFullAppName,
            string bstrCommandLineArg,
            string bstrName,
            string bstrDescription,
            string bstrIcon
        );
        void RegisterEventCallbackCLSID(
            int lFlags,
            string bstrDeviceID,
            in Guid pEventGUID,
            in Guid pClsID,
            string bstrName,
            string bstrDescription,
            string bstrIcon
        );

        [PreserveSig]
        int GetImageDlg(
            int lFlags,
            string bstrDeviceID,
            IntPtr hwndParent,
            string bstrFolderName,
            string bstrFilename,
            ref int plNumFiles,
            [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 5, ArraySubType = UnmanagedType.BStr)] ref string[] ppbstrFilePaths,
            ref IWiaItem2 ppItem
        );
    }

    [ComImport, Guid("B6C292BC-7C88-41ee-8B54-8EC92617E599")]
    public class WiaDevMgr2
    {
    };

    internal static class NativeWiaMethods
    {
        public static uint GetDeviceManager1(out IWiaDevMgr deviceManager)
        {
            try
            {
                deviceManager = (IWiaDevMgr)new WiaDevMgr();
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                deviceManager = null;
                return (uint)ex.HResult;
            }
        }

        public static uint GetDeviceManager2(out IWiaDevMgr2 deviceManager)
        {
            try
            {
                deviceManager = (IWiaDevMgr2)new WiaDevMgr2();
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                deviceManager = null;
                return (uint)ex.HResult;
            }
        }

        public static uint GetDevice1(IWiaDevMgr deviceManager, string deviceId, out IWiaItem device)
        {
            try
            {
                device = deviceManager.CreateDevice(deviceId);
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                device = null;
                return (uint)ex.HResult;
            }
        }

        public static uint GetDevice2(IWiaDevMgr2 deviceManager, string deviceId, out IWiaItem2 device)
        {
            try
            {
                device = deviceManager.CreateDevice(0, deviceId);
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                device = null;
                return (uint)ex.HResult;
            }
        }

        public delegate void EnumDeviceCallback(IWiaPropertyStorage devicePropStorage);

        public static uint EnumerateDevices1(IWiaDevMgr deviceManager, EnumDeviceCallback func)
        {
            const int WIA_DEVINFO_ENUM_LOCAL = 0x00000010;

            try
            {
                IEnumWIA_DEV_INFO enumDevInfo = deviceManager.EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL);
                var hr = Hresult.S_OK;
                do
                {
                    var propStorage = new IWiaPropertyStorage[1];
                    hr = (uint)enumDevInfo.Next(1, propStorage, out var fetched);
                    if (hr == Hresult.S_OK)
                    {
                        if (fetched != 1)
                            break;

                        func(propStorage[0]);
                    }
                } while (hr == Hresult.S_OK);

                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                return (uint)ex.HResult;
            }
        }

        public static uint EnumerateDevices2(IWiaDevMgr2 deviceManager, EnumDeviceCallback func)
        {
            const int WIA_DEVINFO_ENUM_LOCAL = 0x00000010;

            try
            {
                IEnumWIA_DEV_INFO enumDevInfo = deviceManager.EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL);
                var hr = Hresult.S_OK;
                do
                {
                    var propStorage = new IWiaPropertyStorage[1];
                    hr = (uint)enumDevInfo.Next(1, propStorage, out var fetched);
                    if (hr == Hresult.S_OK)
                    {
                        if (fetched != 1)
                            break;

                        func(propStorage[0]);
                    }
                } while (hr == Hresult.S_OK);

                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                return (uint)ex.HResult;
            }
        }

        public delegate void EnumItemCallback1(IWiaItem item);

        public delegate void EnumItemCallback2(IWiaItem2 item);

        public static uint EnumerateItems1(IWiaItem item, EnumItemCallback1 func)
        {
            const int WiaItemTypeFolder = 0x00000004;
            const int WiaItemTypeHasAttachments = 0x00008000;

            try
            {
                int itemType = item.GetItemType();
                if (((itemType & WiaItemTypeFolder) == WiaItemTypeFolder) || ((itemType & WiaItemTypeHasAttachments) == WiaItemTypeHasAttachments))
                {
                    IEnumWiaItem enumerator = item.EnumChildItems();
                    if (enumerator != null)
                    {
                        uint hr;
                        do
                        {
                            var items = new IWiaItem[1];
                            hr = (uint)enumerator.Next(1, items, out var fetched);
                            if (hr == Hresult.S_OK)
                            {
                                if (fetched != 1)
                                    break;

                                func(items[0]);
                            }

                        } while (hr == Hresult.S_OK);
                    }
                }

                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                return (uint)ex.HResult;
            }
        }

        public static uint EnumerateItems2(IWiaItem2 item, EnumItemCallback2 func)
        {
            const int WiaItemTypeFolder = 0x00000004;
            const int WiaItemTypeHasAttachments = 0x00008000;

            try
            {
                int itemType = item.GetItemType();
                if (((itemType & WiaItemTypeFolder) == WiaItemTypeFolder) || ((itemType & WiaItemTypeHasAttachments) == WiaItemTypeHasAttachments))
                {
                    IEnumWiaItem2 enumerator = item.EnumChildItems(IntPtr.Zero);
                    if (enumerator != null)
                    {
                        uint hr;
                        do
                        {
                            var items = new IWiaItem2[1];
                            hr = (uint)enumerator.Next(1, items, out var fetched);
                            if (hr == Hresult.S_OK)
                            {
                                if (fetched != 1)
                                    break;

                                func(items[0]);
                            }

                        } while (hr == Hresult.S_OK);
                    }
                }

                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                return (uint)ex.HResult;
            }
        }

        public static uint GetItemPropertyStorage(object item, out IWiaPropertyStorage propStorage)
        {
            try
            {
                propStorage = (IWiaPropertyStorage)item;
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                propStorage = null;
                return (uint)ex.HResult;
            }
        }

        public delegate void EnumPropertyCallback(int propId, [MarshalAs(UnmanagedType.LPWStr)] string propName, ushort propType);

        public static uint EnumerateProperties(IWiaPropertyStorage propStorage, EnumPropertyCallback func)
        {
            try
            {
                IEnumSTATPROPSTG enumProps = propStorage.Enum();
                if (enumProps != null)
                {
                    uint hr;
                    do
                    {
                        var props = new STATPROPSTG[1];
                        hr = (uint)enumProps.Next(1, props, out var fetched);
                        if (hr == Hresult.S_OK)
                        {
                            if (fetched != 1)
                                break;

                            func((int)props[0].propid, props[0].lpwstrName, props[0].vt);
                        }
                    }
                    while (hr == Hresult.S_OK);
                }
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                return (uint)ex.HResult;
            }
        }

        public static uint GetPropertyBstr(IWiaPropertyStorage propStorage, int propId, out string value)
        {
            var PropSpec = new PROPSPEC[1] { new PROPSPEC { ulKind = 1, unionmember = new IntPtr((uint)propId) } };
            var PropVariant = new PROPVARIANT[1];
            try
            {
                try
                {
                    propStorage.ReadMultiple(1, PropSpec, PropVariant);
                    PropVariantToBSTR(PropVariant[0], out value);

                    return Hresult.S_OK;
                }
                finally
                {
                    PropVariantClear(PropVariant[0]);
                }
            }
            catch (COMException ex)
            {
                value = String.Empty;
                return (uint)ex.HResult;
            }
        }

        public static uint GetPropertyInt(IWiaPropertyStorage propStorage, int propId, out int value)
        {
            var PropSpec = new PROPSPEC[1] { new PROPSPEC { ulKind = 1, unionmember = new IntPtr((uint)propId) } };
            var PropVariant = new PROPVARIANT[1];
            try
            {
                try
                {
                    propStorage.ReadMultiple(1, PropSpec, PropVariant);
                    PropVariantToInt32(PropVariant[0], out value);

                    return Hresult.S_OK;
                }
                finally
                {
                    PropVariantClear(PropVariant[0]);
                }
            }
            catch (COMException ex)
            {
                value = 0;
                return (uint)ex.HResult;
            }
        }

        public static uint SetPropertyInt(IWiaPropertyStorage propStorage, int propId, int value)
        {
            const int WIA_IPA_FIRST = 4098;

            try
            {
                var PropSpec = new PROPSPEC[1] { new PROPSPEC { ulKind = 1, unionmember = new IntPtr((uint)propId) } };
                var PropVariants = new PROPVARIANT[1];
                try
                {
                    PropVariants[0].Init(value);
                    propStorage.WriteMultiple(1, PropSpec, PropVariants, WIA_IPA_FIRST);
                }
                finally
                {
                    PropVariantClear(PropVariants[0]);
                }
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                return (uint)ex.HResult;
            }
        }

        public static void GetPropertyAttributes(
            IWiaPropertyStorage propStorage,
            int propId,
            out int flags,
            out int min,
            out int nom,
            out int max,
            out int step,
            out int numElems,
            out int[] elems)
        {
            const uint WIA_PROP_RANGE = 0x10;
            const uint WIA_PROP_LIST = 0x20;

            const uint WIA_RANGE_MIN = 0;
            const uint WIA_RANGE_NOM = 1;
            const uint WIA_RANGE_MAX = 2;
            const uint WIA_RANGE_STEP = 3;

            const uint WIA_LIST_NOM = 1;

            var PropSpec = new PROPSPEC[1] { new PROPSPEC { ulKind = 1, unionmember = new IntPtr((uint)propId) } };
            var PropFlags = new uint[1];
            var PropVariant = new PROPVARIANT[1];

            elems = null;
            flags = 0;
            max = 0;
            min = 0;
            nom = 0;
            numElems = 0;
            step = 0;
            try
            {
                propStorage.GetPropertyAttributes(1, PropSpec, PropFlags, PropVariant);

                flags = (int)PropFlags[0];
                if ((flags & WIA_PROP_RANGE) == WIA_PROP_RANGE)
                {
                    PropVariantGetInt32Elem(PropVariant[0], WIA_RANGE_MIN, out min);
                    PropVariantGetInt32Elem(PropVariant[0], WIA_RANGE_NOM, out nom);
                    PropVariantGetInt32Elem(PropVariant[0], WIA_RANGE_MAX, out max);
                    PropVariantGetInt32Elem(PropVariant[0], WIA_RANGE_STEP, out step);
                }
                if ((flags & WIA_PROP_LIST) == WIA_PROP_LIST)
                {
                    numElems = (int)PropVariantGetElementCount(PropVariant[0]);
                    PropVariantGetInt32Elem(PropVariant[0], WIA_LIST_NOM, out nom);
                    PropVariantToInt32VectorAlloc(PropVariant[0], out elems, out _);
                }
            }
            finally
            {
                PropVariantClear(PropVariant[0]);
            }
        }

        public static uint StartTransfer1(IWiaItem item, out IWiaDataTransfer transfer)
        {
            const uint WIA_IPA_FIRST = 4098;
            const uint WIA_IPA_FORMAT = 4106;
            const uint WIA_IPA_TYMED = 4108;

            const int TYMED_CALLBACK = 128;

            var WiaImgFmt_BMP = new Guid(0xb96b3cab, 0x0728, 0x11d3, 0x9d, 0x7b, 0x00, 0x00, 0xf8, 0x1e, 0xf3, 0x2e);

            try
            {
                var PropSpec = new PROPSPEC[2] {
                        new PROPSPEC {ulKind = 1, unionmember = new IntPtr(WIA_IPA_FORMAT) },
                        new PROPSPEC {ulKind = 1, unionmember = new IntPtr(WIA_IPA_TYMED) },
                };
                var PropVariant = new PROPVARIANT[2];
                PropVariant[0].Init(WiaImgFmt_BMP);
                PropVariant[1].Init(TYMED_CALLBACK);

                var propStorage = (IWiaPropertyStorage)item;
                propStorage.WriteMultiple(2, PropSpec, PropVariant, WIA_IPA_FIRST);
                transfer = (IWiaDataTransfer)item;
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                transfer = null;
                return (uint)ex.HResult;
            }
        }

        public static uint StartTransfer2(IWiaItem2 item, out IWiaTransfer transfer)
        {
            try
            {
                transfer = (IWiaTransfer)item;
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                transfer = null;
                return (uint)ex.HResult;
            }
        }

        public delegate bool TransferStatusCallback(int msgType, int percent, ulong bytesTransferred, uint hresult, [MarshalAs(UnmanagedType.Interface)] IStream stream);

        public static unsafe uint Download1(IWiaDataTransfer transfer, TransferStatusCallback func)
        {
            var callbackClass = new CWiaTransferCallback1(func);
            /*STGMEDIUM StgMedium = { 0 };
            StgMedium.tymed = TYMED_ISTREAM;
            StgMedium.pstm*/
            var WiaDataTransferInfo = new WIA_DATA_TRANSFER_INFO
            {
                ulSize = (uint)sizeof(WIA_DATA_TRANSFER_INFO),
                ulBufferSize = 131072 * 2,
                bDoubleBuffer = true
            };
            return (uint)transfer.idtGetBandedData(WiaDataTransferInfo, callbackClass);
        }

        public static uint Download2(IWiaTransfer transfer, TransferStatusCallback func)
        {
            try
            {
                var callbackClass = new CWiaTransferCallback2(func);
                transfer.Download(0, callbackClass);
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                return (uint)ex.HResult;
            }
        }

        public static uint SelectDevice1(IWiaDevMgr deviceManager, IntPtr hwnd, int deviceType, int flags, out string deviceId, out IWiaItem device)
        {
            try
            {
                var _deviceId = String.Empty;
                device = deviceManager.SelectDeviceDlg(hwnd, deviceType, flags, ref _deviceId);
                deviceId = _deviceId;
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                deviceId = String.Empty;
                device = null;
                return (uint)ex.HResult;
            }
        }

        public static uint SelectDevice2(IWiaDevMgr2 deviceManager, IntPtr hwnd, int deviceType, int flags, out string deviceId, out IWiaItem2 device)
        {
            try
            {
                var _deviceId = String.Empty;
                device = deviceManager.SelectDeviceDlg(hwnd, deviceType, flags, ref _deviceId);
                deviceId = _deviceId;
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                deviceId = String.Empty;
                device = null;
                return (uint)ex.HResult;
            }
        }

        public static uint GetImage1(IWiaDevMgr deviceManager, IntPtr hwnd, int deviceType, int flags, int intent, [MarshalAs(UnmanagedType.BStr)] string filePath, IWiaItem item)
        {
            var format = new Guid(0xb96b3cae, 0x0728, 0x11d3, 0x9d, 0x7b, 0x00, 0x00, 0xf8, 0x1e, 0xf3, 0x2e);
            try
            {
                deviceManager.GetImageDlg(hwnd, deviceType, flags, intent, item, filePath, format);
                return Hresult.S_OK;
            }
            catch (COMException ex)
            {
                return (uint)ex.HResult;

            }
        }

        public static uint GetImage2(
            IWiaDevMgr2 deviceManager,
            int flags,
            string deviceId,
            IntPtr hwnd,
            string folder,
            string fileName, ref int numFiles,
            ref string[] filePaths,
            ref IWiaItem2 item)
        {
            return (uint)deviceManager.GetImageDlg(flags, deviceId, hwnd, folder, fileName, ref numFiles, ref filePaths, ref item);
        }

        public static uint ConfigureDevice1(IWiaItem device, IntPtr hwnd, int flags, int intent, out int itemCount, out IWiaItem[] items)
        {
            return (uint)device.DeviceDlg(hwnd, flags, intent, out itemCount, out items);
        }

        [DllImport("propsys.dll", PreserveSig = true)]
        public static extern uint PropVariantGetElementCount(in PROPVARIANT propvar);

        [DllImport("propsys.dll", PreserveSig = false)]
        public static extern void PropVariantGetInt32Elem(in PROPVARIANT propvar, uint iElem, out int pnVal);

        [DllImport("propsys.dll", PreserveSig = false)]
        public static extern void PropVariantToInt32(in PROPVARIANT propvarIn, out int plRet);

        [DllImport("propsys.dll", PreserveSig = false)]
        public static extern void InitPropVariantFromInt32(int lVal, out PROPVARIANT ppropvar);
        [DllImport("propsys.dll", PreserveSig = false)]
        public static extern void PropVariantToBSTR(in PROPVARIANT propvar, [MarshalAs(UnmanagedType.BStr)] out string pbstrOut);

        [DllImport("propsys.dll", PreserveSig = false)]
        public static extern void PropVariantToInt32VectorAlloc(
            in PROPVARIANT propvar,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] out int[] pprgn,
            out uint pcElem
        );

        [DllImport("ole32.dll", PreserveSig = false)]
        public static extern void PropVariantClear(in PROPVARIANT pvar);

        [DllImport("shlwapi.dll", PreserveSig = true)]
        public static extern IStream SHCreateMemStream([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pInit, uint cbInit);

        internal class CWiaTransferCallback2 : IWiaTransferCallback
        {
            public const int WIA_TRANSFER_MSG_STATUS = 0x00001;
            public const int WIA_TRANSFER_MSG_END_OF_STREAM = 0x00002;
            public const int WIA_TRANSFER_MSG_END_OF_TRANSFER = 0x00003;
            public const int WIA_TRANSFER_MSG_DEVICE_STATUS = 0x00005;
            public const int WIA_TRANSFER_MSG_NEW_PAGE = 0x00006;

            internal CWiaTransferCallback2(TransferStatusCallback statusCallback)
            {
                m_statusCallback = statusCallback;
            }

            public int TransferCallback(int lFlags, in WiaTransferParams pWiaTransferParams)
            {
                int hr = (int)Hresult.S_OK;
                IStream stream = null;

                switch (pWiaTransferParams.lMessage)
                {
                    case WIA_TRANSFER_MSG_STATUS:
                        break;
                    case WIA_TRANSFER_MSG_END_OF_STREAM:
                        //...
                        stream = m_streams.Dequeue();
                        break;
                    case WIA_TRANSFER_MSG_END_OF_TRANSFER:
                        break;
                    default:
                        break;
                }

                if (!m_statusCallback(
                    pWiaTransferParams.lMessage,
                    pWiaTransferParams.lPercentComplete,
                    pWiaTransferParams.ulTransferredBytes,
                    (uint)pWiaTransferParams.hrErrorStatus,
                    stream))
                {
                    hr = 1;
                }

                if (stream != null)
                {
                    Marshal.ReleaseComObject(stream);
                }

                return hr;
            }

            public void GetNextStream(
                int lFlags,
                string bstrItemName,
                string bstrFullItemName,
                out IStream ppDestination)
            {
                //  Return a new stream for this item's data.
                //
                IStream stream = SHCreateMemStream(null, 0);
                ppDestination = stream;
                m_streams.Enqueue(stream);
            }

            readonly Queue<IStream> m_streams = new Queue<IStream>();
            readonly TransferStatusCallback m_statusCallback;
        };

        internal class CWiaTransferCallback1 : IWiaDataCallback
        {
            const int IT_MSG_DATA_HEADER = 0x0001;
            const int IT_MSG_DATA = 0x0002;
            const int IT_MSG_STATUS = 0x0003;
            const int IT_MSG_TERMINATION = 0x0004;
            const int IT_MSG_NEW_PAGE = 0x0005;
            const int IT_MSG_FILE_PREVIEW_DATA = 0x0006;
            const int IT_MSG_FILE_PREVIEW_DATA_HEADER = 0x0007;

            const int STREAM_SEEK_SET = 0;
            const int STREAM_SEEK_CUR = 1;
            const int STREAM_SEEK_END = 2;

            internal CWiaTransferCallback1(TransferStatusCallback statusCallback)
            {
                m_statusCallback = statusCallback;

            }

            [PreserveSig]
            public int BandedDataCallback(
                int lMessage,
                int lStatus,
                int lPercentComplete,
                int lOffset,
                int lLength,
                int lReserved,
                int lResLength,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)] byte[] pbBuffer
            )
            {
                uint hr = Hresult.S_OK;
                IStream stream = null;

                switch (lMessage)
                {
                    case IT_MSG_DATA_HEADER:
                        break;
                    case IT_MSG_DATA:
                        if (m_stream == null)
                        {
                            m_stream = SHCreateMemStream(null, 0);
                            var empty_header = new byte[] { 0x42, 0x4D, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                            m_stream.Write(empty_header, 14, IntPtr.Zero);
                        }
                        m_stream.Write(pbBuffer, lLength, IntPtr.Zero);
                        break;
                    case IT_MSG_STATUS:
                        break;
                    case IT_MSG_NEW_PAGE:
                        stream = m_stream;
                        m_stream = null;
                        break;
                    case IT_MSG_TERMINATION:
                        stream = m_stream;
                        break;
                    default:
                        break;
                }

                if (stream != null)
                {
                    stream.Seek(14, STREAM_SEEK_SET, IntPtr.Zero);
                    var bytes = new byte[4];
                    stream.Read(bytes, 4, IntPtr.Zero);
                    int header_bytes = BitConverter.ToInt32(bytes, 0);
                    header_bytes += 14;
                    stream.Seek(2, STREAM_SEEK_SET, IntPtr.Zero);
                    stream.Stat(out var stat, 1);
                    bytes = BitConverter.GetBytes(stat.cbSize);
                    stream.Write(bytes, 4, IntPtr.Zero);
                    stream.Seek(10, STREAM_SEEK_SET, IntPtr.Zero);
                    bytes = BitConverter.GetBytes(header_bytes);
                    stream.Write(bytes, 4, IntPtr.Zero);
                    stream.Seek(0, STREAM_SEEK_SET, IntPtr.Zero);
                    m_statusCallback(
                        2,
                        lPercentComplete,
                        (ulong)(lOffset + lLength),
                        (uint)lStatus,
                        stream);
                }

                if (!m_statusCallback(
                    lMessage == IT_MSG_TERMINATION ? 3 : 1,
                    lPercentComplete,
                    (ulong)(lOffset + lLength),
                    (uint)lStatus,
                    null))
                {
                    hr = 1;
                }

                return (int)hr;
            }

            bool cancel;
            IStream m_stream;
            readonly TransferStatusCallback m_statusCallback;
        };
    }
}

