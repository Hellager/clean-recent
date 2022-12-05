using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Shell32;



// https://social.msdn.microsoft.com/Forums/sqlserver/en-US/155aba8d-2fa4-49fe-b5ef-1b114a19e5f2/how-to-programmatically-invoke-unpinfromhome-from-c?forum=windowssdk
public enum HRESULT : int
{
    S_OK = 0,
    S_FALSE = 1,
    E_NOINTERFACE = unchecked((int)0x80004002),
    E_NOTIMPL = unchecked((int)0x80004001),
    E_FAIL = unchecked((int)0x80004005)
}

public enum SIGDN : uint
{
    SIGDN_NORMALDISPLAY = 0,
    SIGDN_PARENTRELATIVEPARSING = 0x80018001,
    SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000,
    SIGDN_PARENTRELATIVEEDITING = 0x80031001,
    SIGDN_DESKTOPABSOLUTEEDITING = 0x8004c000,
    SIGDN_FILESYSPATH = 0x80058000,
    SIGDN_URL = 0x80068000,
    SIGDN_PARENTRELATIVEFORADDRESSBAR = 0x8007c001,
    SIGDN_PARENTRELATIVE = 0x80080001,
    SIGDN_PARENTRELATIVEFORUI = 0x80094001
}

public enum SICHINTF : uint
{
    SICHINT_DISPLAY = 0,
    SICHINT_ALLFIELDS = 0x80000000,
    SICHINT_CANONICAL = 0x10000000,
    SICHINT_TEST_FILESYSPATH_IF_NOT_EQUAL = 0x20000000
}

public enum MenuItemInfo_fMask : uint
{
    MIIM_BITMAP = 0x00000080,           //Retrieves or sets the hbmpItem member.
    MIIM_CHECKMARKS = 0x00000008,       //Retrieves or sets the hbmpChecked and hbmpUnchecked members.
    MIIM_DATA = 0x00000020,             //Retrieves or sets the dwItemData member.
    MIIM_FTYPE = 0x00000100,            //Retrieves or sets the fType member.
    MIIM_ID = 0x00000002,               //Retrieves or sets the wID member.
    MIIM_STATE = 0x00000001,            //Retrieves or sets the fState member.
    MIIM_STRING = 0x00000040,           //Retrieves or sets the dwTypeData member.
    MIIM_SUBMENU = 0x00000004,          //Retrieves or sets the hSubMenu member.
    MIIM_TYPE = 0x00000010,             //Retrieves or sets the fType and dwTypeData members.
                                        //MIIM_TYPE is replaced by MIIM_BITMAP, MIIM_FTYPE, and MIIM_STRING.
}

public enum MenuString_Pos : uint
{
    MF_BYCOMMAND = 0x00000000,
    MF_BYPOSITION = 0x00000400,
}

public enum ShowWindowCommands : uint
{
    SW_HIDE = 0,
    SW_SHOWNORMAL = 1,
    SW_NORMAL = 1,
    SW_SHOWMINIMIZED = 2,
    SW_SHOWMAXIMIZED = 3,
    SW_MAXIMIZE = 3,
    SW_SHOWNOACTIVATE = 4,
    SW_SHOW = 5,
    SW_MINIMIZE = 6,
    SW_SHOWMINNOACTIVE = 7,
    SW_SHOWNA = 8,
    SW_RESTORE = 9,
    SW_SHOWDEFAULT = 10,
    SW_FORCEMINIMIZE = 11,
    SW_MAX = 11
}

public enum QuickAccessType
{
    FrequentFolder = 1,
    RecentFile
};

public enum ShellAddToRecentDocsFlags
{
    SHARD_PIDL = 0x001,
    SHARD_PATHA = 0x002,
    SHARD_PATHW = 0x003
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
public class MENUITEMINFO
{
    public int cbSize;
    public uint fMask;
    public uint fType;
    public uint fState;
    public uint wID;
    public IntPtr hSubMenu;
    public IntPtr hbmpChecked;
    public IntPtr hbmpUnchecked;
    public IntPtr dwItemData;
    public IntPtr dwTypeData;
    public uint cch;
    public IntPtr hbmpItem;

    public MENUITEMINFO()
    {
        cbSize = Marshal.SizeOf(typeof(MENUITEMINFO));
    }
}

[ComImport]
[Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IShellItem
{
    HRESULT BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
    HRESULT GetParent(out IShellItem ppsi);
    HRESULT GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
    HRESULT GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    HRESULT Compare(IShellItem psi, SICHINTF hint, out int piOrder);
}

[ComImport]
[Guid("70629033-e363-4a28-a567-0db78006e6d7")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IEnumShellItems
{
    HRESULT Next(uint celt, out IShellItem rgelt, out uint pceltFetched);
    HRESULT Skip(uint celt);
    HRESULT Reset();
    HRESULT Clone(out IEnumShellItems ppenum);
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
public struct CMINVOKECOMMANDINFO
{
    public int cbSize;
    public int fMask;
    public IntPtr hwnd;
    //public string lpVerb;
    public int lpVerb;
    public string lpParameters;
    public string lpDirectory;
    public int nShow;
    public int dwHotKey;
    public IntPtr hIcon;
}

[ComImport]
[Guid("000214e4-0000-0000-c000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IContextMenu
{
    HRESULT QueryContextMenu(IntPtr hmenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
    HRESULT InvokeCommand(ref CMINVOKECOMMANDINFO pici);
    HRESULT GetCommandString(uint idCmd, uint uType, out uint pReserved, StringBuilder pszName, uint cchMax);
}


namespace QuickAccess
{
    public class QuickAccessHandler
    {
        dynamic quickAccessShell;
        List<string> QuickAccessMenuName;
        List<string> FileExplorerMenuName;
        string QuickAccessRegistryKeyPath;
        public const int SEE_MASK_ASYNCOK = 0x00100000;
        public const int CMIC_MASK_ASYNCOK = SEE_MASK_ASYNCOK;

        // <resource name, resource path>
        Dictionary<string, string> frequentFolders;
        Dictionary<string, string> recentFiles;
        Dictionary<string, string> unspecificContent;

        public enum QuickAccessType
        {
            FrequentFolder = 0,
            RecentFile,
            Undefined
        }

        [DllImport("Shell32.dll", BestFitMapping = false, ThrowOnUnmappableChar = true)]
        private static extern void SHAddToRecentDocs(ShellAddToRecentDocsFlags flags, [MarshalAs(UnmanagedType.LPWStr)] string file);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern HRESULT SHILCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, out IntPtr ppIdl, ref uint rgflnOut);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern HRESULT SHCreateItemFromIDList(IntPtr pidl, [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreatePopupMenu();

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetMenuItemCount(IntPtr hMenu);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool GetMenuItemInfo(IntPtr hMenu, UInt32 uItem, bool fByPosition, [In, Out] MENUITEMINFO lpmii);

        [DllImport("user32.dll")]
        static extern int GetMenuString(IntPtr hMenu, uint uIDItem, [Out] StringBuilder lpString, int nMaxCount, uint uFlag);

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool DestroyMenu(IntPtr hMenu);

        public QuickAccessHandler()
        {
            this.QuickAccessMenuName = new List<string> { "快速访问", "Quick access", "Accès rapide" };
            this.FileExplorerMenuName = new List<string> { "文件资源管理器", "File Explorer", "Explorateur de fichiers" };
            this.QuickAccessRegistryKeyPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer";
            this.quickAccessShell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));

            this.frequentFolders = new Dictionary<string, string>();
            this.recentFiles = new Dictionary<string, string>();
            this.unspecificContent = new Dictionary<string, string>();

            GetQuickAccess();
        }

        private bool IsValidPath(string path)
        {
            return (File.Exists(path) ^ Directory.Exists(path));
        }

        public void AddQuickAccessMenuName(string menuName)
        {
            this.QuickAccessMenuName.Add(menuName);
        }

        public void AddFileExplorerMenuName(string menuName)
        {
            this.FileExplorerMenuName.Add(menuName);
        }

        // Refered from Adam
        // https://stackoverflow.com/questions/2488727/refresh-windows-explorer-in-win7
        private void RefreshFileExplorer()
        {
            Guid CLSID_ShellApplication = new Guid("13709620-C279-11CE-A49E-444553540000");
            Type shellApplicationType = Type.GetTypeFromCLSID(CLSID_ShellApplication, true);

            object shellApplication = Activator.CreateInstance(shellApplicationType);
            object windows = shellApplicationType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shellApplication, new object[] { });

            Type windowsType = windows.GetType();
            object count = windowsType.InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, null);
            for (int i = 0; i < (int)count; i++)
            {
                object item = windowsType.InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, new object[] { i });
                Type itemType = item.GetType();

                // only refresh windows explorers
                string itemName = (string)itemType.InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, item, null);
                foreach(string menuName in this.FileExplorerMenuName)
                {
                    if (itemName == menuName)
                    {
                        itemType.InvokeMember("Refresh", System.Reflection.BindingFlags.InvokeMethod, null, item, null);
                    }
                }
            }
        }

        // Refered from Simon Mourier
        // https://stackoverflow.com/questions/41048080/how-do-i-get-the-name-of-each-item-in-windows-10-quick-access-items-list-and-p
        private void GetQuickAccess()
        {
            var CLSID_HomeFolder = new Guid("679f85cb-0220-4080-b29b-5540cc05aab6");
            var quickAccess = this.quickAccessShell.Namespace("shell:::" + CLSID_HomeFolder.ToString("B"));

            foreach (var item in quickAccess.Items())
            {
                var grouping = (int)item.ExtendedProperty("System.Home.Grouping");
                switch (grouping)
                {
                    case (int)QuickAccessType.FrequentFolder:
                        if (this.frequentFolders.ContainsKey(item.name))
                        {
                            this.frequentFolders[item.name] = item.path;
                        }
                        else
                        {
                            this.frequentFolders.Add(item.name, item.path);
                        }
                        break;

                    case (int)QuickAccessType.RecentFile:
                        // File name will include its type, like myText.txt
                        if (this.recentFiles.ContainsKey(item.name))
                        {
                            this.recentFiles[item.name] = item.path;
                        }
                        else
                        {
                            this.recentFiles.Add(item.name, item.path);
                        }
                        break;

                    default:
                        if (this.unspecificContent.ContainsKey(item.name))
                        {
                            this.unspecificContent[item.name] = item.path;
                        }
                        else
                        {
                            this.unspecificContent.Add(item.name, item.path);
                        }
                        break;
                }
            }
        }

        public Dictionary<string, string> GetQuickAccessDict()
        {
            List<Dictionary<string, string>> quickAccessList = new List<Dictionary<string, string>> { this.frequentFolders, this.recentFiles, this.unspecificContent };
            Dictionary<string, string> quickAccessDict = new Dictionary<string, string>();

            foreach(Dictionary<string, string> listItem in quickAccessList)
            {
                foreach(KeyValuePair<string, string> quickAccessItem in listItem)
                {
                    quickAccessDict[quickAccessItem.Key] = quickAccessItem.Value;
                }
            }

            return quickAccessDict;
        }

        public Dictionary<string, string> GetFrequentFolders()
        {
            return this.frequentFolders;
        }

        public Dictionary<string, string> GetRecentFiles()
        {
            return this.recentFiles;
        }
        

        private bool IsPathInQuickAccess(string path)
        {
            if (IsValidPath(path))
            {
                var qucikAccessList = new List<Dictionary<string, string>> { this.frequentFolders, this.recentFiles };

                foreach (Dictionary<string, string> item in qucikAccessList)
                {
                    if (item.ContainsValue(path))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private bool IsKeywordInQuickAccess(string keyword)
        {
            var qucikAccessList = new List<Dictionary<string, string>> { this.frequentFolders, this.recentFiles };

            foreach (Dictionary<string, string> accessDict in qucikAccessList)
            {
                foreach (KeyValuePair<string, string> item in accessDict)
                {
                    if (item.Key.Contains(keyword) || item.Value.Contains(keyword))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public bool IsInQuickAccess(string data)
        {
            return (IsPathInQuickAccess(data) || IsKeywordInQuickAccess(data));
        }

        public bool RemoveFromQuickAccess(string path)
        {
            if (!IsValidPath(path)) return false;

            // declare variables
            HRESULT hr = HRESULT.E_FAIL;
            IntPtr pidlFull = IntPtr.Zero;
            uint rgflnOut = 0;
            string sPath = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";

            // Creates a pointer to an item identifier list (PIDL) from a path.
            hr = SHILCreateFromPath(sPath, out pidlFull, ref rgflnOut);
            if (hr == HRESULT.S_OK)
            {
                IntPtr pszName = IntPtr.Zero;
                IShellItem pShellItem;

                // Creates and initializes a Shell item object from a pointer to an item identifier list (PIDL).
                hr = SHCreateItemFromIDList(pidlFull, typeof(IShellItem).GUID, out pShellItem);
                if (hr == HRESULT.S_OK)
                {
                    // Get Windows Quick Access Folder
                    hr = pShellItem.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, out pszName);
                    if (hr == HRESULT.S_OK)
                    {
                        string sDisplayName = Marshal.PtrToStringUni(pszName);
                        // Console.WriteLine(string.Format("Folder Name : {0}", sDisplayName));
                        Marshal.FreeCoTaskMem(pszName);
                    }

                    IEnumShellItems pEnumShellItems = null;
                    IntPtr pEnum;
                    Guid BHID_EnumItems = new Guid("94f60519-2850-4924-aa5a-d15e84868039");
                    Guid BHID_SFUIObject = new Guid("3981e225-f559-11d3-8e3a-00c04f6837d5");

                    hr = pShellItem.BindToHandler(IntPtr.Zero, BHID_EnumItems, typeof(IEnumShellItems).GUID, out pEnum);
                    if (hr == HRESULT.S_OK)
                    {
                        pEnumShellItems = Marshal.GetObjectForIUnknown(pEnum) as IEnumShellItems;
                        IShellItem psi = null;
                        uint nFetched = 0;

                        while (HRESULT.S_OK == pEnumShellItems.Next(1, out psi, out nFetched) && nFetched == 1)
                        {
                            // Get Quick Access Item Absolute Path
                            pszName = IntPtr.Zero;
                            hr = psi.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out pszName);

                            if (hr == HRESULT.S_OK)
                            {
                                string sDisplayName = Marshal.PtrToStringUni(pszName);
                                // Console.WriteLine(string.Format("\tItem Name : {0}", sDisplayName));
                                Marshal.FreeCoTaskMem(pszName);

                                if (sDisplayName == path)
                                {
                                    IContextMenu pContextMenu;
                                    IntPtr pcm;
                                    hr = psi.BindToHandler(IntPtr.Zero, BHID_SFUIObject, typeof(IContextMenu).GUID, out pcm);
                                    if (hr == HRESULT.S_OK)
                                    {
                                        pContextMenu = Marshal.GetObjectForIUnknown(pcm) as IContextMenu;
                                        IntPtr hMenu = CreatePopupMenu();
                                        hr = pContextMenu.QueryContextMenu(hMenu, 0, 1, 0x7fff, 0);

                                        if (hr == HRESULT.S_OK)
                                        {
                                            // http://pastebin.fr/111943
                                            // Handle Quick Access File
                                            int nCommand = -1;
                                            int nNbItems = GetMenuItemCount(hMenu);
                                            for (int i = nNbItems - 1; i >= 0; i--)
                                            {
                                                MENUITEMINFO mii = new MENUITEMINFO();
                                                mii.fMask = (uint)(MenuItemInfo_fMask.MIIM_FTYPE |
                                                            MenuItemInfo_fMask.MIIM_ID |
                                                            MenuItemInfo_fMask.MIIM_SUBMENU |
                                                            MenuItemInfo_fMask.MIIM_DATA);

                                                // Check Whether Target File has an 'remove from quick access' menu option
                                                if (GetMenuItemInfo(hMenu, (uint)i, true, mii))
                                                {
                                                    StringBuilder menuName = new StringBuilder();
                                                    GetMenuString(hMenu, (uint)i, menuName, menuName.Capacity, (uint)MenuString_Pos.MF_BYPOSITION);

                                                    foreach(string item in this.QuickAccessMenuName)
                                                    {
                                                        if (menuName.ToString().Contains(item))
                                                        {
                                                            nCommand = (int)mii.wID;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }

                                            CMINVOKECOMMANDINFO ici = new CMINVOKECOMMANDINFO();
                                            ici.cbSize = Marshal.SizeOf(ici);
                                            ici.lpVerb = nCommand - 1;
                                            ici.nShow = (int)ShowWindowCommands.SW_NORMAL;
                                            ici.fMask = CMIC_MASK_ASYNCOK;

                                            hr = pContextMenu.InvokeCommand(ici);

                                            if (hr == HRESULT.S_OK)
                                            {
                                                break;
                                            }
                                        }
                                        DestroyMenu(hMenu);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            this.GetQuickAccess();

            return this.IsInQuickAccess(path);
        }

        public bool AddToQuickAccess(string path)
        {
            if (!IsValidPath(path)) return false;

            SHAddToRecentDocs(ShellAddToRecentDocsFlags.SHARD_PATHW, path);

            this.GetQuickAccess();

            return IsInQuickAccess(path);
        }

        public void ClearRecent()
        {
            SHAddToRecentDocs(ShellAddToRecentDocsFlags.SHARD_PIDL, null);
        }

        private void UpdateQuickAccessRegistryKey(string keyName, bool keyValue)
        {
            RegistryKey hklm = Registry.CurrentUser;
            RegistryKey hkExplorer = hklm.OpenSubKey(this.QuickAccessRegistryKeyPath, true);

            hkExplorer.SetValue(keyName, keyValue ? 1 : 0, RegistryValueKind.DWord);
        }

        public void IsShowQuickAccess(uint accessType, bool isShow)
        {
            UpdateQuickAccessRegistryKey((accessType == (uint)QuickAccessType.FrequentFolder ? "ShowRecent" : "ShowFrequent"), isShow);

            this.RefreshFileExplorer();
        }
    }
}
