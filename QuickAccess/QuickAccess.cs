using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
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
    /// <summary>
    /// Class <c>QuickAccessHandler</c> Handle windows quick access list.
    /// </summary>
    public class QuickAccessHandler
    {
        /// <summary>
        /// Instance variable <c>quickAccessShell</c> A shell instance to handle various actions.
        /// </summary>
        private dynamic quickAccessShell;

        // check language code at https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/faq/languages

        /// <summary>
        /// Instance variable <c>QuickAccessMenuName</c> Store the quick access menu name in windows. In en-US, it should be "Quick access". <br /> If not set properly, will fail to remove item from quick access list.
        /// </summary>
        private Dictionary<string, string> QuickAccessMenuName;

        /// <summary>
        /// Instance variable <c>FileExplorerMenuName</c> Store the file explorer menu name in windows. <br /> In en-US, it should be "File Explorer"
        /// </summary>
        private Dictionary<string, string> FileExplorerMenuName;

        /// <summary>
        /// Instance variable <c>QuickAccessRegistryKeyPath</c> Represents the regisrty key path about quick access. <br /> Should be "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer".
        /// </summary>
        private string QuickAccessRegistryKeyPath;

        /// <summary>
        /// Instance variable <c>SEE_MASK_ASYNCOK</c> fMask value for "unpinfromhome".
        /// </summary>
        private const int SEE_MASK_ASYNCOK = 0x00100000;

        /// <summary>
        /// Instance variable <c>SEE_MASK_ASYNCOK</c> fMask value for "remove from quick access".
        /// </summary>
        private const int CMIC_MASK_ASYNCOK = SEE_MASK_ASYNCOK;

        // <resource name, resource path>

        /// <summary>
        /// Instance variable <c>frequentFolders</c> Store frequent folders. key: folder path, value: folder name.
        /// </summary>
        private Dictionary<string, string> frequentFolders;

        /// <summary>
        /// Instance variable <c>recentFiles</c> Store recent files. key: file path, value: file name.
        /// </summary>
        private Dictionary<string, string> recentFiles;

        /// <summary>
        /// Instance variable <c>unspecificContent</c> Unknown type items in quick access list. key: item path, value: item name.
        /// </summary>
        private Dictionary<string, string> unspecificContent;

        ///// <summary>
        ///// Instance variable <c>QuickAccessType</c> Enumeration of quick access items type: FrequentFolder, RecentFile, Undefined.
        ///// </summary>
        private enum QuickAccessType
        {
            FrequentFolder = 1,
            RecentFile_Win10 = 2,
            RecentFile_Win11 = 3,
        };

        [DllImport("Shell32.dll", BestFitMapping = false, ThrowOnUnmappableChar = true)]
        private static extern void SHAddToRecentDocs(ShellAddToRecentDocsFlags flags, [MarshalAs(UnmanagedType.LPWStr)] string file);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern HRESULT SHILCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, out IntPtr ppIdl, ref uint rgflnOut);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern HRESULT SHCreateItemFromIDList(IntPtr pidl, [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreatePopupMenu();

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetMenuItemCount(IntPtr hMenu);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool GetMenuItemInfo(IntPtr hMenu, UInt32 uItem, bool fByPosition, [In, Out] MENUITEMINFO lpmii);

        [DllImport("user32.dll")]
        private static extern int GetMenuString(IntPtr hMenu, uint uIDItem, [Out] StringBuilder lpString, int nMaxCount, uint uFlag);

        [DllImport("User32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DestroyMenu(IntPtr hMenu);

        public QuickAccessHandler()
        {
            this.QuickAccessMenuName = new Dictionary<string, string>
            {
                {"zh-CN", "快速访问"},
                {"zh-TW", "快速存取" },
                {"en-US", "Quick access"},
                {"fr-FR", "Accès rapide"},
                {"ru-RU", "доступа" }
            };
            this.FileExplorerMenuName = new Dictionary<string, string>
            {
                {"zh-CN", "文件资源管理器"},
                {"zh-TW", "檔案總管" },
                {"en-US", "File Explorer"},
                {"fr-FR", "Explorateur de fichiers"},
                {"ru-RU", "проводник" }
            };
            this.QuickAccessRegistryKeyPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer";
            this.quickAccessShell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));

            this.frequentFolders = new Dictionary<string, string>();
            this.recentFiles = new Dictionary<string, string>();
            this.unspecificContent = new Dictionary<string, string>();

            GetQuickAccess();
        }

        /// <summary>
        /// This method checks whether given path is exists.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given path exists whether it's file or folder, else false.
        /// </returns>
        /// <param><c>path</c> is the given path string.</param>
        private bool IsValidPath(string path)
        {
            return (File.Exists(path) ^ Directory.Exists(path));
        }

        /// <summary>
        /// This method checks whether given menuName is in QuickAccessMenuName dict.
        /// </summary>
        /// (<paramref name="menuName"/>).
        /// <returns>
        /// True if the given menuName is in QuickAccessMenuName dict, else false.
        /// </returns>
        /// <param><c>menuName</c> is the given menuName string.</param>
        public bool IsInQuickAccessMenuName(string menuName)
        {
            foreach (string value in this.QuickAccessMenuName.Values)
            {
                if (menuName == value)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// This method checks whether given languageCode is supported language about QuickAccessMenuName.
        /// </summary>
        /// (<paramref name="languageCode"/>).
        /// <returns>
        /// True if the given languageCode is supported language about QuickAccessMenuName, else false.
        /// </returns>
        /// <param><c>languageCode</c> is the given languageCode string. For example 'en-US'.</param>
        public bool IsSupportedQuickAccessLanguage(string languageCode)
        {
            foreach (string key in this.QuickAccessMenuName.Keys)
            {
                if (languageCode == key)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// This method adds given languageCode with given menuName to QuickAccessMenuName dict.
        /// </summary>
        /// (<paramref name="languageCode"/>,<paramref name="menuName"/>).
        /// <param><c>languageCode</c> is the given languageCode string. For example 'en-US'.</param>
        /// <param><c>menuName</c> is the given menuName string.</param>
        public void AddQuickAccessMenuName(string languageCode, string menuName)
        {
            if (IsSupportedQuickAccessLanguage(languageCode) || IsInQuickAccessMenuName(menuName)) return;

            this.QuickAccessMenuName.Add(languageCode, menuName);  
        }

        /// <summary>
        /// This method checks whether given menuName is in FileExplorerMenuName dict.
        /// </summary>
        /// (<paramref name="menuName"/>).
        /// <returns>
        /// True if the given menuName is in FileExplorerMenuName dict, else false.
        /// </returns>
        /// <param><c>menuName</c> is the given menuName string.</param>
        public bool IsInFileExplorerMenuName(string menuName)
        {
            foreach (string value in this.FileExplorerMenuName.Values)
            {
                if (menuName == value)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// This method checks whether given languageCode is supported language about FileExplorerMenuName.
        /// </summary>
        /// (<paramref name="languageCode"/>).
        /// <returns>
        /// True if the given languageCode is supported language about FileExplorerMenuName, else false.
        /// </returns>
        /// <param><c>languageCode</c> is the given languageCode string. For example 'en-US'.</param>
        public bool IsSupportedFileExplorerLanguage(string languageCode)
        {
            foreach (string key in this.FileExplorerMenuName.Keys)
            {
                if (languageCode == key)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// This method adds given languageCode with given menuName to FileExplorerMenuName dict.
        /// </summary>
        /// (<paramref name="languageCode"/>,<paramref name="menuName"/>).
        /// <param><c>languageCode</c> is the given languageCode string. For example 'en-US'.</param>
        /// <param><c>menuName</c> is the given menuName string.</param>
        public void AddFileExplorerMenuName(string languageCode, string menuName)
        {
            if (IsSupportedFileExplorerLanguage(languageCode) || IsInFileExplorerMenuName(menuName)) return;

            this.FileExplorerMenuName.Add(languageCode, menuName);
        }

        /// <summary>
        /// This method gets the supported langugaes. Support zh-CN, zh-TW, en-US, fr-FR, ru-RU by default.
        /// </summary>
        /// <returns>
        /// True if the given languageCode is supported language about FileExplorerMenuName, else false.
        /// </returns>
        public List<string> GetSupportLanguages()
        {
            List<string> supportQuickAccessLanguage = new List<string>(this.QuickAccessMenuName.Keys);
            List<string> supportFileExplorerLanguage = new List<string>(this.FileExplorerMenuName.Keys);

            return supportQuickAccessLanguage.Intersect(supportFileExplorerLanguage).ToList();
        }

        /// <summary>
        /// This method refreshes the file explorer. Should correctly set FileExplorerMenuName dict to support different language. <br />
        /// Refered from @Adam https://stackoverflow.com/questions/2488727/refresh-windows-explorer-in-win7
        /// </summary>
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
                foreach (string menuName in this.FileExplorerMenuName.Values)
                {
                    if (itemName == menuName)
                    {
                        itemType.InvokeMember("Refresh", System.Reflection.BindingFlags.InvokeMethod, null, item, null);
                    }
                }
            }
        }

        /// <summary>
        /// This method gets the quick access items throuth shell instance and stores them in frequentFolders/recentFiles/unspecificContent. <br />
        /// Refered from @Simon Mourier https://stackoverflow.com/questions/41048080/how-do-i-get-the-name-of-each-item-in-windows-10-quick-access-items-list-and-p
        /// </summary>
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
                        if (this.frequentFolders.ContainsKey(item.path))
                        {
                            this.frequentFolders[item.path] = item.name;
                        }
                        else
                        {
                            this.frequentFolders.Add(item.path, item.name);
                        }
                        break;

                    case (int)QuickAccessType.RecentFile_Win10:
                        if (this.recentFiles.ContainsKey(item.path))
                        {
                            this.recentFiles[item.path] = item.name;
                        }
                        else
                        {
                            this.recentFiles.Add(item.path, item.name);
                        }
                        break;

                    case (int)QuickAccessType.RecentFile_Win11:
                        if (this.recentFiles.ContainsKey(item.path))
                        {
                            this.recentFiles[item.path] = item.name;
                        }
                        else
                        {
                            this.recentFiles.Add(item.path, item.name);
                        }
                        break;

                    default:
                        if (this.unspecificContent.ContainsKey(item.path))
                        {
                            this.unspecificContent[item.path] = item.name;
                        }
                        else
                        {
                            this.unspecificContent.Add(item.path, item.name);
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// This method gets the quick access items in the form of a dictionary. <item name, item path>
        /// </summary>
        /// <returns>
        /// The quick access items dictionary combines frequentFolders/recentFiles/unspecificContent.
        /// </returns>
        public Dictionary<string, string> GetQuickAccessDict()
        {
            GetQuickAccess();

            List<Dictionary<string, string>> quickAccessList = new List<Dictionary<string, string>> { this.frequentFolders, this.recentFiles, this.unspecificContent };
            Dictionary<string, string> quickAccessDict = new Dictionary<string, string>();

            foreach (Dictionary<string, string> listItem in quickAccessList)
            {
                foreach (KeyValuePair<string, string> quickAccessItem in listItem)
                {
                    quickAccessDict[quickAccessItem.Key] = quickAccessItem.Value;
                }
            }

            return quickAccessDict;
        }

        /// <summary>
        /// This method gets the frequent folders in quick access in the form of a dictionary. <folder name, folder path>
        /// </summary>
        /// <returns>
        /// The frequentFolders dictionary.
        /// </returns>
        public Dictionary<string, string> GetFrequentFolders()
        {
            GetQuickAccess();

            return this.frequentFolders;
        }

        /// <summary>
        /// This method gets the recent files in quick access in the form of a dictionary. <file name, file path>
        /// </summary>
        /// <returns>
        /// The recentFiles dictionary.
        /// </returns>
        public Dictionary<string, string> GetRecentFiles()
        {
            GetQuickAccess();

            return this.recentFiles;
        }

        /// <summary>
        /// This method checks whether given path is in quick access.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given path is valid and in quick access, else false.
        /// </returns>
        /// <param><c>path</c> is the given path string,</param>
        private bool IsPathInQuickAccess(string path)
        {
            GetQuickAccess();

            if (IsValidPath(path))
            {
                var qucikAccessList = new List<Dictionary<string, string>> { this.frequentFolders, this.recentFiles, this.unspecificContent };

                foreach (Dictionary<string, string> item in qucikAccessList)
                {
                    if (item.ContainsKey(path))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// This method checks whether given keyword is in quick access.
        /// </summary>
        /// (<paramref name="keyword"/>).
        /// <returns>
        /// True if the given keyword is in quick access, else false.
        /// </returns>
        /// <param><c>keyword</c> is the given keyword string.</param>
        private bool IsKeywordInQuickAccess(string keyword)
        {
            GetQuickAccess();

            var qucikAccessList = new List<Dictionary<string, string>> { this.frequentFolders, this.recentFiles, this.unspecificContent };

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

        /// <summary>
        /// This method checks whether given string is in quick access.
        /// </summary>
        /// (<paramref name="data"/>).
        /// <returns>
        /// True if the given string is in quick access, else false.
        /// </returns>
        /// <param><c>data</c> is the given string, whether it's full path or just keyword..</param>
        public bool IsInQuickAccess(string data)
        {
            return (IsPathInQuickAccess(data) || IsKeywordInQuickAccess(data));
        }

        /// <summary>
        /// This method pins folder to quick access with runspace factory.
        /// Using try catch to adapt tauri app.
        /// https://stackoverflow.com/questions/36739317/programatically-pin-unpin-the-folder-from-quick-access-menu-in-windows-10
        /// </summary>
        /// (<paramref name="path"/>).
        /// <param><c>path</c> is the given folder path.</param>
        private void PinFolderToQuickAccess(string path)
        {
            try
            {
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    runspace.Open();

                    var ps = PowerShell.Create();
                    var shellApplication =
                        ps.AddCommand("New-Object").AddParameter("ComObject", "shell.application").Invoke();

                    dynamic nameSpace = shellApplication.FirstOrDefault()?.Methods["NameSpace"].Invoke(path);
                    nameSpace?.Self.InvokeVerb("pintohome");
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Err: {0}", e);
            }
        }

        /// <summary>
        /// This method adds given path to quick access.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given string is valid and in quick access after adding, else false.
        /// </returns>
        /// <param><c>path</c> is the given path string.</param>
        public bool AddToQuickAccess(string path)
        {
            if (!IsValidPath(path)) return false;
            if (IsPathInQuickAccess(path)) return true;

            if (File.Exists(path))
            {
                SHAddToRecentDocs(ShellAddToRecentDocsFlags.SHARD_PATHW, path);
            }
            else if (Directory.Exists(path))
            {
                PinFolderToQuickAccess(path);
            }
            
            return IsPathInQuickAccess(path);
        }

        /// <summary>
        /// This method unpins folder from quick access with runspace factory.
        /// Using try catch to adapt tauri app.
        /// https://stackoverflow.com/questions/36739317/programatically-pin-unpin-the-folder-from-quick-access-menu-in-windows-10
        /// </summary>
        /// (<paramref name="path"/>).
        /// <param><c>path</c> is the given folder path.</param>
        private void UnpinFolderFromQuickAccess(string path)
        {
            try
            {
                using (var runspace = RunspaceFactory.CreateRunspace())
                {
                    runspace.Open();
                    var ps = PowerShell.Create();
                    var removeScript =
                        $"((New-Object -ComObject shell.application).Namespace(\"shell:::{{679f85cb-0220-4080-b29b-5540cc05aab6}}\").Items() | Where-Object {{ $_.Path -EQ \"{path}\" }}).InvokeVerb(\"unpinfromhome\")";

                    ps.AddScript(removeScript);
                    ps.Invoke();
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Err: {0}", e);
            }
        }

        /// <summary>
        /// This method removes given full path item from quick access.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given full path item is not in quick access after removing, else false.
        /// </returns>
        /// <param><c>path</c> is the given path string.</param>
        private bool RemoveFromQuickAccessWithFullPath(string path)
        {
            if (!IsValidPath(path)) return false;

            if (Directory.Exists(path))
            {
                UnpinFolderFromQuickAccess(path);

                if (!IsPathInQuickAccess(path)) return true;
            }

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

                                                    foreach (string item in this.QuickAccessMenuName.Values)
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

            return !IsPathInQuickAccess(path);
        }

        /// <summary>
        /// This method removes given keyword item from quick access.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given keyword item is not in quick access after removing, else false.
        /// </returns>
        /// <param><c>keyword</c> is the given keyword string.</param>
        private bool RemoveFromQuickAccessWithKeyword(string keyword)
        {
            var CurrentQuickAccessDict = GetQuickAccessDict();

            foreach(var item in CurrentQuickAccessDict)
            {
                if (item.Key.Contains(keyword))
                {
                    RemoveFromQuickAccessWithFullPath(item.Key);
                }
            }

            return !IsKeywordInQuickAccess(keyword);
        }

        /// <summary>
        /// This method removes given data from quick access.
        /// </summary>
        /// (<paramref name="path"/>).
        /// <returns>
        /// True if the given data is not in quick access after removing, else false.
        /// </returns>
        /// <param><c>data</c> is the given data string.</param>
        public bool RemoveFromQuickAccess(string data)
        {
            if (IsValidPath(data))
            {
                RemoveFromQuickAccessWithFullPath(data);
            } 
            else
            {
                RemoveFromQuickAccessWithKeyword(data);
            }

            return !IsInQuickAccess(data);
        }

        /// <summary>
        /// This method clears the recent files.
        /// </summary>
        public void EmptyRecentFiles()
        {
            SHAddToRecentDocs(ShellAddToRecentDocsFlags.SHARD_PIDL, null);
        }

        /// <summary>
        /// This method clears the frequent folders.
        /// </summary>
        public void EmptyFrequentFolders()
        {
            var CurrentFrequentFolders = this.GetFrequentFolders();

            foreach(var item in CurrentFrequentFolders)
            {
                RemoveFromQuickAccessWithFullPath(item.Key);
            }
        }

        /// <summary>
        /// This method clears the quick access.
        /// </summary>
        public void EmptyQuickAccess()
        {
            EmptyRecentFiles();

            EmptyFrequentFolders();
        }

        /// <summary>
        /// This method checks whether current user has system administrator privilege.
        /// </summary>
        /// <returns>
        /// True if current user has system administrator privilege, else false.
        /// </returns>
        public bool IsAdminPrivilege()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        /// <summary>
        /// This method updates the registry key value about quick access.
        /// If keyName is 'HubMode', should check admin privilege first.
        /// </summary>
        /// (<paramref name="keyName"/>,<paramref name="keyValue"/>).
        /// <param><c>keyName</c> is the quick access registry key name.</param>
        /// <param><c>keyValue</c> is new value about given registry key.</param>
        private void UpdateQuickAccessRegistryKey(string keyName, bool keyValue)
        {
            if (keyName == "HubMode" && !IsAdminPrivilege()) return;

            RegistryKey hklm = (keyName == "HubMode") ? Registry.LocalMachine : Registry.CurrentUser;
            RegistryKey hkExplorer = hklm.OpenSubKey(this.QuickAccessRegistryKeyPath, true);

            if (keyName == "HubMode")
            {
                hkExplorer.SetValue(keyName, keyValue ? 0 : 1, RegistryValueKind.DWord);
                return;
            }

            hkExplorer.SetValue(keyName, keyValue ? 1 : 0, RegistryValueKind.DWord);
        }

        /// <summary>
        /// This method gets the registry key value about quick access.
        /// </summary>
        /// (<paramref name="keyName"/>).
        /// <returns>
        /// The given registry key's value, -1 for no such key.
        /// </returns>
        /// <param><c>keyName</c> is the quick access registry key name.</param>
        public int GetQuickAccessRegistryKey(string keyName)
        {
            /// In a x86 machine or program, key 'HubMode' won't be about to get, which means it will always return -1.
            /// https://stackoverflow.com/questions/13324920/regedit-shows-keys-that-are-not-listed-using-getsubkeynames
            RegistryKey hklm = (keyName == "HubMode") ? Registry.LocalMachine : Registry.CurrentUser;
            RegistryKey hkExplorer = hklm.OpenSubKey(this.QuickAccessRegistryKeyPath);

            if (!hkExplorer.GetValueNames().Contains(keyName))
            {
                return -1;
            }

            var currentValue = (int)hkExplorer.GetValue(keyName);

            return currentValue;
        }

        /// <summary>
        /// This method checks whether show frequent folders or recent files.
        /// </summary>
        /// (<paramref name="accessType"/>).
        /// <param><c>keyName</c> is the quick access type. 0 for all type, 1 for frequent folder, 2 for recent files, 3 for side menu quick access, other returns false.</param>
        public bool IsShowQuickAccess(uint accessType)
        {
            bool isShow = false;
            int registrykeyValue;
            string[] registryKey = new string[3] { "ShowFrequent", "ShowRecent", "HubMode" };

            if (accessType == 0)
            {
                int ShowFrequentValue = GetQuickAccessRegistryKey("ShowFrequent");
                int ShowRecentValue = GetQuickAccessRegistryKey("ShowRecent");
                int HubModeValue = GetQuickAccessRegistryKey("HubMode");

                isShow = ((ShowFrequentValue == 0) || (ShowRecentValue == 0) || (HubModeValue == 1)) ? false : true;
            }
            else if (accessType > 0 && accessType <= 3)
            {
                registrykeyValue = GetQuickAccessRegistryKey(registryKey[accessType - 1]);

                // If no such key about quick access, system will show quick access by default.
                if (registrykeyValue == -1) return true;

                // For key 'HubMode' about side menu quick access, 0 for showing, 1 for hiding
                if (accessType == 3)
                {
                    isShow = registrykeyValue > 0 ? false : true;
                }
                else
                {
                    isShow = registrykeyValue > 0 ? true : false;
                }
            }
            else
            {
                isShow = false;
            }

            return isShow;
        }

        /// <summary>
        /// This method updates whether show frequent folders or recent files and will auto refresh file explorer if set language properly.
        /// </summary>
        /// (<paramref name="accessType"/>, <paramref name="isShow"/>).
        /// <param><c>keyName</c> is the quick access type. 0 for all type, 1 for frequent folder, 2 for recent files, 3 for side menu quick access, other returns false.</param>
        /// <param><c>isShow</c> whether show.</param>
        public void UpdateShowQuickAccess(uint accessType, bool isShow)
        {
            switch(accessType)
            {
                case 0:
                    UpdateQuickAccessRegistryKey("ShowFrequent", isShow);
                    UpdateQuickAccessRegistryKey("ShowRecent", isShow);
                    UpdateQuickAccessRegistryKey("HubMode", isShow);
                    break;

                case 1:
                    UpdateQuickAccessRegistryKey("ShowFrequent", isShow);
                    break;

                case 2:
                    UpdateQuickAccessRegistryKey("ShowRecent", isShow);
                    break;

                case 3:
                    UpdateQuickAccessRegistryKey("HubMode", isShow);
                    break;

                default:
                    break;
            }    

            RefreshFileExplorer();
        }
    }
}
