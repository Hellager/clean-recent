# Clean Recent DLL

## `QuickAccessHandler`

```csharp
public class QuickAccess.QuickAccessHandler

```

Methods

| Type | Name | Summary |
| --- | --- | --- |
| `Boolean` | IsInQuickAccessMenuName(`String` menuName) | check whether given menuName is in QuickAccessMenuName |
| `Boolean` | IsSupportedQuickAccessLanguage(`String` languageCode) | check whether given languageCode is supported language for QuickAccessMenuName, like `en-us` |
| `void` | AddQuickAccessMenuName(`String` languageCode, `String` menuName) | add given languageCode with given menuName to support different language system |
| `Boolean` | IsInFileExplorerMenuName(`String` menuName) | check whether given menuName is in FileExplorerMenuName |
| `Boolean` | IsSupportedFileExplorerLanguage(`String` languageCode) | check whether given languageCode is supported language for FileExplorerMenuName |
| `void` | AddFileExplorerMenuName(`String` languageCode, `String` menuName) | add given languageCode with given menuName to support different language system |
| `List<String>` | GetSupportLanguages() | get supported langugaes in the form of a string list <br /> support `zh-CN, zh-TW, en-US, fr-FR, ru-RU` by default |
| `Dictionary<String, String>` | GetQuickAccessDict() | get quick access items in the form of Dictionary <br /> `<item path, item name>` |
| `Dictionary<String, String>` | GetFrequentFolders() | get frequent folders in quick access in the form of Dictionary  <br /> `<folder path, folder name>` |
| `Dictionary<String, String>` | GetRecentFiles() | get recent files in quick access in the form of Dictionary  <br /> `<file path, file name>` |
| `Boolean` | IsInQuickAccess(`String` data) | check whether given data is in quick access |
| `Boolean` | AddToQuickAccess(`String` path) | add given path to quick access |
| `Boolean` | RemoveFromQuickAccess(`String` data) | remove given data from quick access |
| `void` | EmptyRecentFiles() | empty recent files |
| `void` | EmptyFrequentFolders() | empty frequent folders |
| `void` | EmptyQuickAccess() | empty quick access |
| `Boolean` | IsAdminPrivilege() | check whether current user has admin priviledge |
| `Int32` | GetQuickAccessRegistryKey(`String` keyName) | get registry key value about quick access |
| `Boolean` | IsShowQuickAccess(`UInt32` accessType) | check whether show quick access related <br />0 for all quick access related <br />1 for Frequent folders<br />2 for Recent files<br />3 for Quick access |
| `void` | UpdateShowQuickAccess(`UInt32` accessType, `Boolean` isShow) | update whether show quick access related <br />0 for all quick access related <br />1 for Frequent folders<br />2 for Recent files<br />3 for Quick access |


## Notice

### ℹ️ Please check whether your windows system language is supported first!

Function `RemoveFromQuickAccess` relies on properly setting `QuickAccessMenuName` for removing items from quick access.

Function `UpdateShowQuickAccess` relies on properly setting `FileExplorerMenuName` for refreshing file explorer.

Language `zh-CN, zh-TW, en-US, fr-FR, ru-RU` is supported by default.

If your system language is not supported by default, you should add first using `AddQuickAccessMenuName` or `AddFileExplorerMenuName`.

You can get those menu names through [Microsoft: Find and open File Explorer](https://support.microsoft.com/en-us/windows/find-and-open-file-explorer-ef370130-1cca-9dc5-e0df-2f7416fe1cb1) and [Microsoft: Help in File Explorer](https://support.microsoft.com/en-us/windows/help-in-file-explorer-a2d33543-5242-788d-8994-b0be10ae5bca) by changing the url language code.

### ℹ️ `UpdateShowQuickAccess` is not the same in `Folder Options` !

Disable `Show recently used files in Quick Access` or `Show frequently used folders in Quick acess` in Folder Options will also clear your recently used files or frequently used folders.

Function `UpdateShowQuickAccess` won't clear any of your quick access items.

### ℹ️ Show/hide `Quick access` in file explorer requires administrator privileges!

Since the registry key about show/hide `Quick access` is in `HKEY_LOCAL_MACHINE`, we are good to get the key value, but need administrator privileges to change it.

You have to check whether the current user has administrator privileges before calling `UpdateShowQuickAccess` to show/hide `Quick access`.

### ⚠️ Change `Quick access` show status may mix your recent files and frequent folders!

Normally in `Quick access`, we will see `Frequent folders` and `Recent files` are in separate state like accordion.

After changing the `Quick access` show status, you may see the `Frequent folders` and `Recent files` are mixed with each other without any separation, usually back to normal after restart the system.

Haven't figured out why this happens, if you know about it, please help to fix it.

### ⚠️ `AddToQuickAccess` is not working properly!

You can run the tests to check, it has been temporarily commented out.

I have asked about it but no progress for now: [function SHAddToRecentDocs seems not working on win10 for folders](https://learn.microsoft.com/en-us/answers/questions/1115223/function-shaddtorecentdocs-seems-not-working-on-wi.html).

If you know how to solve it or have any good ideas, please let me know to make it working.

