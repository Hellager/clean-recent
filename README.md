# Clean Recent DLL

## `QuickAccessHandler`

```csharp
public class QuickAccess.QuickAccessHandler

```

Methods

| Type | Name | Summary |
| --- | --- | --- |
| `Boolean` | IsInQuickAccessMenuName(`String` menuName) | check whether given menuName is in QuickAccessMenuName |
| `Boolean` | IsSupportedQuickAccessLanguage(`String` languageCode) | check whether given languageCode is supported language for QuickAccessMenuName |
| `void` | AddQuickAccessMenuName(`String` languageCode, `String` menuName) | add given languageCode with given menuName for removing item from quick access |
| `Boolean` | IsInFileExplorerMenuName(`String` menuName) | check whether given menuName is in FileExplorerMenuName |
| `Boolean` | IsSupportedFileExplorerLanguage(`String` languageCode) | check whether given languageCode is supported language for FileExplorerMenuName |
| `void` | AddFileExplorerMenuName(`String` languageCode, `String` menuName) | add given languageCode with given menuName for refreshing file explorer |
| `List<String>` | GetSupportLanguages() | get supported langugaes in the form of a string list <br /> support zh-CN, zh-TW, en-US, fr-FR, ru-RU by default |
| `Dictionary<String, String>` | GetQuickAccessDict() | get quick access items in the form of a dictionary <br /> <item name, item path> |
| `Dictionary<String, String>` | GetFrequentFolders() | get frequent folders in quick access in the form of a dictionary <br /> <folder name, folder path> |
| `Dictionary<String, String>` | GetRecentFiles() | get recent files in quick access in the form of a dictionary <br /> <file name, file path> |
| `Boolean` | IsInQuickAccess(`String` data) | check whether given string is in quick access |
| `Boolean` | AddToQuickAccess(`String` path) | add given path to quick access |
| `Boolean` | RemoveFromQuickAccess(`String` data) | removes given data from quick access |
| `void` | EmptyRecentFiles() | empty recent files |
| `void` | EmptyFrequentFolders() | empty frequent folders |
| `void` | EmptyQuickAccess() | empty quick access |
| `Boolean` | IsShowQuickAccess(`UInt32` accessType) | check whether show frequent folders or recent files |
| `void` | UpdateShowQuickAccess(`UInt32` accessType, `Boolean` isShow) | update whether show frequent folders or recent files <br /> will auto refresh file explorer if set language properly. |


## Notice

### ℹ️ Please check whether your windows system language is supported first!

Function `RemoveFromQuickAccess` relies on properly setting `QuickAccessMenuName` for removing items from quick access.

Function `UpdateShowQuickAccess` relies on properly setting `FileExplorerMenuName` for refreshing file explorer.

Language `zh-CN, zh-TW, en-US, fr-FR, ru-RU` is supported by default.

If your system language is not supported by default, you should add first using `AddQuickAccessMenuName` or `AddFileExplorerMenuName`.

You can get those menu names through [Microsoft: Find and open File Explorer](https://support.microsoft.com/en-us/windows/find-and-open-file-explorer-ef370130-1cca-9dc5-e0df-2f7416fe1cb1) and [Microsoft: Help in File Explorer](https://support.microsoft.com/en-us/windows/help-in-file-explorer-a2d33543-5242-788d-8994-b0be10ae5bca) by changing the url language code.

### ℹ️ `UpdateShowQuickAccess` is not the same in Folder Options!

Disable `Show recently used files in Quick Access` or `Show frequently used folders in Quick acess` in Folder Options will also clear your recently used files or frequently used folders.

Function `UpdateShowQuickAccess` won't clear any of your quick access items.

### ⚠️ `AddToQuickAccess` is not working properly!

You can run the tests to check, it has been temporarily commented out.

I have asked about it but no progress for now: [function SHAddToRecentDocs seems not working on win10 for folders](https://learn.microsoft.com/en-us/answers/questions/1115223/function-shaddtorecentdocs-seems-not-working-on-wi.html).

If you know how to solve it or have any good ideas, please let me know to make it working.

