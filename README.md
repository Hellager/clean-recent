# Clean Recent DLL

## `QuickAccessHandler`

```csharp
public class QuickAccess.QuickAccessHandler

```

Methods

| Type                         | Name                                                         | Summary |
| ---------------------------- | ------------------------------------------------------------ | ------- |
| `void`                       | AddFileExplorerMenuName(`String` languageCode, `String` menuName) |         |
| `void`                       | AddQuickAccessMenuName(`String` languageCode, `String` menuName) |         |
| `Boolean`                    | AddToQuickAccess(`String` path)                              |         |
| `void`                       | ClearRecent()                                                |         |
| `Dictionary<String, String>` | GetFrequentFolders()                                         |         |
| `Dictionary<String, String>` | GetQuickAccessDict()                                         |         |
| `Dictionary<String, String>` | GetRecentFiles()                                             |         |
| `Boolean`                    | IsInFileExplorerMenuName(`String` menuName)                  |         |
| `Boolean`                    | IsInQuickAccess(`String` data)                               |         |
| `Boolean`                    | IsInQuickAccessMenuName(`String` menuName)                   |         |
| `Boolean`                    | IsShowQuickAccess(`UInt32` accessType)                       |         |
| `Boolean`                    | IsSupportedFileExplorerLanguage(`String` languageCode)       |         |
| `Boolean`                    | IsSupportedQuickAccessLanguage(`String` languageCode)        |         |
| `Boolean`                    | RemoveFromQuickAccess(`String` path)                         |         |
| `void`                       | UpdateShowQuickAccess(`UInt32` accessType, `Boolean` isShow) |         |

