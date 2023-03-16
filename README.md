# QuickAccessShell

QuickAccessShell is a c# console application based on .net framework for handling windows quick access.

This is customized for Tauri sidecar with JSON output.

## Installation

### 1. From release page

Check [release page](https://github.com/Hellager/clean-recent/releases) and download `QuickAccessShell.exe`.

### 2. Build it yourself

```c
$ git clone https://github.com/Hellager/clean-recent.git
$ git fetch origin shell
$ git checkout -b shell origin/shell
```

Open project in Visual Studio and build solution.

## Usage

```shell
$ .\QuickAccessShell.exe

QuickAccessShell 1.0
Copyright (c) 2023 Steins

ERROR(S):
  No verb selected.

  list       List current quick acess or supported language.
  add        Add file or folder to quick access.
  remove     Remove items from quick access.
  show       Show/Hide quick access related.
  check      Check whether in quick access or show quick access or supported
             language
  empty      Empty quick access.
  help       Display more information on a specific command.
  version    Display version information.
```



## Verbs & Options

### list

```shell
  -a, --all                     List both recent files and frequent folders.

  -r, --recent-files            List recent files.

  -f, --frequent-folders        List frequent folders.

  -i, --internationalization    List supported language
```

### add

```shell
  value pos. 0    Targets to add.
```

### remove

```shell
  value pos. 0                  Targets to remove.
  
  -i, --internationalization    To support unsupported language.
```

### show

```shell
  value pos. 0                    (Default: true) Determine whether show or hide.
  
  -a, --all                       Show/Hide all quick access related.

  -r, --recent-files              Show/Hide recent files.

  -f, --frequent-folders          Show/Hide frequent folders.

  -s, --side-menu-quick-access    Show/Hide side menu quick access.
```

### check

```shell
  value pos. 0                    Targets to check.
  
  -q, --quick-access              Check whether in quick access.

  -a, --all                       Check whether show all quick access related.

  -r, --recent-files              Check whether show recent files.

  -f, --frequent-folders          Check whether show frequent folders.

  -s, --side-menu-quick-access    Show/Hide side menu quick access.

  -i, --internationalization      Check whether supported language
```

### empty

```shell
  -a, --all                 Empty all quick access items.

  -r, --recent-files        Empty recent files.

  -f, --frequent-folders    Empty frequent folders.
```



## Contributing

Pull requests are welcome. Since this console application is customized for Tauri sidecar, if you want to change the ouput, you'd better build your own console application.



## License

[Apache License 2.0](https://choosealicense.com/licenses/apache-2.0/)
