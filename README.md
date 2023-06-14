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

QuickAccessShell 0.0.0.1
Copyright ©  2023 Steins.

ERROR(S):
  No verb selected.

  list       List current quick acess or supported languages.

  remove     Remove items from quick access.

  check      Check whether in quick access or show quick access or supported language.

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
  
  -u, --ui-culture              List system ui culture name.
```

### remove

```shell
  value pos. 0                  Targets to remove.
  
  -i, --internationalization    Add unsported language info
```

### check

```shell
  value pos. 0                    Targets to check.
  
  -q, --quick-access              Check whether in quick access.

  -s, --supported                 Check whether current system is supported

  -i, --internationalization      Add unsported language info
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
