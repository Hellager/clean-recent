using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Newtonsoft.Json;
using CommandLine;
using CommandLine.Text;
using QuickAccess;

namespace QuickAccessShell
{
    internal class Program
    {
        [Verb("list", HelpText = "List current quick acess or supported language.")]
        class ListOptions
        {
            [Option('a', "all", Required = false, HelpText = "List both recent files and frequent folders.")]
            public bool IsListAllAccess { get; set; }

            [Option('r', "recent-files", Required = false, HelpText = "List recent files.")]
            public bool IsListRecentFiles { get; set; }

            [Option('f', "frequent-folders", Required = false, HelpText = "List frequent folders.")]
            public bool IsListFrequentFolders { get; set; }

            [Option('i', "internationalization", Required = false, HelpText = "List supported language")]
            public bool IsListSupportedLanguage { get; set; }
        }

        [Verb("add", HelpText = "Add file or folder to quick access.")]
        class AddOptions
        {
            [Value(0, HelpText = "Targets to add.")]
            public IEnumerable<string> AddItems { get; set; }
        }


        [Verb("remove", HelpText = "Remove items from quick access.")]
        class RemoveOptions
        {
            [Value(0, HelpText = "Targets to remove.")]
            public IEnumerable<string> RemoveItems { get; set; }

            [Option('i', "internationalization", Required = false, HelpText = "To support unsupported language.")]
            public string I18nQuickAccess { get; set; }
        }

        [Verb("show", HelpText = "Show/Hide quick access related.")]
        class ShowOptions
        {
            [Value(0, Default = true, HelpText = "Determine whether show or hide.")]
            public bool IsShow { get; set; }

            [Option('a', "all", Required = false, HelpText = "Show/Hide all quick access related.")]
            public bool IsShowAll { get; set; }

            [Option('r', "recent-files", Required = false, HelpText = "Show/Hide recent files.")]
            public bool IsShowRecentFiles { get; set; }

            [Option('f', "frequent-folders", Required = false, HelpText = "Show/Hide frequent folders.")]
            public bool IsShowFrequentFolders { get; set; }

            [Option('s', "side-menu-quick-access", Required = false, HelpText = "Show/Hide side menu quick access.")]
            public bool IsShowSideMenuQuickAccess { get; set; }
        }

        [Verb("check", HelpText = "Check whether in quick access or show quick access or supported language")]
        class CheckOptions
        {
            [Value(0, HelpText = "Targets to check.")]
            public IEnumerable<string> Target { get; set; }

            [Option('q', "quick-access", Required = false, HelpText = "Check whether in quick access.")]
            public bool IsInQuickAccess { get; set; }

            [Option('a', "all", Required = false, HelpText = "Check whether show all quick access related.")]
            public bool IsShowAll { get; set; }

            [Option('r', "recent-files", Required = false, HelpText = "Check whether show recent files.")]
            public bool IsShowRecentFiles { get; set; }

            [Option('f', "frequent-folders", Required = false, HelpText = "Check whether show frequent folders.")]
            public bool IsShowFrequentFolders { get; set; }

            [Option('s', "side-menu-quick-access", Required = false, HelpText = "Show/Hide side menu quick access.")]
            public bool IsShowSideMenuQuickAccess { get; set; }

            [Option('i', "internationalization", Required = false, HelpText = "Check whether supported language")]
            public bool IsCheckSupportedLanguage { get; set; }
        }

        [Verb("clean", HelpText = "Empty quick access.")]
        class CleanOptions
        {
            [Option('a', "all", Required = false, HelpText = "Empty all quick access items.")]
            public bool IsEmptyAll { get; set; }

            [Option('r', "recent-files", Required = false, HelpText = "Empty recent files.")]
            public bool IsEmptyRecentFiles { get; set; }

            [Option('f', "frequent-folders", Required = false, HelpText = "Empty frequent folders.")]
            public bool IsEmptyFrequentFolders { get; set; }
        }

        //[Verb("test", HelpText = "Test some feature.")]
        //class TestOptions
        //{
        //    [Option('t', "test", Required = false, HelpText = "Test with option.")]
        //    public bool IsTestOption { get; set; }
        //}

        private class OutputData
        {
            public ArrayList Data { get; set; }

            public string Type { get; set; }

            public ArrayList Debug { get; set; }
        }

        private static readonly QuickAccessHandler _handler = new QuickAccessHandler();

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            var parserResult = new Parser(with => with.HelpWriter = null).ParseArguments<ListOptions, AddOptions, RemoveOptions, ShowOptions, CheckOptions, CleanOptions>(args);

            parserResult.MapResult(
                    (ListOptions _option) => HandleListOptions(_option),
                    (AddOptions _option) => HandleAddOptions(args, _option),
                    (RemoveOptions _option) => HandleRemoveOptions(args, _option),
                    (ShowOptions _option) => HandleShowOptions(_option),
                    (CheckOptions _option) => HandleCheckOptions(args, _option),
                    (CleanOptions _option) => HandleCleanOptions(_option),
                    //(TestOptions _option) => HandleTestOptions(_option),
                    errs => DisplayHelp(parserResult)
                );
        }

        static int DisplayHelp<T>(ParserResult<T> result)
        {
            var helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                h.Heading = "QuickAccessShell 1.0";
                h.Copyright = "Copyright (c) 2023 Steins";
                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);
            Console.WriteLine(helpText);

            return -1;
        }

        private static string BuildOutputData(ArrayList input, string type, ArrayList debug)
        {
            var output = new OutputData
            {
                Data = input,
                Type = type,
                Debug = debug
            };

            return JsonConvert.SerializeObject(output);
        }

        private static int HandleListOptions(ListOptions options)
        {
            Dictionary<string, string> res;

            if (options.IsListRecentFiles)
            {
                res = _handler.GetRecentFiles();
            }
            else if (options.IsListFrequentFolders)
            {
                res = _handler.GetFrequentFolders();
            }
            else
            {
                res = _handler.GetQuickAccessDict();
            }

            ArrayList _data = new ArrayList { };
            foreach (var item in res)
            {
                _data.Add(item.Key);
            }

            if (options.IsListSupportedLanguage)
            {
                _data = new ArrayList(_handler.GetSupportLanguages());
            }

            ArrayList _debug = new ArrayList { };
            Console.WriteLine(BuildOutputData(_data, "list", _debug));

            return 0;
        }

        private static int HandleAddOptions(IEnumerable<string> args, AddOptions options)
        {
            ArrayList _data = new ArrayList { };
            ArrayList _debug = new ArrayList { };

            foreach (var item in options.AddItems)
            {
                _debug.Add(item);
                _handler.AddToQuickAccess(item);
            }

            Thread.Sleep(500);

            foreach (var item in options.AddItems)
            {
                bool res = _handler.IsInQuickAccess(item);

                _data.Add(res);
            }

            Console.WriteLine(BuildOutputData(_data, "add", _debug));

            return 0;
        }

        private static int HandleRemoveOptions(IEnumerable<string> args, RemoveOptions options)
        {
            ArrayList _data = new ArrayList { };
            ArrayList _debug = new ArrayList { };

            foreach (var item in options.RemoveItems)
            {
                _handler.RemoveFromQuickAccess(item);
            }

            Thread.Sleep(500);

            foreach (var item in options.RemoveItems)
            {
                bool res = _handler.IsInQuickAccess(item);

                _data.Add(res);
            }

            Console.WriteLine(BuildOutputData(_data, "remove", _debug));

            return 0;
        }

        private static int HandleShowOptions(ShowOptions options)
        {
            UInt32 accessType = 0;
            bool isShow = options.IsShow;

            if (options.IsShowAll)
            {
                accessType = 0;
            }
            else if (options.IsShowFrequentFolders)
            {
                accessType = 1;
            }
            else if (options.IsShowRecentFiles)
            {
                accessType = 2;
            }
            else if (options.IsShowSideMenuQuickAccess)
            {
                accessType = 3;
            }

            if (options.IsShowAll || options.IsShowSideMenuQuickAccess)
            {
                if (!_handler.IsAdminPrivilege())
                {
                    ArrayList res = new ArrayList { };

                    res.Add("Admin Privilege required.");

                    ArrayList _debug = new ArrayList { };
                    Console.WriteLine(BuildOutputData(res, "error", _debug));

                    return 0;
                }
            }

            _handler.UpdateShowQuickAccess(accessType, isShow);

            return 0;
        }

        private static int HandleCheckOptions(IEnumerable<string> args, CheckOptions options)
        {
            ArrayList res = new ArrayList { };

            if (!options.IsInQuickAccess && (options.IsShowAll || options.IsShowRecentFiles || options.IsShowFrequentFolders || options.IsShowSideMenuQuickAccess))
            {
                UInt32 accessType = 0;

                if (options.IsShowAll)
                {
                    accessType = 0;
                }
                else if (options.IsShowFrequentFolders)
                {
                    accessType = 1;
                }
                else if (options.IsShowRecentFiles)
                {
                    accessType = 2;
                }
                else if (options.IsShowSideMenuQuickAccess)
                {
                    accessType = 3;
                }

                res.Add(_handler.IsShowQuickAccess(accessType));
            }
            else if (options.IsInQuickAccess)
            {
                foreach (var item in options.Target)
                {
                    bool isInQuickAccess = _handler.IsInQuickAccess(item);
                    res.Add(isInQuickAccess);
                }
            }
            else if (options.IsCheckSupportedLanguage)
            {
                foreach (var item in options.Target)
                {
                    bool isSupportedLang = _handler.IsSupportedQuickAccessLanguage(item);
                    res.Add(isSupportedLang);
                }
            }

            ArrayList _debug = new ArrayList { };
            Console.WriteLine(BuildOutputData(res, "check", _debug));

            return 0;
        }

        private static int HandleCleanOptions(CleanOptions options)
        {
            if (options.IsEmptyAll)
            {
                _handler.EmptyQuickAccess();
            }
            else if (options.IsEmptyRecentFiles)
            {
                _handler.EmptyRecentFiles();
            }
            else if (options.IsEmptyFrequentFolders)
            {
                _handler.EmptyFrequentFolders();
            }

            return 0;
        }

        //private static int HandleTestOptions(TestOptions options)
        //{
        //    var path = @"C:\Intel";

        //    _handler.AddToQuickAccess(path);

        //    _handler.RemoveFromQuickAccess(path);

        //    return 0;
        //}
    }
}
