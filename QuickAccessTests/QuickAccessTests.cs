namespace QuickAccessTests
{
    /// <summary>
    /// Do not directly run or debug or tests!!! <br />
    /// Some of them are risky actions!!! <br />
    /// </summary>
    [TestClass]
    public class QuickAccessTests
    {
        [TestMethod]
        public void GetIsCurrentSystemDefaultSupported()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var isSupportedByDefault = handler.isSupportedSystem();

            Console.WriteLine("Current system whether supported by default: " + isSupportedByDefault);
        }

        [TestMethod]
        public void AddQuickAccessCommandName_WithGivenName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();
            bool isDefaultSupported = handler.isSupportedSystem();

            string commandName = "";

            if (!isDefaultSupported)
            {
                handler.AddquickAccessCommandName(commandName);

                bool isCurrentSupported = handler.isSupportedSystem();

                Assert.AreNotEqual(isDefaultSupported, isCurrentSupported, "Invalid command name for current system");
            }
        }

        [TestMethod]
        public void AddFileExplorerMenuName_WithGivenName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            string fileExplorerName = "";

            handler.AddFileExplorerMenuName(fileExplorerName);
        }

        [TestMethod]
        public void CheckSupportLanguage_ByDefault()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            List<string> defaultSupportLanguage = new List<string> { "zh-CN", "zh-TW", "en-US", "fr-FR", "ru-RU" };
            var handlerSupportLanguage = handler.GetDefaultSupportLanguages();

            var isSame = defaultSupportLanguage.All(handlerSupportLanguage.Contains) && (defaultSupportLanguage.Count == handlerSupportLanguage.Count);

            Assert.IsTrue(isSame, "Missing default support language");
        }

        [TestMethod]
        public void GetQuickAccessDict_CheckReturnType()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var quickAccess = handler.GetQuickAccessDict();

            bool isEqualObj = quickAccess is Dictionary<string, string>;

            Assert.IsTrue(isEqualObj, "Incorrect return type");
        }

        [TestMethod]
        public void GetFrequentFolders_CheckReturnType()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var frequentFolders = handler.GetFrequentFolders();

            bool isEqualObj = frequentFolders is Dictionary<string, string>;
            Assert.IsTrue(isEqualObj, "Incorrect return type");
        }

        [TestMethod]
        public void GetRecentFiles_CheckReturnType()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var recentFiles = handler.GetRecentFiles();

            bool isEqualObj = recentFiles is Dictionary<string, string>;
            Assert.IsTrue(isEqualObj, "Incorrect return type");
        }

        [TestMethod]
        public void CheckIsInQuickAccess()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var recentFiles = handler.GetRecentFiles().Keys;

            bool pathRes = handler.IsInQuickAccess(recentFiles.ElementAt(0));
            Assert.IsFalse(pathRes, "Should not be in quick access");

            bool keywordRes = handler.IsInQuickAccess("pneumonoultramicroscopicsilicovolcanoconiosis");
            Assert.IsFalse(keywordRes, "Should not be in quick access");
        }

        [TestMethod]
        public void AddToQuickAccess_WithGivenPath()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            string testPath = @"D:\test_add.txt";
            try
            {
                if (!File.Exists(testPath))
                {
                    File.Create(testPath);
                }
            }
            catch (IOException e)
            {
                Console.WriteLine("Failed to delete test file. Errmsg: " + e);
            }

            bool isTestPathExists = File.Exists(testPath);

            Assert.IsTrue(isTestPathExists, "Failed to create test file");

            // Since the function will directly return IsInQuickAccess(), no need to check it twice
            bool addRes = handler.AddToQuickAccess(testPath);
            Assert.IsTrue(addRes, "Failed to add test file to quick access");

            bool addAnyRes = handler.AddToQuickAccess(@"Steins");
            Assert.IsFalse(addAnyRes, "Should be false by adding a not valid path to quick access");

            try
            {
                File.Delete(testPath);
            }
            catch (IOException e)
            {
                Console.WriteLine("Failed to delete test file. Errmsg: " + e);
            }
        }

        [TestMethod]
        public void RemoveFromQuickAccess_WithAddFirst()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            string testPath = @"D:\test_remove.txt";
            try
            {
                if (!File.Exists(testPath))
                {
                    File.Create(testPath);
                }
            }
            catch (IOException e)
            {
                Console.WriteLine("Failed to delete test file. Errmsg: " + e);
            }

            bool isTestPathExists = File.Exists(testPath);

            Assert.IsTrue(isTestPathExists, "Failed to create test file");

            // Since the function will directly return IsInQuickAccess(), no need to check it twice
            bool addRes = handler.AddToQuickAccess(testPath);
            Assert.IsTrue(addRes, "Failed to add test file to quick access");

            handler.RemoveFromQuickAccess(new List<string> { testPath });
            bool removeRes = handler.IsInQuickAccess(testPath);
            Assert.IsTrue(removeRes, "Failed to remove test file from quick access");

            try
            {
                File.Delete(testPath);
            }
            catch (IOException e)
            {
                Console.WriteLine("Failed to delete test file. Errmsg: " + e);
            }
        }

        [TestMethod]
        public void EmptyRecentFiles()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            handler.EmptyRecentFiles();

            var currentRecentFiles = handler.GetRecentFiles();
            var numberOfCurrentQuickAccess = currentRecentFiles.Count;

            Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty recent files");
        }

        [TestMethod]
        public void EmptyFrequentFolders()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            handler.EmptyFrequentFolders();

            var currentFrequentFolders = handler.GetFrequentFolders();
            var numberOfCurrentQuickAccess = currentFrequentFolders.Count;

            Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty frequent folders");
        }

        [TestMethod]
        public void EmptyQuickAccess()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            handler.EmptyQuickAccess();

            var currentQuickAccess = handler.GetQuickAccessDict();
            var numberOfCurrentQuickAccess = currentQuickAccess.Count;

            Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty quick access");
        }

        [TestMethod]
        public void IsAdminPrivilege_DefaultFalse()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var isAdmin = handler.IsAdminPrivilege();

            Assert.IsFalse(isAdmin, "Current user is admin?");
        }

        [TestMethod]
        public void UpdateShowQuickAccess_WithGivenAccessType()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            if (handler.IsAdminPrivilege())
            {
                var isShow = handler.IsShowQuickAccess(0);
                handler.UpdateShowQuickAccess(0, !isShow);

                bool currentIsShow = handler.IsShowQuickAccess(0);

                Assert.AreEqual(isShow, !currentIsShow, "Failed to show/hide all quick access");

                handler.UpdateShowQuickAccess(0, isShow);
            }
            else
            {
                var isShowFrequent = handler.IsShowQuickAccess(0);
                handler.UpdateShowQuickAccess(0, !isShowFrequent);

                bool currentIsShowFrequent = handler.IsShowQuickAccess(0);

                Assert.AreEqual(isShowFrequent, !currentIsShowFrequent, "Failed to show/hide frequent folders");

                handler.UpdateShowQuickAccess(0, isShowFrequent);

                var isShowRecent = handler.IsShowQuickAccess(0);
                handler.UpdateShowQuickAccess(0, !isShowRecent);

                bool currentIsShowRecent = handler.IsShowQuickAccess(0);

                Assert.AreEqual(isShowRecent, !currentIsShowRecent, "Failed to show/hide recent files");

                handler.UpdateShowQuickAccess(0, isShowRecent);
            }
        }
    }
}