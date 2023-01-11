namespace QuickAccessTests
{
    [TestClass]
    public class QuickAccessTests
    {
        [TestMethod]
        public void CheckIsSupportedLanguage_InQuickAccessMenuName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var defaultSupportedLanguage = "en-US";
            bool isSupported = handler.IsSupportedQuickAccessLanguage(defaultSupportedLanguage);
            Assert.IsTrue(isSupported, "Should support en-US by default");

            var checkSupportedLangeage = "ru-Ru";
            bool isNewLanguageSupported = handler.IsSupportedQuickAccessLanguage(checkSupportedLangeage);
            Assert.IsFalse(isNewLanguageSupported, "Should not support ru-Ru by default");
        }

        [TestMethod]
        public void CheckIsSupportedLanguage_InFileExplorerMenuName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var defaultSupportedLanguage = "en-US";
            bool isSupported = handler.IsSupportedFileExplorerLanguage(defaultSupportedLanguage);
            Assert.IsTrue(isSupported, "Should support en-US by default");

            var checkSupportedLangeage = "ru-Ru";
            bool isNewLanguageSupported = handler.IsSupportedFileExplorerLanguage(checkSupportedLangeage);
            Assert.IsFalse(isNewLanguageSupported, "Should not support ru-Ru by default");
        }

        [TestMethod]
        public void CheckSupportLanguage_ByDefault()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            List<string> defaultSupportLanguage = new List<string> { "zh-CN", "zh-TW", "en-US", "fr-FR", "ru-RU"};
            var handlerSupportLanguage = handler.GetSupportLanguages();

            var isSame = defaultSupportLanguage.All(handlerSupportLanguage.Contains) && (defaultSupportLanguage.Count == handlerSupportLanguage.Count);

            Assert.IsTrue(isSame, "Missing default support language");
        }

        [TestMethod]
        public void AddQuickAccessMenuName_WithGivenName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            handler.AddQuickAccessMenuName("ja-JP", "¥Õ¥©¥ë¥À©`¤ò¥¯¥¤¥Ã¥¯ ¥¢¥¯¥»¥¹");

            bool addRes = handler.IsInQuickAccessMenuName("¥Õ¥©¥ë¥À©`¤ò¥¯¥¤¥Ã¥¯ ¥¢¥¯¥»¥¹");
            Assert.IsTrue(addRes, "Failed add menuName to handler");
        }

        [TestMethod]
        public void AddFileExplorerMenuName_WithGivenName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            handler.AddFileExplorerMenuName("ja-JP", "¥¨¥¯¥¹¥×¥í©`¥é©`");

            bool addRes = handler.IsInFileExplorerMenuName("¥¨¥¯¥¹¥×¥í©`¥é©`");
            Assert.IsTrue(addRes, "Failed add menuName to handler");
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
        public void IsInQuickAccess_WithNoExistsGiven()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            bool pathRes = handler.IsInQuickAccess("Z:\\679f85cb-0220-4080-b29b-5540cc05aab6");
            Assert.IsFalse(pathRes, "Should not be in quick access");

            bool keywordRes = handler.IsInQuickAccess("pneumonoultramicroscopicsilicovolcanoconiosis");
            Assert.IsFalse(keywordRes, "Should not be in quick access");
        }

        [TestMethod]
        public void IsInQuickAccess_WithExistsGiven()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            bool pathRes = handler.IsInQuickAccess("Z:\\679f85cb-0220-4080-b29b-5540cc05aab6");
            Assert.IsFalse(pathRes, "Should not be in quick access");

            bool keywordRes = handler.IsInQuickAccess("pneumonoultramicroscopicsilicovolcanoconiosis");
            Assert.IsFalse(keywordRes, "Should not be in quick access");
        }

        //[TestMethod]
        //public void AddToQuickAccess_WithGivenPath()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    string testPath = @"D:\test_add.txt";
        //    try
        //    {
        //        if (!File.Exists(testPath))
        //        {
        //            File.Create(testPath);
        //        }
        //    }
        //    catch (IOException e)
        //    {
        //        Console.WriteLine("Failed to delete test file. Errmsg: " + e);
        //    }

        //    bool isTestPathExists = File.Exists(testPath);

        //    Assert.IsTrue(isTestPathExists, "Failed to create test file");

        //    // Since the function will directly return IsInQuickAccess(), no need to check it twice
        //    bool addRes = handler.AddToQuickAccess(testPath);
        //    Assert.IsTrue(addRes, "Failed to add test file to quick access");

        //    bool addAnyRes = handler.AddToQuickAccess(@"Steins");
        //    Assert.IsFalse(addAnyRes, "Should be false by adding a not valid path to quick access");

        //    //bool removeRes = handler.RemoveFromQuickAccess(testPath);
        //    //Assert.IsTrue(removeRes, "Failed to remove test file from quick access");

        //    try
        //    {
        //        File.Delete(testPath);
        //    }
        //    catch (IOException e)
        //    {
        //        Console.WriteLine("Failed to delete test file. Errmsg: " + e);
        //    }
        //}

        //[TestMethod]
        //public void RemoveFromQuickAccess_WithAddFirst()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    string testPath = @"D:\test_remove.txt";
        //    try
        //    {
        //        if (!File.Exists(testPath))
        //        {
        //            File.Create(testPath);
        //        }
        //    }
        //    catch (IOException e)
        //    {
        //        Console.WriteLine("Failed to delete test file. Errmsg: " + e);
        //    }

        //    bool isTestPathExists = File.Exists(testPath);

        //    Assert.IsTrue(isTestPathExists, "Failed to create test file");

        //    // Since the function will directly return IsInQuickAccess(), no need to check it twice
        //    bool addRes = handler.AddToQuickAccess(testPath);
        //    Assert.IsTrue(addRes, "Failed to add test file to quick access");

        //    bool removeRes = handler.RemoveFromQuickAccess(testPath);
        //    Assert.IsTrue(removeRes, "Failed to remove test file from quick access");

        //    try
        //    {
        //        File.Delete(testPath);
        //    }
        //    catch (IOException e)
        //    {
        //        Console.WriteLine("Failed to delete test file. Errmsg: " + e);
        //    }
        //}

        //[TestMethod]
        //public void EmptyRecentFiles_IsDangerAction()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    handler.EmptyRecentFiles();

        //    var currentRecentFiles = handler.GetRecentFiles();
        //    var numberOfCurrentQuickAccess = currentRecentFiles.Count;

        //    Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty recent files");
        //}

        //[TestMethod]
        //public void EmptyFrequentFolders_IsDangerAction()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    handler.EmptyFrequentFolders();

        //    var currentFrequentFolders = handler.GetFrequentFolders();
        //    var numberOfCurrentQuickAccess = currentFrequentFolders.Count;

        //    Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty frequent folders");
        //}

        //[TestMethod]
        //public void ClearRecent_IsDangerAction()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    handler.EmptyQuickAccess();

        //    var currentQuickAccess = handler.GetQuickAccessDict();
        //    var numberOfCurrentQuickAccess = currentQuickAccess.Count;

        //    Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to empty quick access");
        //}

        [TestMethod]
        public void IsAdminPrivilege_DefaultFalse()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            var isAdmin = handler.IsAdminPrivilege();

            Assert.IsFalse(isAdmin, "Current user is admin?");
        }

        [TestMethod]
        public void GetQuickAccessRegistryKey_WithGivenKeyname()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            RegistryKey hklm = Registry.CurrentUser;
            RegistryKey hkExplorer = hklm.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer");

            var frequentKey = handler.GetQuickAccessRegistryKey("ShowFrequent");
            var currentFrequentKey = (int)hkExplorer.GetValue("ShowFrequent");

            Assert.AreEqual(frequentKey, currentFrequentKey, "Registry Key 'ShowFrequent' doesn't match");

            var recentKey = handler.GetQuickAccessRegistryKey("ShowRecent");
            var currentRecentKey = (int)hkExplorer.GetValue("ShowRecent");

            Assert.AreEqual(recentKey, currentRecentKey, "Registry Key 'ShowRecent' doesn't match");

            hklm = Registry.LocalMachine;
            hkExplorer = hklm.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer");

            var hubModeKey = handler.GetQuickAccessRegistryKey("HubMode");
            int currentHubModeKey;

            if (!hkExplorer.GetValueNames().Contains("HubMode"))
            {
                currentHubModeKey = -1;
            }
            else
            {
                currentHubModeKey = (int)hkExplorer.GetValue("HubMode");
            }

            Assert.AreEqual(hubModeKey, currentHubModeKey, "Registry Key 'HubMode' doesn't match");
        }

        [TestMethod]
        public void IsShowQuickAccess_WithGivenAccessType()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            int ShowFrequentValue = handler.GetQuickAccessRegistryKey("ShowFrequent");
            int ShowRecentValue = handler.GetQuickAccessRegistryKey("ShowRecent");
            int HubModeValue = handler.GetQuickAccessRegistryKey("HubMode");

            var isShowAll = handler.IsShowQuickAccess(0);
            var currentIsShowAll = ((ShowFrequentValue == 0) || (ShowRecentValue == 0) || (HubModeValue == 1)) ? false : true;

            Assert.AreEqual(isShowAll, currentIsShowAll, "Not showing all quick access properly");

            var isShowFrequent = handler.IsShowQuickAccess(1);
            var currentIsShowFrequent = ShowFrequentValue > 0 ? true : false;
            if (ShowFrequentValue == -1) currentIsShowFrequent = true;

            Assert.AreEqual(isShowFrequent, currentIsShowFrequent, "Not showing frequent folders properly");

            var isShowRecent = handler.IsShowQuickAccess(2);
            var currentIsShowRecent = ShowRecentValue > 0 ? true : false;
            if (ShowRecentValue == -1) currentIsShowRecent = true;

            Assert.AreEqual(isShowRecent, currentIsShowRecent, "Not showing recent files properly");

            var isShowSideMenuQuickAccess = handler.IsShowQuickAccess(3);
            var currentIsShowSideMenuQuickAccess = HubModeValue > 0 ? false : true;
            if (HubModeValue == -1) currentIsShowSideMenuQuickAccess = true;

            Assert.AreEqual(isShowSideMenuQuickAccess, currentIsShowSideMenuQuickAccess, "Not showing side menu quick access properly");
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