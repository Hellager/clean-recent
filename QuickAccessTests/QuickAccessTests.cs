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
        //public void ClearRecent_IsDangerAction()
        //{
        //    QuickAccessHandler handler = new QuickAccessHandler();

        //    handler.ClearRecent();

        //    var currentQuickAccess = handler.GetQuickAccessDict();
        //    var numberOfCurrentQuickAccess = currentQuickAccess.Count;

        //    Assert.AreEqual(numberOfCurrentQuickAccess, 0, "Failed to clear quick access list");
        //}

        [TestMethod]
        public void IsShowQuickAccess_WithGivenKey()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            RegistryKey hklm = Registry.CurrentUser;
            RegistryKey hkExplorer = hklm.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer");

            if (!hkExplorer.GetValueNames().Contains("ShowFrequent"))
            {
                hkExplorer = hklm.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer");
                hkExplorer.SetValue("ShowFrequent", 1, RegistryValueKind.DWord);
            }
            var currentFrequentKey = (int)hkExplorer.GetValue("ShowFrequent");
            handler.UpdateShowQuickAccess(1, currentFrequentKey > 0 ? false : true);
            bool isUpdateShowFrequent = handler.IsShowQuickAccess(1);

            Assert.AreEqual(isUpdateShowFrequent, currentFrequentKey > 0 ? false : true, "Failed to update frequent folder show status");

            handler.UpdateShowQuickAccess(1, !isUpdateShowFrequent);

            if (!hkExplorer.GetValueNames().Contains("ShowFrequent"))
            {
                hkExplorer = hklm.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer");
                hkExplorer.SetValue("ShowRecent", 1, RegistryValueKind.DWord);
            }
            var currentRecentKey = (int)hkExplorer.GetValue("ShowRecent");
            handler.UpdateShowQuickAccess(2, currentFrequentKey > 0 ? false : true);
            bool isUpdateShowRecent = handler.IsShowQuickAccess(2);

            Assert.AreEqual(isUpdateShowRecent, currentRecentKey > 0 ? false : true, "Failed to update recent file show status");

            handler.UpdateShowQuickAccess(2, !isUpdateShowRecent);
        }
    }
}