namespace QuickAccessTests
{
    [TestClass]
    public class QuickAccessTests
    {
        //[TestMethod]
        //public void Debit_WithValidAmount_UpdatesBalance()
        //{
        //    // Arrange
        //    double beginningBalance = 11.99;
        //    double debitAmount = 4.55;
        //    double expected = 7.44;
        //    BankAccount account = new BankAccount("Mr. Bryan Walton", beginningBalance);

        //    // Act
        //    account.Debit(debitAmount);

        //    // Assert
        //    double actual = account.Balance;
        //    Assert.AreEqual(expected, actual, 0.001, "Account not debited correctly");
        //}

        [TestMethod]
        public void AddQuickAccessMenuName_WithGivenName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            handler.AddQuickAccessMenuName("SomeMenuName");

            bool addRes = handler.IsInQuickAccessMenuName("SomeMenuName");
            Assert.IsTrue(addRes, "Failed add menuName to handler");
        }

        [TestMethod]
        public void AddFileExplorerMenuName_WithGivenName()
        {
            QuickAccessHandler handler = new QuickAccessHandler();

            handler.AddFileExplorerMenuName("SomeMenuName");

            bool addRes = handler.IsInFileExplorerMenuName("SomeMenuName");
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

            //bool removeRes = handler.RemoveFromQuickAccess(testPath);
            //Assert.IsTrue(removeRes, "Failed to remove test file from quick access");

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

            bool removeRes = handler.RemoveFromQuickAccess(testPath);
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
            bool isUpdateShowFrequent = handler.IsShowQuickAccess(QuickAccessHandler.QuickAccessType.FrequentFolder, currentFrequentKey > 0 ? false : true);

            Assert.AreEqual(isUpdateShowFrequent, currentFrequentKey > 0 ? false : true, "Failed to update frequent folder show status");

            handler.IsShowQuickAccess(QuickAccessHandler.QuickAccessType.FrequentFolder, !isUpdateShowFrequent);

            if (!hkExplorer.GetValueNames().Contains("ShowFrequent"))
            {
                hkExplorer = hklm.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer");
                hkExplorer.SetValue("ShowRecent", 1, RegistryValueKind.DWord);
            }
            var currentRecentKey = (int)hkExplorer.GetValue("ShowRecent");
            bool isUpdateShowRecent = handler.IsShowQuickAccess(QuickAccessHandler.QuickAccessType.RecentFile, currentRecentKey > 0 ? false : true);

            Assert.AreEqual(isUpdateShowRecent, currentRecentKey > 0 ? false : true, "Failed to update recent file show status");

            handler.IsShowQuickAccess(QuickAccessHandler.QuickAccessType.RecentFile, !isUpdateShowRecent);
        }
    }
}