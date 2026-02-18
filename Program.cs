using System;
using System.Security.Principal;
using System.Globalization;
using System.Reflection;
using System.Threading;

namespace AutoWindowsUpdate
{
    /// <summary>
    /// Windows Updateを自動的に実行するためのメインクラスです。
    /// </summary>
    class Program
    {
        /// <summary>
        /// 言語リソース
        /// </summary>
        static class Loc
        {
            public static bool IsJapanese => CultureInfo.CurrentUICulture.Name.Equals("ja-JP", StringComparison.OrdinalIgnoreCase);

            // 無理やり感あるが一旦これで
            public static string RunAsAdmin => IsJapanese 
                ? "このアプリケーションは管理者として実行する必要があります。" 
                : "This application must be run as Administrator.";
            public static string PressKeyToExit => IsJapanese
                ? "キーを押して終了してください..."
                : "Press any key to exit...";
            public static string CheckingUpdates => IsJapanese
                ? "Windows Updateを確認しています..."
                : "Checking for Windows Updates...";
            public static string SessionCreateFail => IsJapanese
                ? "Microsoft.Update.Sessionの作成に失敗しました。Windows環境であることを確認してください。"
                : "Failed to create Microsoft.Update.Session. Ensure you are on a Windows machine.";
            public static string NoUpdatesFound => IsJapanese
                ? "更新プログラムは見つかりませんでした。"
                : "No updates found.";
            public static string ErrorOccurred(string msg) => IsJapanese
                ? $"エラーが発生しました: {msg}"
                : $"An error occurred: {msg}";
            public static string SearchingUpdates => IsJapanese
                ? "更新プログラムを検索中..."
                : "Searching for updates...";
            public static string UpdatesFound(int count) => IsJapanese
                ? $"{count} 個の更新プログラムが見つかりました。"
                : $"Found {count} updates.";
            public static string DownloadingUpdates => IsJapanese
                ? "\n更新プログラムをダウンロード中..."
                : "\nDownloading updates...";
            public static string DownloadComplete => IsJapanese
                ? "ダウンロードが完了しました。"
                : "Download complete.";
            public static string InstallingUpdates => IsJapanese
                ? "\n更新プログラムをインストール中..."
                : "\nInstalling updates...";
            public static string NoUpdatesToInstall => IsJapanese
                ? "インストール可能な更新プログラムがありません。"
                : "No updates ready to install.";
            public static string InstallResult(int code) => IsJapanese
                ? $"\nインストール結果: コード {code}"
                : $"\nInstallation Result: code {code}";

            public static string RebootRequired(bool required) => IsJapanese
                ? $"再起動が必要: {required}"
                : $"Reboot Required: {required}";
            public static string RebootNeededMessage => IsJapanese
                ? "インストールを完了するために再起動が必要です。"
                : "A reboot is required to complete the installation.";
        }

        static void Main(string[] args)
        {
            if (!IsAdministrator())
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(Loc.RunAsAdmin);
                Console.ResetColor();
                Console.WriteLine(Loc.PressKeyToExit);
                Console.ReadKey();
                return;
            }

            Console.WriteLine(Loc.CheckingUpdates);

            try
            {
                // Update Sessionの作成
                Type? updateSessionType = Type.GetTypeFromProgID("Microsoft.Update.Session");
                if (updateSessionType == null)
                {
                    Console.WriteLine(Loc.SessionCreateFail);
                    return;
                }
                
                dynamic updateSession = Activator.CreateInstance(updateSessionType!)!;

                // 1. 検索 / Search
                dynamic searchResult = SearchUpdates(updateSession);

                if (searchResult.Updates.Count == 0)
                {
                    Console.WriteLine(Loc.NoUpdatesFound);
                    return;
                }

                // 2. ダウンロード / Download
                DownloadUpdates(updateSession, searchResult);

                // 3. インストール / Install
                InstallUpdates(updateSession, searchResult);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                //Console.WriteLine(Loc.ErrorOccurred(ex.Message));
                Console.WriteLine(Loc.ErrorOccurred(ex.ToString()));
                Console.ResetColor();
            }

            Console.WriteLine($"\n{Loc.PressKeyToExit}");
            Console.ReadKey();
        }

        /// <summary>
        /// アプリケーションが管理者権限で実行されているか確認
        /// </summary>
        /// <returns>管理者として実行されている場合は <c>true</c></returns>
        static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        /// <summary>
        /// 未インストールの更新プログラムを検索
        /// </summary>
        /// <param name="updateSession">Windows Updateセッションオブジェクト</param>
        /// <returns>検索結果を含む <c>ISearchResult</c> オブジェクト</returns>
        static dynamic SearchUpdates(dynamic updateSession)
        {
            dynamic updateSearcher = updateSession.CreateUpdateSearcher();
            Console.WriteLine(Loc.SearchingUpdates);
            
            // "IsInstalled=0" 未インストールの更新プログラム
            dynamic searchResult = updateSearcher.Search("IsInstalled=0");
            Console.WriteLine(Loc.UpdatesFound(searchResult.Updates.Count));
            
            return searchResult;
        }

        /// <summary>
        /// 検索された更新プログラムをダウンロード
        /// </summary>
        /// <param name="updateSession">Windows Updateセッションオブジェクト</param>
        /// <param name="searchResult">検索結果オブジェクト</param>
        static void DownloadUpdates(dynamic updateSession, dynamic searchResult)
        {
            dynamic updatesToDownload = Activator.CreateInstance(Type.GetTypeFromProgID("Microsoft.Update.UpdateColl")!)!;

            for (int i = 0; i < searchResult.Updates.Count; i++)
            {
                dynamic update = searchResult.Updates[i];
                Console.WriteLine($"[{i + 1}] {update.Title}");
                updatesToDownload.Add(update);
            }

            Console.WriteLine(Loc.DownloadingUpdates);
            dynamic downloader = updateSession.CreateUpdateDownloader();
            downloader.Updates = updatesToDownload;
            
            // 同期ダウンロード / Synchronous Download
            // (dynamicでの非同期コールバックが困難なため同期処理を使用)
            downloader.Download();
            
            Console.WriteLine(Loc.DownloadComplete);
        }

        /// <summary>
        /// ダウンロードされた更新プログラムをインストール
        /// </summary>
        /// <param name="updateSession">Windows Updateセッションオブジェクト</param>
        /// <param name="searchResult">検索結果オブジェクト</param>
        static void InstallUpdates(dynamic updateSession, dynamic searchResult)
        {
            Console.WriteLine(Loc.InstallingUpdates);
            dynamic updatesToInstall = Activator.CreateInstance(Type.GetTypeFromProgID("Microsoft.Update.UpdateColl")!)!;

            // 未インストールのプログラムを探す
            for (int i = 0; i < searchResult.Updates.Count; i++)
            {
                dynamic update = searchResult.Updates[i];
                if (update.IsDownloaded)
                {
                    updatesToInstall.Add(update);
                }
            }

            if (updatesToInstall.Count == 0)
            {
                Console.WriteLine(Loc.NoUpdatesToInstall);
                return;
            }

            dynamic installer = updateSession.CreateUpdateInstaller();
            installer.Updates = updatesToInstall;

            // 同期インストール / Synchronous Install
            dynamic installationResult = installer.Install();

            Console.WriteLine(Loc.InstallResult(installationResult.ResultCode));
            Console.WriteLine(Loc.RebootRequired(installationResult.RebootRequired));

            if (installationResult.RebootRequired)
            {
                Console.WriteLine(Loc.RebootNeededMessage);
            }
        }
    }
}
