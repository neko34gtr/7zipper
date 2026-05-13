using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace Zipper7
{
    /// <summary>
    /// 7zipper: 7-Zipをベースとした直感的な圧縮・解凍ラッパーアプリ
    /// ドラッグ&ドロップ位置による挙動の変化、ファジーな命名、スマートな解凍判定を特徴とする
    /// </summary>
    public partial class MainWindow : Window
    {
        // 設定ファイル名（アプリ名に基づいた命名）
        private const string ConfigFile = "7zipper.json";
        // 7-Zip実行ファイルの標準的なインストールパス
        private string _sevenZipPath = @"C:\Program Files\7-Zip\7z.exe";

        // 設定パネル展開時のサイズ復元用バッファ
        private double _originalWidth = 150;
        private double _originalHeight = 150;

        // タスクバーアイコン点滅用のWin32 API
        [DllImport("user32.dll")]
        private static extern bool FlashWindow(IntPtr hwnd, bool bInvert);

        public MainWindow()
        {
            InitializeComponent();
            LoadConfig();

            // OSのシャットダウンやログオフ時のイベントを登録（設定の確実な保存）
            if (Application.Current != null)
            {
                Application.Current.SessionEnding += (s, e) => SaveConfig();
            }
        }

        // ウィンドウが閉じられる（×ボタンやメニュー終了）直前のイベント
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            SaveConfig();
        }

        /// <summary>
        /// json形式の設定ファイルを読み込む
        /// </summary>
        private void LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFile))
                {
                    // UTF-8 (BOM無し) で読み込み
                    string json = File.ReadAllText(ConfigFile, new UTF8Encoding(false));
                    var config = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

                    if (config != null)
                    {
                        if (config.TryGetValue("Ultra", out var u)) ChkUltra.IsChecked = bool.Parse(u);
                        if (config.TryGetValue("Thread", out var t)) ChkThread.IsChecked = bool.Parse(t);
                        if (config.TryGetValue("Timestamp", out var ts)) ChkTimestamp.IsChecked = bool.Parse(ts);

                        // 最前面設定の読み込み（デフォルト: True）
                        bool isTopmost = true;
                        if (config.TryGetValue("Topmost", out var tm)) isTopmost = bool.Parse(tm);
                        ChkTopmost.IsChecked = isTopmost;
                        this.Topmost = isTopmost;

                        // 拡張子選択の復元
                        if (config.TryGetValue("Ext", out var e))
                        {
                            foreach (ComboBoxItem item in ComboExt.Items)
                            {
                                if (item.Content.ToString() == e)
                                {
                                    ComboExt.SelectedItem = item;
                                    break;
                                }
                            }
                        }

                        // ウィンドウ座標の復元
                        if (config.TryGetValue("WindowTop", out var top) && double.TryParse(top, out double tVal)) this.Top = tVal;
                        if (config.TryGetValue("WindowLeft", out var left) && double.TryParse(left, out double lVal)) this.Left = lVal;

                        // 画面外へ消えないためのセーフティ
                        if (this.Top < 0) this.Top = 0;
                        if (this.Left < 0) this.Left = 0;
                    }
                }
                else
                {
                    // 初回起動時のデフォルト設定
                    ChkUltra.IsChecked = true;
                    ChkThread.IsChecked = true;
                    ChkTimestamp.IsChecked = true;
                    ChkTopmost.IsChecked = true;
                    this.Topmost = true;
                    ComboExt.SelectedIndex = 0;
                    SaveConfig();
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Config Load Error: {ex.Message}");
            }
        }

        /// <summary>
        /// 設定をjson形式で保存（アトミックな書き込みによるファイル破損防止）
        /// </summary>
        private void SaveConfig()
        {
            try
            {
                var config = new Dictionary<string, string>
                {
                    { "Ultra", ChkUltra.IsChecked?.ToString() ?? "True" },
                    { "Thread", ChkThread.IsChecked?.ToString() ?? "True" },
                    { "Timestamp", ChkTimestamp.IsChecked?.ToString() ?? "True" },
                    { "Topmost", ChkTopmost.IsChecked?.ToString() ?? "True" },
                    { "Ext", ((ComboBoxItem)ComboExt.SelectedItem).Content?.ToString() ?? ".zip" },
                    { "WindowTop", this.Top.ToString() },
                    { "WindowLeft", this.Left.ToString() }
                };

                string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });

                // 一時ファイルに書き込んでから置換するアトミック書き込み
                string tempFile = ConfigFile + ".tmp";
                File.WriteAllText(tempFile, json, new UTF8Encoding(false));

                if (File.Exists(ConfigFile))
                {
                    File.Delete(ConfigFile);
                }
                File.Move(tempFile, ConfigFile);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Save Config Error: {ex.Message}");
            }
        }

        private void ChkTopmost_Changed(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded)
            {
                this.Topmost = ChkTopmost.IsChecked == true;
            }
        }

        // ⚙ボタンクリックで設定パネルを表示
        private void ConfigButton_Click(object sender, RoutedEventArgs e)
        {
            _originalWidth = this.Width;
            _originalHeight = this.Height;

            // パネル表示用にウィンドウサイズを拡張
            this.Width = 250;
            this.Height = 300;
            DropZone.Visibility = Visibility.Collapsed;
            ConfigPanel.Visibility = Visibility.Visible;
        }

        // 設定を閉じてメインUIに戻る
        private void ConfigClose_Click(object sender, RoutedEventArgs e)
        {
            this.Width = _originalWidth;
            this.Height = _originalHeight;
            ConfigPanel.Visibility = Visibility.Collapsed;
            DropZone.Visibility = Visibility.Visible;
            SaveConfig();
        }

        // ウィンドウのドラッグ移動
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        // コンテキストメニューからの終了
        private void ExitMenu_Click(object sender, RoutedEventArgs e)
        {
            this.Close(); // OnClosingイベントを経由させる
        }

        // エリア進入時に選択オーバーレイを表示
        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                SelectionGrid.Visibility = Visibility.Visible;
                string[]? paths = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                UpdateSelectionUI(paths);
            }
        }

        /// <summary>
        /// ドラッグ内容に応じて「個別」エリアの表示を動的に警告へ切り替える
        /// </summary>
        private void UpdateSelectionUI(string[]? paths)
        {
            if (paths == null) return;

            // ファイルが含まれているかチェック
            bool hasFiles = paths.Any(p => !Directory.Exists(p));

            var indivBorder = (Border)SelectionGrid.Children[0];
            var indivText = (TextBlock)indivBorder.Child;

            if (hasFiles && paths.Length > 1)
            {
                // ファイル混在ドロップで個別を選ぼうとしている場合に警告
                indivText.Text = "ファイル混在\n(一括推奨)";
                indivText.Foreground = Brushes.Orange;
            }
            else
            {
                indivText.Text = "個別";
                indivText.Foreground = Brushes.White;
            }
        }

        private void OnDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                DropZone.Background = new SolidColorBrush(Color.FromArgb(0xCC, 0x44, 0x44, 0x44));
            }
            e.Handled = true;
        }

        private void OnDragLeave(object sender, DragEventArgs e)
        {
            DropZone.Background = new SolidColorBrush(Color.FromArgb(0xCC, 0x22, 0x22, 0x22));
            SelectionGrid.Visibility = Visibility.Collapsed;
        }

        // ドロップ実行
        private void OnDrop(object sender, DragEventArgs e)
        {
            DropZone.Background = new SolidColorBrush(Color.FromArgb(0xCC, 0x22, 0x22, 0x22));
            SelectionGrid.Visibility = Visibility.Collapsed;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[]? paths = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                if (paths != null && paths.Length > 0)
                {
                    Point dropPoint = e.GetPosition(DropZone);

                    // 事故防止ロジック: ファイルが含まれている場合は「個別」を拒否して強制的に「一括」扱いとする
                    bool hasFiles = paths.Any(p => !Directory.Exists(p));
                    bool isIndividual = paths.Length > 1 &&
                                       dropPoint.X < DropZone.ActualWidth / 2 &&
                                       !hasFiles;

                    // UIスレッドを止めないようバックグラウンド処理
                    Task.Run(() => {
                        ProcessItems(paths, isIndividual);
                        NotifyCompletion();
                    });
                }
            }
        }

        /// <summary>
        /// アイテムの性質（圧縮か解凍か）を判定して処理を振り分ける
        /// </summary>
        private void ProcessItems(string[] paths, bool isIndividual)
        {
            string[] archiveExtensions = { ".zip", ".7z", ".lzh", ".rar", ".tar", ".gz" };

            if (isIndividual)
            {
                // 最大3並列で個別処理を実行
                var options = new ParallelOptions { MaxDegreeOfParallelism = 3 };
                Parallel.ForEach(paths, options, path =>
                {
                    string ext = Path.GetExtension(path).ToLower();
                    if (archiveExtensions.Contains(ext))
                        ExecuteExtract(path);
                    else
                        Execute7Zip(new string[] { path });
                });
            }
            else
            {
                // 一括処理
                string firstExt = Path.GetExtension(paths[0]).ToLower();
                // 1つのアーカイブのみがドロップされた場合は解凍、それ以外は一括圧縮
                if (archiveExtensions.Contains(firstExt) && paths.Length == 1)
                {
                    ExecuteExtract(paths[0]);
                }
                else
                {
                    Execute7Zip(paths);
                }
            }
        }

        /// <summary>
        /// アーカイブの解凍（展開）を実行
        /// </summary>
        private void ExecuteExtract(string archivePath)
        {
            if (!File.Exists(_sevenZipPath)) return;

            try
            {
                string targetBaseDir = Path.GetDirectoryName(archivePath) ?? "";

                // アーカイブ構造を診察し、展開先フォルダを自動決定
                bool shouldExtractDirectly = CheckIfArchiveHasSingleRootFolder(archivePath);

                string outputDir = shouldExtractDirectly
                    ? targetBaseDir
                    : Path.Combine(targetBaseDir, Path.GetFileNameWithoutExtension(archivePath));

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = _sevenZipPath,
                    // x:パス維持解凍, -o:出力先, -y:強制上書き承認
                    Arguments = $"x \"{archivePath}\" -o\"{outputDir}\" -y",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process? p = Process.Start(psi))
                {
                    p?.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Extraction Error: {ex.Message}");
            }
        }

        /// <summary>
        /// アーカイブのルート階層をリスト化し、展開時にフォルダを作成すべきか判定する
        /// </summary>
        private bool CheckIfArchiveHasSingleRootFolder(string archivePath)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = _sevenZipPath,
                    Arguments = $"l \"{archivePath}\" -slt", // 詳細リストモード
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (Process? p = Process.Start(psi))
                {
                    string output = p?.StandardOutput.ReadToEnd() ?? "";
                    p?.WaitForExit();

                    var lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                    var rootItems = new HashSet<string>();
                    var folderItems = new HashSet<string>();

                    string currentPath = "";
                    bool currentIsFolder = false;

                    foreach (var line in lines)
                    {
                        if (line.StartsWith("Path = ")) currentPath = line.Substring(7);
                        if (line.StartsWith("Attributes = ")) currentIsFolder = line.Contains("D");

                        // 7-Zipリストの区切り線等は無視
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("----------")) continue;

                        if (!string.IsNullOrEmpty(currentPath))
                        {
                            // ルート直下のアイテムのみを抽出
                            string cleanPath = currentPath.TrimEnd('\\');
                            if (!cleanPath.Contains("\\"))
                            {
                                rootItems.Add(cleanPath);
                                if (currentIsFolder) folderItems.Add(cleanPath);
                            }
                        }
                    }
                    // ルートに唯一のフォルダが1つだけ存在する場合のみ、そのまま解凍(true)とする
                    return rootItems.Count == 1 && folderItems.Count == 1;
                }
            }
            catch { return false; }
        }

        /// <summary>
        /// 7-Zipを使用した圧縮処理を実行
        /// </summary>
        private void Execute7Zip(string[] paths)
        {
            if (!File.Exists(_sevenZipPath)) return;

            try
            {
                string ext = "";
                bool isUltra = true;
                bool isThread = true;
                bool isTimestamp = true;

                // UI要素からの設定取得はDispatcher経由で行う
                Dispatcher.Invoke(() =>
                {
                    ext = ((ComboBoxItem)ComboExt.SelectedItem).Content?.ToString() ?? ".zip";
                    isUltra = ChkUltra.IsChecked == true;
                    isThread = ChkThread.IsChecked == true;
                    isTimestamp = ChkTimestamp.IsChecked == true;
                });

                string targetDir = Path.GetDirectoryName(paths[0]) ?? "";

                // ファジー・ロジックによる最適なアーカイブ名の決定
                string baseName = DetermineArchiveName(paths);
                string outputArchive = Path.Combine(targetDir, $"{baseName}{ext}");

                List<string> args = new List<string> { "a", $"\"{outputArchive}\"" };
                args.AddRange(paths.Select(p => $"\"{p}\""));

                if (ext == ".zip") args.Add("-tzip");
                if (isUltra) args.Add("-mx=9");
                if (isThread) args.Add("-mmt=on");
                if (isTimestamp) args.Add("-stl");
                args.Add("-y");

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = _sevenZipPath,
                    Arguments = string.Join(" ", args),
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process? p = Process.Start(psi))
                {
                    p?.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Compression Error: {ex.Message}");
            }
        }

        /// <summary>
        /// 複数ファイル名の中から、アーカイブ名のベースを決定するロジック。
        /// 1. 実行ファイル(.exe)があれば最優先。
        /// 2. なければ区切り文字に基づいた出現頻度の高いワードを抽出。
        /// </summary>
        private string DetermineArchiveName(string[] paths)
        {
            // 単体ドロップならそのままの名前
            if (paths.Length == 1) return Path.GetFileNameWithoutExtension(paths[0]);

            // 【追加ロジック】実行ファイル(.exe)が含まれているかスキャン
            string? firstExeName = paths
                .Where(p => Path.GetExtension(p).Equals(".exe", StringComparison.OrdinalIgnoreCase))
                .Select(p => Path.GetFileNameWithoutExtension(p))
                .FirstOrDefault();

            // .exeが見つかれば、それを確定名称として即座に返す
            if (!string.IsNullOrEmpty(firstExeName)) return firstExeName;

            var wordFrequencies = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            // 一般的に使用される区切り文字
            char[] delimiters = { '.', '-', '_', ' ' };

            foreach (var path in paths)
            {
                string fileName = Path.GetFileNameWithoutExtension(path);
                string[] words = fileName.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

                foreach (var word in words)
                {
                    if (wordFrequencies.ContainsKey(word)) wordFrequencies[word]++;
                    else wordFrequencies[word] = 1;
                }
            }

            // 最も出現頻度が高く、かつ長い単語を優先的に選出
            var mostFrequentWord = wordFrequencies
                .OrderByDescending(x => x.Value)
                .ThenByDescending(x => x.Key.Length)
                .FirstOrDefault();

            // 共通項（頻度2以上）があれば採用。なければ最初のファイル名を採用
            if (mostFrequentWord.Value > 1) return mostFrequentWord.Key;

            return Path.GetFileNameWithoutExtension(paths[0]);
        }

        /// <summary>
        /// 処理完了をユーザーに通知（タスクバー点滅 + トースト）
        /// </summary>
        private void NotifyCompletion()
        {
            Dispatcher.Invoke(() =>
            {
                var helper = new WindowInteropHelper(this);
                FlashWindow(helper.Handle, true);
                ShowToast("7zipper", "処理が完了しました。");
            });
        }

        /// <summary>
        /// PowerShell経由でWindows標準のトースト通知を発行
        /// </summary>
        private void ShowToast(string title, string message)
        {
            try
            {
                string script = $"[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null;" +
                                $"$xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02);" +
                                $"$xml.GetElementsByTagName('text').Item(0).AppendChild($xml.CreateTextNode('{title}')) | Out-Null;" +
                                $"$xml.GetElementsByTagName('text').Item(1).AppendChild($xml.CreateTextNode('{message}')) | Out-Null;" +
                                $"$toast = [Windows.UI.Notifications.ToastNotification]::new($xml);" +
                                $"[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('7zipper').Show($toast);";

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{script}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Toast Notification Error: {ex.Message}");
            }
        }
    }
}