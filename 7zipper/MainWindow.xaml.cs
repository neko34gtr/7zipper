using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Zipper7
{
    public partial class MainWindow : Window
    {
        private const string ConfigFile = "7zipper.json";
        private string _sevenZipPath = @"C:\Program Files\7-Zip\7z.exe";

        private double _originalWidth = 150;
        private double _originalHeight = 150;

        public MainWindow()
        {
            InitializeComponent();
            LoadConfig();
        }

        private void LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFile))
                {
                    string json = File.ReadAllText(ConfigFile, new UTF8Encoding(false));
                    var config = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

                    if (config != null)
                    {
                        if (config.TryGetValue("Ultra", out var u)) ChkUltra.IsChecked = bool.Parse(u);
                        if (config.TryGetValue("Thread", out var t)) ChkThread.IsChecked = bool.Parse(t);
                        if (config.TryGetValue("Timestamp", out var ts)) ChkTimestamp.IsChecked = bool.Parse(ts);

                        // 最前面設定の読み込み（デフォルトON）
                        bool isTopmost = true;
                        if (config.TryGetValue("Topmost", out var tm)) isTopmost = bool.Parse(tm);
                        ChkTopmost.IsChecked = isTopmost;
                        this.Topmost = isTopmost;

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

                        if (config.TryGetValue("WindowTop", out var top) && double.TryParse(top, out double tVal)) this.Top = tVal;
                        if (config.TryGetValue("WindowLeft", out var left) && double.TryParse(left, out double lVal)) this.Left = lVal;

                        if (this.Top < 0) this.Top = 0;
                        if (this.Left < 0) this.Left = 0;
                    }
                }
                else
                {
                    // デフォルト設定
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
                Trace.WriteLine($"Config Error: {ex.Message}");
            }
        }

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
                File.WriteAllText(ConfigFile, json, new UTF8Encoding(false));
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

        private void ConfigButton_Click(object sender, RoutedEventArgs e)
        {
            _originalWidth = this.Width;
            _originalHeight = this.Height;

            this.Width = 250;
            this.Height = 300;

            DropZone.Visibility = Visibility.Collapsed;
            ConfigPanel.Visibility = Visibility.Visible;
        }

        private void ConfigClose_Click(object sender, RoutedEventArgs e)
        {
            this.Width = _originalWidth;
            this.Height = _originalHeight;

            ConfigPanel.Visibility = Visibility.Collapsed;
            DropZone.Visibility = Visibility.Visible;

            SaveConfig();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
                SaveConfig();
            }
        }

        private void ExitMenu_Click(object sender, RoutedEventArgs e)
        {
            SaveConfig();
            Application.Current.Shutdown();
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                SelectionGrid.Visibility = Visibility.Visible;
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

                    if (paths.Length > 1 && dropPoint.X < DropZone.ActualWidth / 2)
                    {
                        Task.Run(() => ExecuteParallel7Zip(paths));
                    }
                    else
                    {
                        Task.Run(() => Execute7Zip(paths));
                    }
                }
            }
        }

        private void ExecuteParallel7Zip(string[] paths)
        {
            var options = new ParallelOptions { MaxDegreeOfParallelism = 3 };

            Parallel.ForEach(paths, options, path =>
            {
                Execute7Zip(new string[] { path });
            });
        }

        private void Execute7Zip(string[] paths)
        {
            if (!File.Exists(_sevenZipPath))
            {
                Dispatcher.Invoke(() => MessageBox.Show("7z.exe が見つかりません。"));
                return;
            }

            try
            {
                string ext = "";
                bool isUltra = true;
                bool isThread = true;
                bool isTimestamp = true;

                Dispatcher.Invoke(() =>
                {
                    ext = ((ComboBoxItem)ComboExt.SelectedItem).Content?.ToString() ?? ".zip";
                    isUltra = ChkUltra.IsChecked == true;
                    isThread = ChkThread.IsChecked == true;
                    isTimestamp = ChkTimestamp.IsChecked == true;
                });

                string targetDir = Path.GetDirectoryName(paths[0]) ?? "";
                string baseName = Path.GetFileName(paths[0]);
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
                Trace.WriteLine($"圧縮エラー: {ex.Message}");
            }
        }
    }
}