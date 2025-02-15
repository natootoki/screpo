using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

class Program
{
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    [DllImport("shcore.dll")]
    private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

    private const int MDT_EFFECTIVE_DPI = 0;

    private const int HOTKEY_ID = 1;
    private const uint MOD_CONTROL = 0x0002;
    private const uint VK_SNAPSHOT = 0x2C;
    private const int WM_HOTKEY = 0x0312;

    [STAThread]
    static void Main(string[] args)
    {

        // アプリケーションをDPI対応にする
        SetProcessDPIAware();

        Application.Run(new ScreenshotToolForm());
    }

    private class ScreenshotToolForm : Form
    {
        private float dpiScale; // DPIスケール倍率
        private string saveFolder;
        private string screenshotsFolder;
        private TextBox titleInput;
        private Button generateHtmlButton;
        private Button takeScreenshotButton;
        private ComboBox screenSelector;
        private CheckBox autoGenerateHtmlCheckBox;
        private Label dateTimeLabel;
        private Timer dateTimeTimer;
        private TrackBar opacityTrackBar;

        public ScreenshotToolForm()
        {
            // DPI スケールを取得
            dpiScale = GetDpiScale();

            saveFolder = GetFolderWithFileDialog();
            if (string.IsNullOrEmpty(saveFolder))
            {
                MessageBox.Show("フォルダが選択されませんでした。終了します。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                Environment.Exit(0);
            }

            screenshotsFolder = Path.Combine(saveFolder, "img");
            if (!Directory.Exists(screenshotsFolder))
            {
                Directory.CreateDirectory(screenshotsFolder);
            }

            Text = "screpo - " + saveFolder;
            int baseWidth = 400;
            int padding = 20;
            int calculatedWidth = TextRenderer.MeasureText("screpo - " + saveFolder, this.Font).Width + padding;
            Size = new Size((int)(baseWidth * dpiScale), (int)(220 + 100 * dpiScale));
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            this.Load += (sender, e) =>
            {
                TopMost = true;
                this.Activate();
                TopMost = false;
            };

            dateTimeLabel = new Label
            {
                Location = new Point(10, 10),
                Size = new Size((int)(380 * dpiScale), 30),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Arial", 14, FontStyle.Bold)
            };
            Controls.Add(dateTimeLabel);

            dateTimeTimer = new Timer
            {
                Interval = 1000
            };
            dateTimeTimer.Tick += (s, e) => UpdateDateTime();
            dateTimeTimer.Start();

            var screenSelectorLabel = new Label
            {
                Text = "撮影するディスプレイ:",
                Location = new Point(10, 50),
                AutoSize = true,
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            Controls.Add(screenSelectorLabel);

            screenSelector = new ComboBox
            {
                Location = new Point((int)(150 * dpiScale), 45),
                Size = new Size(220, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            foreach (var screen in Screen.AllScreens)
            {
                screenSelector.Items.Add(screen.DeviceName);
            }

            if (screenSelector.Items.Count > 0)
            {
                screenSelector.SelectedIndex = 0;
            }

            Controls.Add(screenSelector);

            var titleLabel = new Label
            {
                Text = "HTMLタイトル:",
                Location = new Point(10, 90),
                AutoSize = true,
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            Controls.Add(titleLabel);

            titleInput = new TextBox
            {
                Location = new Point((int)(100 * dpiScale), 85),
                Size = new Size(250, 20),
                Font = new Font("Arial", 9, FontStyle.Regular),
                Text = "スクリーンショット一覧"
            };
            Controls.Add(titleInput);

            takeScreenshotButton = new Button
            {
                Text = "スクリーンショット撮影",
                Location = new Point(10, 120),
                Size = new Size((int)(150*dpiScale), (int)(25*dpiScale)),
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            takeScreenshotButton.Click += (sender, e) => TakeScreenshot();
            Controls.Add(takeScreenshotButton);

            generateHtmlButton = new Button
            {
                Text = "HTML生成",
                Location = new Point((int)(170*dpiScale), 120),
                Size = new Size((int)(100 * dpiScale), (int)(25*dpiScale)),
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            generateHtmlButton.Click += async (sender, e) =>
            {
                await GenerateHtmlFile(titleInput.Text);
                var selectedScreen = Screen.AllScreens[screenSelector.SelectedIndex];
                Rectangle bounds = selectedScreen.Bounds;
                await ShowNotificationOnScreen("HTMLファイルを生成しました。", Color.LightGreen, 3000, bounds);
            };
            Controls.Add(generateHtmlButton);

            Button copyFolderPathButton = new Button
            {
                Text = "パスをコピーする",
                Location = new Point((int)(280 * dpiScale), 120),
                Size = new Size((int)(100 * dpiScale), (int)(25*dpiScale)),
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            copyFolderPathButton.Click += (sender, e) => CopyFolderPathToClipboard();
            Controls.Add(copyFolderPathButton);

            Button saveClipboardImageButton = new Button
            {
                Text = "クリップボードの画像を保存",
                Location = new Point(10, (int)(200 + 50*dpiScale)),
                Size = new Size((int)(200*dpiScale), (int)(25*dpiScale)),
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            saveClipboardImageButton.Click += async (sender, e) => await SaveClipboardImageToFolder();
            Controls.Add(saveClipboardImageButton);

            autoGenerateHtmlCheckBox = new CheckBox
            {
                Text = "HTML自動生成",
                Location = new Point(10, 160),
                AutoSize = true,
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            Controls.Add(autoGenerateHtmlCheckBox);

            var alwaysOnTopCheckBox = new CheckBox
            {
                Text = "最前面固定",
                Location = new Point(10, 190),
                AutoSize = true,
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            alwaysOnTopCheckBox.CheckedChanged += (sender, e) =>
            {
                TopMost = alwaysOnTopCheckBox.Checked;
            };
            Controls.Add(alwaysOnTopCheckBox);

            var opacityLabel = new Label
            {
                Text = "ウィンドウ透明度:",
                Location = new Point(10, 220),
                AutoSize = true,
                Font = new Font("Arial", 9, FontStyle.Regular)
            };
            Controls.Add(opacityLabel);

            opacityTrackBar = new TrackBar
            {
                Location = new Point(150, 215),
                Size = new Size(200, 45),
                Minimum = 20,
                Maximum = 100,
                Value = 100,
                TickFrequency = 10,
                LargeChange = 10,
                SmallChange = 1
            };
            opacityTrackBar.Scroll += (sender, e) =>
            {
                this.Opacity = opacityTrackBar.Value / 100.0;
            };
            Controls.Add(opacityTrackBar);

            if (!RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_SNAPSHOT))
            {
                MessageBox.Show("ホットキーの登録に失敗しました。終了します。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                Environment.Exit(0);
            }
        }

        // DPI スケール倍率を取得
        private float GetDpiScale()
        {
            using (Graphics g = this.CreateGraphics())
            {
                return g.DpiX / 96.0f; // 標準DPI(96)を基準とする
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnFormClosed(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                TakeScreenshot();
            }
            base.WndProc(ref m);
        }

        private void CopyFolderPathToClipboard()
        {
            if (!string.IsNullOrEmpty(saveFolder))
            {
                Clipboard.SetText(saveFolder);
                MessageBox.Show("フォルダパスがクリップボードにコピーされました。");
            }
            else
            {
                MessageBox.Show("保存フォルダが設定されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //
        private async Task SaveClipboardImageToFolder()
        {
            try
            {
                if (Clipboard.ContainsImage())
                {
                    Image image = Clipboard.GetImage();
                    string fileName = GetUniqueFileName(screenshotsFolder, "ClipboardImage", ".png");
                    string filePath = Path.Combine(screenshotsFolder, fileName);

                    image.Save(filePath, ImageFormat.Png);

                    var selectedScreen = Screen.AllScreens[screenSelector.SelectedIndex];
                    Rectangle bounds = selectedScreen.Bounds;

                    await ShowNotificationOnScreen(
                        string.Format("クリップボードの画像を保存しました。\n{0}", filePath),
                        Color.LightGreen,
                        3000,
                        bounds
                    );

                    if (autoGenerateHtmlCheckBox.Checked)
                    {
                        await GenerateHtmlFile(titleInput.Text);
                        await ShowNotificationOnScreen("HTMLファイルを自動生成しました。", Color.LightGreen, 3000, bounds);
                    }
                }
                else
                {
                    MessageBox.Show("クリップボードに画像が保存されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("クリップボードの画像の保存に失敗しました: {0}", ex.Message), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void TakeScreenshot()
        {
            try
            {
                var selectedScreen = Screen.AllScreens[screenSelector.SelectedIndex];
                Rectangle bounds = selectedScreen.Bounds;

                // モニターのDPIを取得
                IntPtr hMonitor = MonitorFromPoint(new Point(bounds.Left, bounds.Top), 2);
                uint dpiX, dpiY;
                GetDpiForMonitor(hMonitor, MDT_EFFECTIVE_DPI, out dpiX, out dpiY);

                float scaleX = dpiX / 96.0f; // DPIスケール倍率
                float scaleY = dpiY / 96.0f;

                // DPIを考慮したウィンドウサイズを取得
                int width = bounds.Width;  // DPIスケール適用しない
                int height = bounds.Height;
                int x = bounds.X;
                int y = bounds.Y;

                using (Bitmap bitmap = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(x, y, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
                    }

                    string fileName = GetUniqueFileName(screenshotsFolder, "Screenshot_", ".png");
                    string filePath = Path.Combine(screenshotsFolder, fileName);

                    bitmap.Save(filePath, ImageFormat.Png);

                    await ShowNotificationOnScreen(
                        string.Format("スクリーンショットを保存しました。\n{0}", filePath),
                        Color.LightGreen,
                        3000,
                        bounds
                    );

                    if (autoGenerateHtmlCheckBox.Checked)
                    {
                        await GenerateHtmlFile(titleInput.Text);
                        await ShowNotificationOnScreen(
                            "HTMLファイルを自動生成しました。",
                            Color.LightGreen,
                            3000,
                            bounds
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("スクリーンショットの保存に失敗しました: {0}", ex.Message),
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification
                );
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(Point pt, uint dwFlags);

        private async Task ShowNotificationOnScreen(string message, Color backgroundColor, int durationMs, Rectangle screenBounds)
        {
            Form notification = new Form
            {
                Size = new Size(400, 100),
                FormBorderStyle = FormBorderStyle.None,
                TopMost = true,
                BackColor = backgroundColor
            };

            Label label = new Label
            {
                Text = message,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Arial", 12, FontStyle.Bold)
            };
            notification.Controls.Add(label);

            notification.StartPosition = FormStartPosition.Manual;
            notification.Location = new Point(
                screenBounds.Right - notification.Width - 10,
                screenBounds.Bottom - notification.Height - 10
            );

            notification.Show();
            await Task.Delay(durationMs);
            notification.Close();
        }

        private string GetUniqueFileName(string folderPath, string baseName, string extension)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            return timestamp + "_" + baseName + extension;
        }

        private async Task GenerateHtmlFile(string title)
        {
            await Task.Run(() =>
            {
                try
                {
                    string htmlFilePath = Path.Combine(saveFolder, "index.html");

                    var files = new List<string>();
                    files.AddRange(Directory.GetFiles(screenshotsFolder, "*.png"));
                    files.AddRange(Directory.GetFiles(screenshotsFolder, "*.jpg"));
                    files.AddRange(Directory.GetFiles(screenshotsFolder, "*.jpeg"));
                    files.AddRange(Directory.GetFiles(screenshotsFolder, "*.bmp"));

                    files.Sort((x, y) => File.GetLastWriteTime(x).CompareTo(File.GetLastWriteTime(y)));

                    using (StreamWriter writer = new StreamWriter(htmlFilePath, false, System.Text.Encoding.UTF8))
                    {
                        writer.WriteLine("<!DOCTYPE html>");
                        writer.WriteLine("<html lang=\"ja\">");
                        writer.WriteLine("<head>");
                        writer.WriteLine("<meta charset=\"UTF-8\">");
                        writer.WriteLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
                        writer.WriteLine(string.Format("<title>{0}</title>", title));
                        writer.WriteLine("<style>");
                        writer.WriteLine("body { font-family: Arial, sans-serif; margin: 20px; }");
                        writer.WriteLine("img { max-width: 800px; display: block; margin: 10px 0; }");
                        writer.WriteLine("</style>");
                        writer.WriteLine("</head>");
                        writer.WriteLine("<body contenteditable=\"true\">");
                        writer.WriteLine(string.Format("<h1>{0}</h1>", title));

                        foreach (string file in files)
                        {
                            string fileName = Path.GetFileName(file);
                            DateTime lastWriteTime = File.GetLastWriteTime(file);

                            writer.WriteLine("<div>");
                            writer.WriteLine(string.Format("<p>▼ {0} - 更新日時: {1}</p>", fileName, lastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")));
                            writer.WriteLine(string.Format("<img src=\"{0}\" alt=\"{1}\">", "img/" + fileName, fileName));
                            writer.WriteLine("</div>");
                        }

                        writer.WriteLine("</body>");
                        writer.WriteLine("</html>");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("HTMLファイルの生成に失敗しました: {0}", ex.Message),
                                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                }
            });
        }

        private void UpdateDateTime()
        {
            dateTimeLabel.Text = string.Format("現在日時: {0:yyyy-MM-dd HH:mm:ss}", DateTime.Now);
        }

        private string GetFolderWithFileDialog()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.ValidateNames = false;
                dialog.CheckFileExists = false;
                dialog.CheckPathExists = true;
                dialog.FileName = "フォルダを選択してください";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return Path.GetDirectoryName(dialog.FileName);
                }
            }
            return null;
        }
    }
}