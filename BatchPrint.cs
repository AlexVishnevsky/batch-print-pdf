// Пакетний друк PDF: друкує PDF з папки пачками, чекає завершення друку на принтері,
// робить паузу і переносить надруковані файли в іншу папку.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace BatchPrint
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    class Settings
    {
        public string Source = "", Done = "", Printer = "", Ip = "", Community = "public";
        public int Duplex = 1, Color = 1, Dpi = 300, Batch = 50, PauseMin = 5, Total = 0, MaxQueue = 3, FallbackSec = 120;

        static string PathFile
        {
            get
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BatchPrint");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "settings.ini");
            }
        }

        public static Settings Load()
        {
            var s = new Settings();
            try
            {
                if (!File.Exists(PathFile)) return s;
                foreach (var line in File.ReadAllLines(PathFile, Encoding.UTF8))
                {
                    int i = line.IndexOf('=');
                    if (i <= 0) continue;
                    string k = line.Substring(0, i), v = line.Substring(i + 1);
                    int n; int.TryParse(v, out n);
                    switch (k)
                    {
                        case "Source": s.Source = v; break;
                        case "Done": s.Done = v; break;
                        case "Printer": s.Printer = v; break;
                        case "Ip": s.Ip = v; break;
                        case "Community": s.Community = v; break;
                        case "Duplex": s.Duplex = n; break;
                        case "Color": s.Color = n; break;
                        case "Dpi": s.Dpi = n; break;
                        case "Batch": s.Batch = n; break;
                        case "PauseMin": s.PauseMin = n; break;
                        case "Total": s.Total = n; break;
                        case "MaxQueue": s.MaxQueue = n; break;
                        case "FallbackSec": s.FallbackSec = n; break;
                    }
                }
            }
            catch { }
            return s;
        }

        public void Save()
        {
            try
            {
                File.WriteAllLines(PathFile, new[] {
                    "Source=" + Source, "Done=" + Done, "Printer=" + Printer, "Ip=" + Ip, "Community=" + Community,
                    "Duplex=" + Duplex, "Color=" + Color, "Dpi=" + Dpi, "Batch=" + Batch, "PauseMin=" + PauseMin,
                    "Total=" + Total, "MaxQueue=" + MaxQueue, "FallbackSec=" + FallbackSec }, Encoding.UTF8);
            }
            catch { }
        }
    }

    class MainForm : Form
    {
        SkField tbSrc, tbDone, tbIp, logField;
        TextBox tbLog;
        SkSelect cbPrinter, cbDuplex, cbColor, cbDpi;
        SkStepper nBatch, nPause, nTotal, nQueue, nFallback;
        SkButton bStart, bPause, bStop, bIp, bSrc, bDone;
        Label lCount, lStatus;
        SkProgress progress;

        Thread worker;
        volatile bool stopReq, pauseReq;
        Settings cfg;

        static readonly Rectangle P1 = new Rectangle(14, 12, 632, 140), P2 = new Rectangle(14, 164, 632, 138),
                                  P3 = new Rectangle(14, 314, 632, 138);
        Rectangle P4 { get { return new Rectangle(14, 518, 632, ClientSize.Height - 518 - 14); } }

        public MainForm()
        {
            Text = "Пакетний друк PDF";
            Font = new Font("Segoe UI", 9.5f);
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            ClientSize = new Size(660, 770);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(24, 24, 27);
            ForeColor = Sk.Text;
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            AddLabel("ПАПКА З PDF ДЛЯ ДРУКУ", 30, 24);
            tbSrc = Field(30, 46, 500);
            bSrc = Btn("Огляд…", 540, 46, 92, 30, (s, e) => Browse(tbSrc));
            tbSrc.Box.TextChanged += (s, e) => UpdateCount();
            AddLabel("ПАПКА ДЛЯ НАДРУКОВАНИХ PDF", 30, 86);
            tbDone = Field(30, 108, 500);
            bDone = Btn("Огляд…", 540, 108, 92, 30, (s, e) => Browse(tbDone));

            AddLabel("ПРИНТЕР", 30, 176);
            AddLabel("IP ПРИНТЕРА", 370, 176);
            cbPrinter = Sel(30, 198, 326);
            foreach (string p in PrinterSettings.InstalledPrinters) cbPrinter.Items.Add(p);
            cbPrinter.SelectedIndexChanged += (s, e) => { tbIp.Text = Printing.DetectIp(cbPrinter.Text); };
            tbIp = Field(370, 198, 150);
            bIp = Btn("Перевірити", 528, 198, 104, 30, (s, e) => CheckIp());

            AddLabel("СТОРОНИ", 30, 238); AddLabel("КОЛІР", 300, 238); AddLabel("ЯКІСТЬ (DPI)", 470, 238);
            cbDuplex = Sel(30, 260, 256, "Односторонній", "Двосторонній (довгий край)", "Двосторонній (короткий край)");
            cbColor = Sel(300, 260, 156, "Чорно-білий", "Кольоровий");
            cbDpi = Sel(470, 260, 162, "150 — швидко", "200", "300 — стандарт", "600 — найкраща");

            AddLabel("ФАЙЛІВ У ПАЧЦІ", 30, 326); AddLabel("ПАУЗА ПІСЛЯ ПАЧКИ, ХВ", 240, 326); AddLabel("ВСЬОГО (0 = ВСЯ ПАПКА)", 450, 326);
            nBatch = Stepper(30, 348, 1, 10000);
            nPause = Stepper(240, 348, 0, 600);
            nTotal = Stepper(450, 348, 0, 1000000);
            AddLabel("У ЧЕРЗІ WINDOWS ОДНОЧАСНО", 30, 388); AddLabel("ЯКЩО ПРИНТЕР МОВЧИТЬ, С", 240, 388);
            nQueue = Stepper(30, 410, 1, 100);
            nFallback = Stepper(240, 410, 0, 3600);
            lCount = new Label { Left = 450, Top = 415, Width = 180, Height = 22, BackColor = Color.Transparent, ForeColor = Sk.Text, Font = new Font(Font, FontStyle.Bold) };
            Controls.Add(lCount);

            bStart = Btn("▶   Старт", 14, 466, 200, 40, (s, e) => StartRun());
            bStart.Accent = true;
            bPause = Btn("❚❚   Пауза", 230, 466, 200, 40, (s, e) => TogglePause());
            bStop = Btn("■   Стоп", 446, 466, 200, 40, (s, e) => RequestStop());
            bStart.Font = bPause.Font = bStop.Font = new Font("Segoe UI Semibold", 11f);
            bPause.Enabled = bStop.Enabled = false;

            Rectangle p4 = P4;
            lStatus = new Label { Left = 30, Top = p4.Top + 12, Width = 600, Height = 22, Text = "Готово до роботи", BackColor = Color.Transparent, ForeColor = Sk.Text, Font = new Font("Segoe UI Semibold", 10.5f) };
            Controls.Add(lStatus);
            progress = new SkProgress { Left = 30, Top = p4.Top + 40, Width = 600, Height = 16 };
            Controls.Add(progress);
            logField = new SkField(true) { Left = 30, Top = p4.Top + 66, Width = 600, Height = p4.Bottom - (p4.Top + 66) - 16 };
            tbLog = logField.Box;
            tbLog.Font = new Font("Consolas", 9f);
            Controls.Add(logField);

            LoadSettings();
            FormClosing += OnClosing;
        }

        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr h, int attr, ref int val, int size);

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            int on = 1;
            try { DwmSetWindowAttribute(Handle, 20, ref on, 4); } catch { } // темний заголовок вікна
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            var g = e.Graphics;
            using (var br = new System.Drawing.Drawing2D.LinearGradientBrush(ClientRectangle, Color.FromArgb(40, 40, 44), Color.FromArgb(14, 14, 16), 90f))
                g.FillRectangle(br, ClientRectangle);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            Sk.Plate(g, P1, 10); Sk.Plate(g, P2, 10); Sk.Plate(g, P3, 10); Sk.Plate(g, P4, 10);
        }

        // ---------- UI helpers ----------
        Label AddLabel(string t, int x, int y)
        {
            var l = new Label { Text = t, Left = x, Top = y, AutoSize = true, BackColor = Color.Transparent, ForeColor = Sk.LabelColor, Font = new Font("Segoe UI Semibold", 8f) };
            Controls.Add(l);
            return l;
        }

        SkButton Btn(string t, int x, int y, int w, int h, EventHandler click)
        {
            var b = new SkButton { Text = t, Left = x, Top = y, Width = w, Height = h };
            b.Click += click;
            Controls.Add(b);
            return b;
        }

        SkField Field(int x, int y, int w)
        {
            var f = new SkField(false) { Left = x, Top = y, Width = w, Height = 30 };
            Controls.Add(f);
            return f;
        }

        SkSelect Sel(int x, int y, int w, params string[] items)
        {
            var c = new SkSelect { Left = x, Top = y, Width = w, Height = 30 };
            c.Items.AddRange(items);
            Controls.Add(c);
            return c;
        }

        SkStepper Stepper(int x, int y, int min, int max)
        {
            var n = new SkStepper { Left = x, Top = y, Width = 160, Height = 30, Minimum = min, Maximum = max };
            Controls.Add(n);
            return n;
        }
        void Browse(SkField tb)
        {
            using (var d = new FolderBrowserDialog())
            {
                if (Directory.Exists(tb.Text)) d.SelectedPath = tb.Text;
                if (d.ShowDialog(this) == DialogResult.OK) tb.Text = d.SelectedPath;
            }
        }

        void UpdateCount()
        {
            try
            {
                lCount.Text = Directory.Exists(tbSrc.Text) ? "У папці: " + Directory.GetFiles(tbSrc.Text, "*.pdf").Length + " PDF" : "";
            }
            catch { lCount.Text = ""; }
        }

        static readonly int[] DpiValues = { 150, 200, 300, 600 };

        void LoadSettings()
        {
            cfg = Settings.Load();
            tbSrc.Text = cfg.Source;
            tbDone.Text = cfg.Done;
            int pi = cbPrinter.Items.IndexOf(cfg.Printer);
            if (pi < 0) pi = cbPrinter.Items.IndexOf(new PrinterSettings().PrinterName);
            if (pi >= 0) cbPrinter.SelectedIndex = pi;
            if (cfg.Ip != "") tbIp.Text = cfg.Ip;
            cbDuplex.SelectedIndex = Clamp(cfg.Duplex, 0, 2);
            cbColor.SelectedIndex = Clamp(cfg.Color, 0, 1);
            int di = Array.IndexOf(DpiValues, cfg.Dpi);
            cbDpi.SelectedIndex = di < 0 ? 2 : di;
            nBatch.Value = Clamp(cfg.Batch, 1, 10000);
            nPause.Value = Clamp(cfg.PauseMin, 0, 600);
            nTotal.Value = Clamp(cfg.Total, 0, 1000000);
            nQueue.Value = Clamp(cfg.MaxQueue, 1, 100);
            nFallback.Value = Clamp(cfg.FallbackSec, 0, 3600);
            UpdateCount();
        }

        void SaveSettings()
        {
            cfg.Source = tbSrc.Text.Trim(); cfg.Done = tbDone.Text.Trim(); cfg.Printer = cbPrinter.Text; cfg.Ip = tbIp.Text.Trim();
            cfg.Duplex = cbDuplex.SelectedIndex; cfg.Color = cbColor.SelectedIndex; cfg.Dpi = DpiValues[cbDpi.SelectedIndex];
            cfg.Batch = (int)nBatch.Value; cfg.PauseMin = (int)nPause.Value; cfg.Total = (int)nTotal.Value;
            cfg.MaxQueue = (int)nQueue.Value; cfg.FallbackSec = (int)nFallback.Value;
            cfg.Save();
        }

        static int Clamp(int v, int a, int b) { return v < a ? a : (v > b ? b : v); }

        void CheckIp()
        {
            string ip = tbIp.Text.Trim();
            if (ip == "") { MessageBox.Show(this, "Вкажіть IP-адресу принтера.", Text); return; }
            Cursor = Cursors.WaitCursor;
            long? st = Snmp.Get(ip, cfg.Community, Snmp.HrPrinterStatus);
            long? cnt = Snmp.Get(ip, cfg.Community, Snmp.PageCounter);
            Cursor = Cursors.Default;
            if (st == null && cnt == null)
                MessageBox.Show(this, "Принтер не відповів по SNMP.\n\nУтиліта працюватиме, але після черги Windows чекатиме фіксований час («Якщо принтер мовчить»).", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                MessageBox.Show(this, "Принтер відповідає.\nСтан: " + Snmp.StatusText(st) + "\nЛічильник сторінок: " + (cnt.HasValue ? cnt.Value.ToString() : "—"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void Log(string msg)
        {
            string line = DateTime.Now.ToString("HH:mm:ss") + "  " + msg;
            Ui(() => { tbLog.AppendText(line + Environment.NewLine); });
            try
            {
                if (Directory.Exists(cfg.Done))
                    File.AppendAllText(Path.Combine(cfg.Done, "журнал друку.txt"), DateTime.Now.ToString("yyyy-MM-dd ") + line + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
        }

        void Status(string s) { Ui(() => lStatus.Text = s); }

        void Ui(Action a)
        {
            if (IsDisposed) return;
            try { if (InvokeRequired) BeginInvoke(a); else a(); } catch { }
        }

        void SetRunning(bool running)
        {
            Ui(() =>
            {
                bStart.Enabled = !running;
                bPause.Enabled = bStop.Enabled = running;
                bPause.Text = "❚❚  Пауза";
                foreach (var c in new SkField[] { tbSrc, tbDone, tbIp }) c.Locked = running;
                foreach (var c in new SkStepper[] { nBatch, nPause, nTotal, nQueue, nFallback }) c.Locked = running;
                foreach (Control c in new Control[] { cbPrinter, bIp, cbDuplex, cbColor, cbDpi, bSrc, bDone }) c.Enabled = !running;
                if (!running) UpdateCount();
            });
        }

        void TogglePause()
        {
            pauseReq = !pauseReq;
            bPause.Text = pauseReq ? "▶  Продовжити" : "❚❚  Пауза";
            Log(pauseReq ? "Пауза (поточний файл допрацює)." : "Продовжуємо.");
        }

        void RequestStop()
        {
            stopReq = true;
            pauseReq = false;
            bStop.Enabled = false;
            Log("Зупинка: допрацює поточний файл і зупиниться.");
        }

        void OnClosing(object s, FormClosingEventArgs e)
        {
            if (worker != null && worker.IsAlive)
            {
                if (MessageBox.Show(this, "Друк ще йде. Закрити програму?\nФайли, які вже в черзі принтера, допрацюють.", Text,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) { e.Cancel = true; return; }
                stopReq = true;
            }
            SaveSettings();
        }

        // ---------- Запуск ----------
        void StartRun()
        {
            string src = tbSrc.Text.Trim(), done = tbDone.Text.Trim();
            if (!Directory.Exists(src)) { MessageBox.Show(this, "Папка з PDF не існує.", Text); return; }
            if (done == "") { MessageBox.Show(this, "Вкажіть папку для надрукованих.", Text); return; }
            if (string.Equals(Path.GetFullPath(src).TrimEnd('\\'), Path.GetFullPath(done).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
            { MessageBox.Show(this, "Папки мають бути різні.", Text); return; }
            if (cbPrinter.SelectedIndex < 0) { MessageBox.Show(this, "Оберіть принтер.", Text); return; }
            try { Directory.CreateDirectory(done); } catch (Exception ex) { MessageBox.Show(this, "Не вдалося створити папку: " + ex.Message, Text); return; }

            SaveSettings();
            var ps = new PrinterSettings { PrinterName = cfg.Printer };
            if (cfg.Duplex > 0 && !ps.CanDuplex)
                Log("Увага: драйвер не повідомляє про двосторонній друк — перевірте, що в налаштуваннях принтера увімкнено дуплекс-модуль.");

            stopReq = false; pauseReq = false;
            tbLog.Clear();
            SetRunning(true);
            var job = new Job(cfg, this);
            worker = new Thread(job.Run) { IsBackground = true };
            worker.Start();
        }

        // Робота у фоновому потоці
        class Job
        {
            readonly Settings c;
            readonly MainForm f;
            public Job(Settings settings, MainForm form)
            {
                // копія налаштувань, щоб зміни у вікні не впливали на запущений друк
                c = (Settings)typeof(Settings).GetMethod("MemberwiseClone", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(settings, null);
                f = form;
            }

            bool Stop { get { return f.stopReq; } }

            void WaitWhilePaused()
            {
                bool shown = false;
                while (f.pauseReq && !f.stopReq)
                {
                    if (!shown) { f.Status("Пауза"); shown = true; }
                    Thread.Sleep(300);
                }
            }

            public void Run()
            {
                int printed = 0, failed = 0, batchNo = 0;
                var skip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                Duplex dup = c.Duplex == 0 ? Duplex.Simplex : (c.Duplex == 1 ? Duplex.Vertical : Duplex.Horizontal);
                bool color = c.Color == 1;
                try
                {
                    int available = Printing.ListPdfs(c.Source).Count;
                    int target = c.Total > 0 ? Math.Min(c.Total, available) : available;
                    f.Log("Старт: " + available + " PDF у папці, до друку " + target + ". Принтер: " + c.Printer +
                          ". " + (c.Duplex == 0 ? "Односторонній" : "Двосторонній") + ", " + (color ? "кольоровий" : "чорно-білий") + ", " + c.Dpi + " DPI.");
                    f.Ui(() => { f.progress.Maximum = Math.Max(1, target); f.progress.Value = 0; });

                    while (printed < target && !Stop)
                    {
                        var files = Printing.ListPdfs(c.Source).Where(p => !skip.Contains(p)).Take(Math.Min(c.Batch, target - printed)).ToList();
                        if (files.Count == 0) break;
                        batchNo++;
                        f.Log("Пачка " + batchNo + ": " + files.Count + " файлів.");
                        long? counterStart = string.IsNullOrEmpty(c.Ip) ? null : Snmp.Get(c.Ip, c.Community, Snmp.PageCounter);

                        int inBatch = 0;
                        foreach (var file in files)
                        {
                            WaitWhilePaused();
                            if (Stop) break;
                            // не завалювати чергу: чекаємо, поки в ній стане менше заданої кількості
                            while (!Stop && Printing.QueueCount(c.Printer) >= c.MaxQueue)
                            {
                                f.Status("Пачка " + batchNo + ": чекаю, поки черга Windows звільниться…");
                                Thread.Sleep(1000);
                            }
                            if (Stop) break;
                            string name = Path.GetFileName(file);
                            f.Status("Пачка " + batchNo + ": друкую " + (inBatch + 1) + " з " + files.Count + " — " + name);
                            try
                            {
                                int pages = Printing.PrintPdf(file, c.Printer, dup, color, c.Dpi, Path.Combine(c.Done, "_тест_pdf"));
                                Printing.MoveUnique(file, c.Done);
                                printed++; inBatch++;
                                f.Log("  ✓ " + name + " (" + pages + " стор.)");
                            }
                            catch (Exception ex)
                            {
                                failed++;
                                skip.Add(file);
                                string err = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                                f.Log("  ✗ " + name + " — помилка: " + err);
                                try { Printing.MoveUnique(file, Path.Combine(c.Done, "_помилки")); f.Log("    файл перенесено в «_помилки»"); } catch { }
                                if (failed >= 5 && printed == 0) { f.Log("Забагато помилок підряд — зупиняюсь. Перевірте принтер."); f.stopReq = true; }
                            }
                            int pv = printed;
                            f.Ui(() => f.progress.Value = Math.Min(f.progress.Maximum, pv));
                        }

                        if (inBatch == 0) continue;
                        WaitPrinterDone(batchNo);
                        if (counterStart.HasValue)
                        {
                            long? counterEnd = Snmp.Get(c.Ip, c.Community, Snmp.PageCounter);
                            if (counterEnd.HasValue) f.Log("Пачка " + batchNo + " завершена. Лічильник принтера: +" + (counterEnd.Value - counterStart.Value) + " стор.");
                        }
                        else f.Log("Пачка " + batchNo + " завершена.");

                        bool more = printed < target && Printing.ListPdfs(c.Source).Any(p => !skip.Contains(p));
                        if (more && !Stop && c.PauseMin > 0) Countdown(c.PauseMin * 60);
                    }
                }
                catch (Exception ex)
                {
                    f.Log("Аварійна зупинка: " + ex.Message);
                }
                string summary = (Stop ? "Зупинено. " : "Готово. ") + "Надруковано файлів: " + printed + (failed > 0 ? ", з помилкою: " + failed : "") + ".";
                f.Log(summary);
                f.Status(summary);
                f.SetRunning(false);
            }

            void Countdown(int seconds)
            {
                DateTime end = DateTime.Now.AddSeconds(seconds);
                f.Log("Пауза " + (seconds / 60) + " хв до наступної пачки.");
                while (!Stop)
                {
                    if (f.pauseReq)
                    {
                        // на паузі таймер стоїть
                        double left = (end - DateTime.Now).TotalSeconds;
                        WaitWhilePaused();
                        end = DateTime.Now.AddSeconds(left);
                        continue;
                    }
                    TimeSpan rem = end - DateTime.Now;
                    if (rem.TotalSeconds <= 0) break;
                    f.Status("Наступна пачка через " + ((int)rem.TotalMinutes).ToString("00") + ":" + rem.Seconds.ToString("00"));
                    Thread.Sleep(250);
                }
            }

            // Чекаємо: спершу черга Windows порожня, потім сам принтер (SNMP) закінчив друкувати
            void WaitPrinterDone(int batchNo)
            {
                bool reportedErr = false;
                while (true)
                {
                    string err;
                    int n = Printing.QueueCount(c.Printer, out err);
                    if (n <= 0) break;
                    if (err != null && !reportedErr) { f.Log("Черга Windows повідомляє: " + err); reportedErr = true; }
                    f.Status("Пачка " + batchNo + ": у черзі Windows ще " + n + " завд." + (err != null ? " (" + err + ")" : ""));
                    Thread.Sleep(2000);
                }

                if (string.IsNullOrEmpty(c.Ip)) { FixedWait(batchNo, "IP принтера не вказано"); return; }

                long? lastCnt = null;
                DateTime lastChange = DateTime.Now;
                bool errLogged = false;
                int silent = 0;
                while (true)
                {
                    long? st = Snmp.Get(c.Ip, c.Community, Snmp.HrPrinterStatus);
                    long? cnt = Snmp.Get(c.Ip, c.Community, Snmp.PageCounter);
                    long? dev = Snmp.Get(c.Ip, c.Community, Snmp.HrDeviceStatus);
                    if (st == null && cnt == null)
                    {
                        if (++silent >= 3) { FixedWait(batchNo, "принтер не відповідає по SNMP"); return; }
                        Thread.Sleep(2000);
                        continue;
                    }
                    silent = 0;
                    if (cnt.HasValue && cnt != lastCnt) { lastCnt = cnt; lastChange = DateTime.Now; }
                    double still = (DateTime.Now - lastChange).TotalSeconds;

                    if (dev == 5) // down: застрягання, нема паперу, відкрита кришка
                    {
                        if (!errLogged) { f.Log("Принтер повідомляє про помилку (папір / застрягання / кришка?) — чекаю."); errLogged = true; }
                        f.Status("Пачка " + batchNo + ": принтер у помилці — усуньте, друк продовжиться");
                        lastChange = DateTime.Now;
                        Thread.Sleep(3000);
                        continue;
                    }
                    if (errLogged) { f.Log("Помилку принтера усунуто."); errLogged = false; }

                    bool printing = st == 4 || st == 5;
                    f.Status("Пачка " + batchNo + ": принтер допрацьовує — " + Snmp.StatusText(st) + (cnt.HasValue ? ", лічильник " + cnt.Value : ""));
                    if (!printing && still >= 15) return;
                    if (st == null && still >= 30) return;
                    Thread.Sleep(3000);
                }
            }

            void FixedWait(int batchNo, string why)
            {
                if (c.FallbackSec <= 0) return;
                f.Log("Черга Windows порожня; " + why + " — чекаю ще " + c.FallbackSec + " с, поки принтер допрацює.");
                DateTime end = DateTime.Now.AddSeconds(c.FallbackSec);
                while (DateTime.Now < end)
                {
                    f.Status("Пачка " + batchNo + ": даю принтеру допрацювати, ще " + (int)(end - DateTime.Now).TotalSeconds + " с");
                    Thread.Sleep(500);
                }
            }
        }
    }


    // ================= Чорний скевоморфізм: власне малювання всіх елементів =================
    static class Sk
    {
        public static readonly Color Text = Color.FromArgb(232, 232, 236);
        public static readonly Color TextDim = Color.FromArgb(120, 120, 128);
        public static readonly Color LabelColor = Color.FromArgb(150, 150, 158);
        public static readonly Color WellBack = Color.FromArgb(15, 15, 17);

        public static System.Drawing.Drawing2D.GraphicsPath Round(RectangleF r, float rad)
        {
            var p = new System.Drawing.Drawing2D.GraphicsPath();
            float d = rad * 2;
            if (d > r.Height) d = r.Height;
            if (d > r.Width) d = r.Width;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        static System.Drawing.Drawing2D.LinearGradientBrush Grad(RectangleF r, Color top, Color bottom)
        {
            return new System.Drawing.Drawing2D.LinearGradientBrush(new RectangleF(r.X, r.Y - 1, r.Width, r.Height + 2), top, bottom, 90f);
        }

        // Піднята панель
        public static void Plate(Graphics g, Rectangle rc, float rad)
        {
            RectangleF r = rc;
            for (int i = 3; i >= 1; i--)
                using (var p = Round(new RectangleF(r.X - i + 1, r.Y + i, r.Width + 2 * i - 2, r.Height), rad + i))
                using (var b = new SolidBrush(Color.FromArgb(28, 0, 0, 0))) g.FillPath(b, p);
            using (var p = Round(r, rad))
            {
                using (var b = Grad(r, Color.FromArgb(38, 38, 42), Color.FromArgb(25, 25, 28))) g.FillPath(b, p);
                using (var pen = new Pen(Color.FromArgb(6, 6, 7))) g.DrawPath(pen, p);
            }
            using (var p = Round(new RectangleF(r.X + 1, r.Y + 1, r.Width - 2, r.Height - 2), rad - 1))
            using (var pen = new Pen(new System.Drawing.Drawing2D.LinearGradientBrush(r, Color.FromArgb(34, 255, 255, 255), Color.FromArgb(0, 255, 255, 255), 90f)))
                g.DrawPath(pen, p);
        }

        // Заглиблене поле
        public static void Well(Graphics g, RectangleF r, float rad)
        {
            // світлий «уступ» під полем
            using (var p = Round(new RectangleF(r.X, r.Y + 1, r.Width, r.Height), rad))
            using (var pen = new Pen(Color.FromArgb(22, 255, 255, 255))) g.DrawPath(pen, p);
            using (var p = Round(r, rad))
            {
                using (var b = Grad(r, Color.FromArgb(8, 8, 9), Color.FromArgb(19, 19, 21))) g.FillPath(b, p);
                var old = g.Clip;
                g.SetClip(p, System.Drawing.Drawing2D.CombineMode.Intersect);
                for (int i = 0; i < 4; i++) // внутрішня тінь зверху
                    using (var pen = new Pen(Color.FromArgb(70 - i * 16, 0, 0, 0)))
                        g.DrawLine(pen, r.X, r.Y + i, r.Right, r.Y + i);
                g.Clip = old;
                using (var pen = new Pen(Color.FromArgb(2, 2, 3))) g.DrawPath(pen, p);
            }
        }

        // Опукла кнопка з глянцем
        public static void Raised(Graphics g, RectangleF r, float rad, bool hot, bool down, bool accent)
        {
            using (var p = Round(new RectangleF(r.X, r.Y + 1.5f, r.Width, r.Height), rad))
            using (var b = new SolidBrush(Color.FromArgb(90, 0, 0, 0))) g.FillPath(b, p);
            int k = hot && !down ? 10 : 0;
            Color top = down ? Color.FromArgb(30, 30, 33) : Color.FromArgb(78 + k, 78 + k, 84 + k);
            Color bot = down ? Color.FromArgb(48, 48, 52) : Color.FromArgb(36 + k, 36 + k, 40 + k);
            if (accent && !down) { top = Color.FromArgb(92 + k, 94 + k, 100 + k); bot = Color.FromArgb(44 + k, 45 + k, 50 + k); }
            using (var p = Round(r, rad))
            {
                using (var b = Grad(r, top, bot)) g.FillPath(b, p);
                if (!down)
                {
                    var old = g.Clip;
                    g.SetClip(p, System.Drawing.Drawing2D.CombineMode.Intersect);
                    var gloss = new RectangleF(r.X, r.Y, r.Width, r.Height * 0.5f);
                    using (var gp = Round(new RectangleF(r.X + 1, r.Y + 1, r.Width - 2, r.Height * 0.5f), rad - 1))
                    using (var b = Grad(gloss, Color.FromArgb(46, 255, 255, 255), Color.FromArgb(8, 255, 255, 255))) g.FillPath(b, gp);
                    g.Clip = old;
                }
                using (var pen = new Pen(Color.FromArgb(4, 4, 5))) g.DrawPath(pen, p);
            }
            if (!down)
                using (var p = Round(new RectangleF(r.X + 1, r.Y + 1, r.Width - 2, r.Height - 2), rad - 1))
                using (var pen = new Pen(new System.Drawing.Drawing2D.LinearGradientBrush(r, Color.FromArgb(70, 255, 255, 255), Color.FromArgb(0, 255, 255, 255), 90f)))
                    g.DrawPath(pen, p);
        }

        public static void EmbossText(Graphics g, string t, Font f, Rectangle r, Color c, TextFormatFlags fl)
        {
            Rectangle sh = r; sh.Offset(0, 1);
            TextRenderer.DrawText(g, t, f, sh, Color.FromArgb(0, 0, 0), fl);
            TextRenderer.DrawText(g, t, f, r, c, fl);
        }

        public static void Prepare(Control c)
        {
            c.BackColor = Color.Transparent;
        }
    }

    class SkControl : Control
    {
        protected bool hot, down;
        public SkControl()
        {
            SetStyle(ControlStyles.SupportsTransparentBackColor | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.Transparent;
            ForeColor = Sk.Text;
            Cursor = Cursors.Hand;
        }
        protected override void OnMouseEnter(EventArgs e) { hot = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hot = false; down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left) { down = true; Invalidate(); } base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnTextChanged(EventArgs e) { Invalidate(); base.OnTextChanged(e); }
        protected Graphics Smooth(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            return e.Graphics;
        }
    }

    class SkButton : SkControl
    {
        public bool Accent;
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = Smooth(e);
            var r = new RectangleF(1, 0, Width - 2, Height - 3);
            Sk.Raised(g, r, 7, hot && Enabled, down && Enabled, Accent && Enabled);
            var tr = new Rectangle(0, down ? 1 : 0, Width, Height - 3);
            Sk.EmbossText(g, Text, Font, tr, Enabled ? Sk.Text : Sk.TextDim, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
            if (Accent && Enabled) // зелений «світлодіод»
            {
                var led = new RectangleF(14, Height / 2f - 5.5f, 9, 9);
                using (var b = new SolidBrush(Color.FromArgb(60, 90, 255, 120))) g.FillEllipse(b, RectangleF.Inflate(led, 3, 3));
                using (var b = new System.Drawing.Drawing2D.LinearGradientBrush(led, Color.FromArgb(170, 255, 190), Color.FromArgb(40, 190, 90), 90f)) g.FillEllipse(b, led);
                using (var pen = new Pen(Color.FromArgb(10, 40, 20))) g.DrawEllipse(pen, led);
            }
        }
    }

    class DarkColors : ProfessionalColorTable
    {
        static readonly Color Bg = Color.FromArgb(30, 30, 34), Sel = Color.FromArgb(62, 62, 68);
        public override Color MenuItemSelected { get { return Sel; } }
        public override Color MenuItemSelectedGradientBegin { get { return Sel; } }
        public override Color MenuItemSelectedGradientEnd { get { return Sel; } }
        public override Color MenuItemBorder { get { return Color.FromArgb(90, 90, 98); } }
        public override Color MenuBorder { get { return Color.FromArgb(4, 4, 5); } }
        public override Color ToolStripDropDownBackground { get { return Bg; } }
        public override Color ImageMarginGradientBegin { get { return Bg; } }
        public override Color ImageMarginGradientMiddle { get { return Bg; } }
        public override Color ImageMarginGradientEnd { get { return Bg; } }
        public override Color CheckBackground { get { return Color.FromArgb(50, 50, 56); } }
        public override Color CheckSelectedBackground { get { return Color.FromArgb(70, 70, 78); } }
        public override Color CheckPressedBackground { get { return Color.FromArgb(70, 70, 78); } }
    }

    class SkSelect : SkControl
    {
        public readonly List<string> Items = new List<string>();
        int sel = -1;
        public event EventHandler SelectedIndexChanged;

        public int SelectedIndex
        {
            get { return sel; }
            set
            {
                if (value == sel) return;
                sel = value;
                Invalidate();
                if (SelectedIndexChanged != null) SelectedIndexChanged(this, EventArgs.Empty);
            }
        }
        public override string Text { get { return sel >= 0 && sel < Items.Count ? Items[sel] : ""; } set { } }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            var m = new ContextMenuStrip { Renderer = new ToolStripProfessionalRenderer(new DarkColors()) { RoundedEdges = false }, Font = Font, ShowImageMargin = false, ShowCheckMargin = true };
            for (int i = 0; i < Items.Count; i++)
            {
                int idx = i;
                var it = new ToolStripMenuItem(Items[i]) { ForeColor = Sk.Text, Checked = i == sel, Padding = new Padding(0, 3, 0, 3) };
                it.Click += (s, a) => SelectedIndex = idx;
                m.Items.Add(it);
            }
            m.MinimumSize = new Size(Width, 0);
            m.Closed += (s, a) => BeginInvoke((Action)(() => m.Dispose()));
            m.Show(this, new Point(0, Height - 1));
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = Smooth(e);
            Sk.Raised(g, new RectangleF(1, 0, Width - 2, Height - 3), 7, hot && Enabled, down && Enabled, false);
            var tr = new Rectangle(12, 0, Width - 40, Height - 3);
            Sk.EmbossText(g, Text, Font, tr, Enabled ? Sk.Text : Sk.TextDim, TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
            // шеврон
            float cx = Width - 18, cy = (Height - 3) / 2f + 1;
            using (var pen = new Pen(Color.FromArgb(0, 0, 0), 2f)) g.DrawLines(pen, new[] { new PointF(cx - 4, cy - 1), new PointF(cx, cy + 3), new PointF(cx + 4, cy - 1) });
            using (var pen = new Pen(Enabled ? Sk.Text : Sk.TextDim, 1.6f)) g.DrawLines(pen, new[] { new PointF(cx - 4, cy - 2), new PointF(cx, cy + 2), new PointF(cx + 4, cy - 2) });
            using (var pen = new Pen(Color.FromArgb(10, 10, 11))) g.DrawLine(pen, Width - 34, 6, Width - 34, Height - 9);
            using (var pen = new Pen(Color.FromArgb(30, 255, 255, 255))) g.DrawLine(pen, Width - 33, 6, Width - 33, Height - 9);
        }
    }

    // Заглиблене поле з текстом
    class SkField : SkControl
    {
        public readonly TextBox Box;
        readonly bool multi;
        bool locked;

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] static extern int SetWindowTheme(IntPtr h, string app, string id);

        public SkField(bool multiline)
        {
            multi = multiline;
            Cursor = Cursors.IBeam;
            Box = new TextBox { BorderStyle = BorderStyle.None, BackColor = Sk.WellBack, ForeColor = Sk.Text, Multiline = multiline };
            if (multiline)
            {
                Box.ReadOnly = true;
                Box.ScrollBars = ScrollBars.Vertical;
                Box.HandleCreated += (s, e) => { try { SetWindowTheme(Box.Handle, "DarkMode_Explorer", null); } catch { } };
            }
            Controls.Add(Box);
        }

        public override string Text { get { return Box.Text; } set { Box.Text = value; } }

        public bool Locked
        {
            get { return locked; }
            set { locked = value; if (!multi) Box.ReadOnly = value; Box.ForeColor = value ? Sk.TextDim : Sk.Text; }
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            if (multi) Box.SetBounds(8, 7, Width - 10, Height - 13);
            else Box.SetBounds(10, (Height - Box.PreferredHeight) / 2 + 1, Width - 20, Box.PreferredHeight);
        }

        protected override void OnMouseDown(MouseEventArgs e) { Box.Focus(); base.OnMouseDown(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Sk.Well(Smooth(e), new RectangleF(0.5f, 0.5f, Width - 1, Height - 2.5f), 7);
        }
    }

    // Поле з числом і кнопками − / +
    class SkStepper : SkControl
    {
        readonly TextBox box;
        int val, min, max = 100;
        bool locked;
        int pressDir;
        readonly System.Windows.Forms.Timer rep = new System.Windows.Forms.Timer();

        public SkStepper()
        {
            box = new TextBox { BorderStyle = BorderStyle.None, BackColor = Sk.WellBack, ForeColor = Sk.Text, Text = "0", Font = new Font("Segoe UI Semibold", 10f) };
            box.KeyPress += (s, e) => { if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar)) e.Handled = true; };
            box.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { Commit(); e.SuppressKeyPress = true; } else if (e.KeyCode == Keys.Up) { Step(1); e.Handled = true; } else if (e.KeyCode == Keys.Down) { Step(-1); e.Handled = true; } };
            box.Leave += (s, e) => Commit();
            box.MouseWheel += (s, e) => Step(e.Delta > 0 ? 1 : -1);
            Controls.Add(box);
            rep.Tick += (s, e) => { rep.Interval = 60; Step(pressDir); };
        }

        public int Minimum { get { return min; } set { min = value; Value = val; } }
        public int Maximum { get { return max; } set { max = value; Value = val; } }
        public int Value
        {
            get { Commit(); return val; }
            set { val = Math.Max(min, Math.Min(max, value)); box.Text = val.ToString(); }
        }
        public bool Locked
        {
            get { return locked; }
            set { locked = value; box.ReadOnly = value; box.ForeColor = value ? Sk.TextDim : Sk.Text; Invalidate(); }
        }

        void Commit()
        {
            int v;
            if (!int.TryParse(box.Text, out v)) v = val;
            v = Math.Max(min, Math.Min(max, v));
            val = v;
            if (box.Text != v.ToString()) box.Text = v.ToString();
        }

        void Step(int d)
        {
            if (locked) return;
            Commit();
            int stepSize = max > 1000 ? 10 : 1;
            if (Control.ModifierKeys == Keys.Shift) stepSize *= 10;
            Value = val + d * stepSize;
        }

        Rectangle MinusRect { get { return new Rectangle(Width - 60, 3, 28, Height - 9); } }
        Rectangle PlusRect { get { return new Rectangle(Width - 31, 3, 28, Height - 9); } }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            box.SetBounds(10, (Height - box.PreferredHeight) / 2, Width - 76, box.PreferredHeight);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            pressDir = MinusRect.Contains(e.Location) ? -1 : PlusRect.Contains(e.Location) ? 1 : 0;
            if (pressDir == 0) { box.Focus(); return; }
            Step(pressDir);
            rep.Interval = 400;
            rep.Start();
        }
        protected override void OnMouseUp(MouseEventArgs e) { rep.Stop(); pressDir = 0; base.OnMouseUp(e); }
        protected override void OnMouseLeave(EventArgs e) { rep.Stop(); pressDir = 0; base.OnMouseLeave(e); }
        protected override void OnMouseWheel(MouseEventArgs e) { Step(e.Delta > 0 ? 1 : -1); base.OnMouseWheel(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = Smooth(e);
            Sk.Well(g, new RectangleF(0.5f, 0.5f, Width - 1, Height - 2.5f), 7);
            DrawMini(g, MinusRect, false, pressDir == -1);
            DrawMini(g, PlusRect, true, pressDir == 1);
        }

        void DrawMini(Graphics g, Rectangle r, bool plus, bool pressed)
        {
            Sk.Raised(g, new RectangleF(r.X, r.Y, r.Width, r.Height - 1), 5, false, pressed, false);
            float cx = r.X + r.Width / 2f, cy = r.Y + r.Height / 2f - 0.5f + (pressed ? 1 : 0);
            Color c = locked ? Sk.TextDim : Sk.Text;
            using (var sh = new Pen(Color.Black, 2f))
            using (var pen = new Pen(c, 2f))
            {
                g.DrawLine(sh, cx - 4.5f, cy + 1, cx + 4.5f, cy + 1);
                g.DrawLine(pen, cx - 4.5f, cy, cx + 4.5f, cy);
                if (plus) { g.DrawLine(sh, cx, cy - 3.5f, cx, cy + 5.5f); g.DrawLine(pen, cx, cy - 4.5f, cx, cy + 4.5f); }
            }
        }
    }

    class SkProgress : SkControl
    {
        int max = 100, val;
        public int Maximum { get { return max; } set { max = Math.Max(1, value); Invalidate(); } }
        public int Value { get { return val; } set { val = Math.Max(0, Math.Min(max, value)); Invalidate(); } }
        public SkProgress() { Cursor = Cursors.Default; }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = Smooth(e);
            var r = new RectangleF(0.5f, 0.5f, Width - 1, Height - 2);
            Sk.Well(g, r, Height / 2f);
            float w = (Width - 4) * (float)val / max;
            if (w < 2) return;
            var fr = new RectangleF(2, 2, Math.Max(w, Height - 4), Height - 6);
            using (var p = Sk.Round(fr, fr.Height / 2))
            {
                using (var b = new System.Drawing.Drawing2D.LinearGradientBrush(new RectangleF(fr.X, fr.Y - 1, fr.Width, fr.Height + 2), Color.FromArgb(225, 226, 232), Color.FromArgb(120, 121, 128), 90f))
                    g.FillPath(b, p);
                using (var pen = new Pen(Color.FromArgb(10, 10, 12))) g.DrawPath(pen, p);
            }
            using (var p = Sk.Round(new RectangleF(fr.X + 2, fr.Y + 1, fr.Width - 4, fr.Height / 2 - 1), fr.Height / 4))
            using (var b = new SolidBrush(Color.FromArgb(90, 255, 255, 255))) g.FillPath(b, p);
        }
    }

    static class Printing
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        static extern int StrCmpLogicalW(string a, string b);

        public static List<string> ListPdfs(string dir)
        {
            var l = Directory.GetFiles(dir, "*.pdf").ToList();
            l.Sort((a, b) => StrCmpLogicalW(Path.GetFileName(a), Path.GetFileName(b)));
            return l;
        }

        public static void MoveUnique(string file, string dir)
        {
            Directory.CreateDirectory(dir);
            string name = Path.GetFileNameWithoutExtension(file), ext = Path.GetExtension(file);
            string dest = Path.Combine(dir, name + ext);
            for (int i = 2; File.Exists(dest); i++) dest = Path.Combine(dir, name + " (" + i + ")" + ext);
            File.Move(file, dest);
        }

        public static int QueueCount(string printer) { string e; return QueueCount(printer, out e); }

        public static int QueueCount(string printer, out string error)
        {
            error = null;
            try
            {
                int n = 0;
                using (var s = new ManagementObjectSearcher("SELECT Name, JobStatus FROM Win32_PrintJob"))
                    foreach (ManagementObject o in s.Get())
                    {
                        string nm = (string)o["Name"] ?? "";
                        int i = nm.LastIndexOf(',');
                        if (i <= 0 || !string.Equals(nm.Substring(0, i), printer, StringComparison.OrdinalIgnoreCase)) continue;
                        n++;
                        string js = o["JobStatus"] as string;
                        if (!string.IsNullOrEmpty(js) && Regex.IsMatch(js, "Error|Offline|Paper|Paused|Blocked|User Intervention", RegexOptions.IgnoreCase))
                            error = js;
                    }
                return n;
            }
            catch { return 0; }
        }

        static string Wql(string s) { return s.Replace("\\", "\\\\").Replace("'", "\\'"); }

        public static string DetectIp(string printer)
        {
            try
            {
                string port = null;
                using (var s = new ManagementObjectSearcher("SELECT PortName FROM Win32_Printer WHERE Name='" + Wql(printer) + "'"))
                    foreach (ManagementObject o in s.Get()) port = o["PortName"] as string;
                if (string.IsNullOrEmpty(port)) return "";
                using (var s = new ManagementObjectSearcher("SELECT HostAddress FROM Win32_TCPIPPrinterPort WHERE Name='" + Wql(port) + "'"))
                    foreach (ManagementObject o in s.Get())
                    {
                        string h = o["HostAddress"] as string;
                        if (!string.IsNullOrEmpty(h)) return h;
                    }
                var m = Regex.Match(port, @"\b\d{1,3}(\.\d{1,3}){3}\b");
                if (m.Success) return m.Value;
            }
            catch { }
            return "";
        }

        // Рендер сторінок PDF вбудованим рушієм Windows і друк окремим завданням
        public static int PrintPdf(string path, string printer, Duplex duplex, bool color, int dpi, string testDir)
        {
            var pages = Render(path, dpi);
            try
            {
                using (var pd = new PrintDocument())
                {
                    pd.PrinterSettings.PrinterName = printer;
                    if (!pd.PrinterSettings.IsValid) throw new Exception("принтер «" + printer + "» недоступний");
                    pd.DocumentName = Path.GetFileName(path);
                    pd.PrintController = new StandardPrintController();
                    if (pd.PrinterSettings.CanDuplex) pd.PrinterSettings.Duplex = duplex;
                    pd.DefaultPageSettings.Color = color;
                    pd.OriginAtMargins = false;
                    if (printer.IndexOf("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // для перевірки без паперу: результат пишеться у файл
                        Directory.CreateDirectory(testDir);
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = Path.Combine(testDir, Path.GetFileName(path));
                    }

                    int idx = 0;
                    pd.QueryPageSettings += (s, e) =>
                    {
                        if (idx < pages.Count) e.PageSettings.Landscape = pages[idx].Width > pages[idx].Height;
                    };
                    pd.PrintPage += (s, e) =>
                    {
                        Bitmap img = pages[idx];
                        // сторінка у фізичних розмірах (1/100 дюйма), без масштабування «під поля»
                        float pw = e.PageBounds.Width, ph = e.PageBounds.Height;
                        float iw = img.Width * 100f / dpi, ih = img.Height * 100f / dpi;
                        float k = Math.Min(1f, Math.Min(pw / iw, ph / ih));
                        iw *= k; ih *= k;
                        float x = (pw - iw) / 2f - e.PageSettings.HardMarginX;
                        float y = (ph - ih) / 2f - e.PageSettings.HardMarginY;
                        var dest = new RectangleF(x, y, iw, ih);
                        if (color)
                            e.Graphics.DrawImage(img, dest);
                        else
                            using (var ia = new ImageAttributes())
                            {
                                ia.SetColorMatrix(new ColorMatrix(new[] {
                                    new[] { .299f, .299f, .299f, 0, 0 }, new[] { .587f, .587f, .587f, 0, 0 },
                                    new[] { .114f, .114f, .114f, 0, 0 }, new[] { 0f, 0, 0, 1, 0 }, new[] { 0f, 0, 0, 0, 1 } }));
                                e.Graphics.DrawImage(img, Rectangle.Round(dest), 0, 0, img.Width, img.Height, GraphicsUnit.Pixel, ia);
                            }
                        idx++;
                        e.HasMorePages = idx < pages.Count;
                    };
                    pd.Print();
                }
                return pages.Count;
            }
            finally
            {
                foreach (var b in pages) b.Dispose();
            }
        }

        static List<Bitmap> Render(string path, int dpi)
        {
            var file = StorageFile.GetFileFromPathAsync(path).AsTask().Result;
            var doc = PdfDocument.LoadFromFileAsync(file).AsTask().Result;
            if (doc.PageCount == 0) throw new Exception("PDF без сторінок");
            var list = new List<Bitmap>();
            try
            {
                for (uint i = 0; i < doc.PageCount; i++)
                {
                    using (var page = doc.GetPage(i))
                    using (var ms = new InMemoryRandomAccessStream())
                    {
                        var opt = new PdfPageRenderOptions();
                        // розмір сторінки — у DIP (1/96 дюйма)
                        opt.DestinationWidth = (uint)Math.Round(page.Size.Width / 96.0 * dpi);
                        opt.DestinationHeight = (uint)Math.Round(page.Size.Height / 96.0 * dpi);
                        page.RenderToStreamAsync(ms, opt).AsTask().Wait();
                        using (var st = ms.AsStream())
                        using (var tmp = new Bitmap(st))
                        {
                            var bmp = new Bitmap(tmp.Width, tmp.Height, PixelFormat.Format24bppRgb);
                            bmp.SetResolution(dpi, dpi);
                            using (var g = Graphics.FromImage(bmp))
                            {
                                g.Clear(Color.White); // прозорий фон PDF -> білий
                                g.DrawImage(tmp, 0, 0, tmp.Width, tmp.Height);
                            }
                            list.Add(bmp);
                        }
                    }
                }
            }
            catch
            {
                foreach (var b in list) b.Dispose();
                throw;
            }
            return list;
        }
    }

    // Мінімальний SNMP v1 GET — щоб дізнатися стан принтера і лічильник сторінок
    static class Snmp
    {
        public const string HrPrinterStatus = "1.3.6.1.2.1.25.3.5.1.1.1"; // 3 idle, 4 printing, 5 warmup
        public const string HrDeviceStatus = "1.3.6.1.2.1.25.3.2.1.5.1";  // 2 running, 3 warning, 5 down
        public const string PageCounter = "1.3.6.1.2.1.43.10.2.1.4.1.1";  // prtMarkerLifeCount
        static int reqId = Environment.TickCount & 0x7FFF;

        public static string StatusText(long? st)
        {
            switch (st.HasValue ? (int)st.Value : -1)
            {
                case 3: return "готовий";
                case 4: return "друкує";
                case 5: return "прогрівається";
                case 1: return "інше";
                case 2: return "невідомо";
                default: return "немає даних";
            }
        }

        public static long? Get(string host, string community, string oid)
        {
            try
            {
                int id = Interlocked.Increment(ref reqId);
                byte[] vb = Tlv(0x30, Cat(Oid(oid), new byte[] { 0x05, 0x00 }));
                byte[] pdu = Tlv(0xA0, Cat(Int(id), Int(0), Int(0), Tlv(0x30, vb)));
                byte[] msg = Tlv(0x30, Cat(Int(0), Tlv(0x04, Encoding.ASCII.GetBytes(community)), pdu));
                using (var udp = new UdpClient())
                {
                    udp.Client.ReceiveTimeout = 1500;
                    udp.Connect(host, 161);
                    udp.Send(msg, msg.Length);
                    IPEndPoint ep = null;
                    return Parse(udp.Receive(ref ep));
                }
            }
            catch { return null; }
        }

        static byte[] Cat(params byte[][] parts) { return parts.SelectMany(p => p).ToArray(); }

        static byte[] Tlv(byte tag, byte[] content)
        {
            int n = content.Length;
            byte[] len = n < 128 ? new[] { (byte)n } : n < 256 ? new byte[] { 0x81, (byte)n } : new byte[] { 0x82, (byte)(n >> 8), (byte)n };
            return Cat(new[] { tag }, len, content);
        }

        static byte[] Int(long v)
        {
            var b = new List<byte>();
            while (true)
            {
                b.Insert(0, (byte)(v & 0xFF));
                v >>= 8;
                if ((v == 0 && (b[0] & 0x80) == 0) || (v == -1 && (b[0] & 0x80) != 0)) break;
            }
            return Tlv(0x02, b.ToArray());
        }

        static byte[] Oid(string oid)
        {
            uint[] p = oid.Split('.').Select(uint.Parse).ToArray();
            var b = new List<byte> { (byte)(p[0] * 40 + p[1]) };
            for (int i = 2; i < p.Length; i++)
            {
                uint v = p[i];
                var t = new Stack<byte>();
                t.Push((byte)(v & 0x7F));
                v >>= 7;
                while (v > 0) { t.Push((byte)((v & 0x7F) | 0x80)); v >>= 7; }
                b.AddRange(t);
            }
            return Tlv(0x06, b.ToArray());
        }

        // повертає позицію початку вмісту; pos переводиться за кінець елемента
        static int Read(byte[] d, ref int pos, out byte tag, out int len)
        {
            tag = d[pos++];
            int l = d[pos++];
            if ((l & 0x80) != 0)
            {
                int nb = l & 0x7F; l = 0;
                for (int i = 0; i < nb; i++) l = (l << 8) | d[pos++];
            }
            len = l;
            int start = pos;
            pos += l;
            return start;
        }

        static long? Parse(byte[] d)
        {
            byte tag; int len; int pos = 0;
            pos = Read(d, ref pos, out tag, out len);        // message
            Read(d, ref pos, out tag, out len);              // version
            Read(d, ref pos, out tag, out len);              // community
            pos = Read(d, ref pos, out tag, out len);        // response PDU
            if (tag != 0xA2) return null;
            Read(d, ref pos, out tag, out len);              // request id
            int es = Read(d, ref pos, out tag, out len);     // error status
            if (len > 0 && d[es] != 0) return null;
            Read(d, ref pos, out tag, out len);              // error index
            pos = Read(d, ref pos, out tag, out len);        // varbind list
            pos = Read(d, ref pos, out tag, out len);        // varbind
            Read(d, ref pos, out tag, out len);              // oid
            int vs = Read(d, ref pos, out tag, out len);     // value
            if (len == 0 || len > 8) return null;
            if (tag == 0x02)
            {
                long v = (d[vs] & 0x80) != 0 ? -1 : 0;
                for (int i = 0; i < len; i++) v = (v << 8) | d[vs + i];
                return v;
            }
            if (tag == 0x41 || tag == 0x42 || tag == 0x43 || tag == 0x46)
            {
                long v = 0;
                for (int i = 0; i < len; i++) v = (v << 8) | d[vs + i];
                return v;
            }
            return null;
        }
    }
}
