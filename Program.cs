using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace Mirror2Out
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp())
            {
                Application.Run();
            }
        }
    }

    /// <summary>
    /// 托盘常驻：SPDIF 低延迟（独占优先），HDMI 高缓冲（共享）
    /// 输入：VB-CABLE 等“录音设备”→ WasapiCapture；否则“渲染设备环回”→ WasapiLoopbackCapture
    /// </summary>
    class TrayApp : IDisposable
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();

        // 设备与音频对象
        MMDevice _inDev;              // 允许为 Capture 或 Render
        MMDevice _outSpdif, _outHdmi; // 两个渲染设备
        IWaveIn _capture;             // WasapiCapture 或 WasapiLoopbackCapture（都实现 IWaveIn）
        BufferedWaveProvider _bufSpdif, _bufHdmi;
        WasapiOut _spdifOut, _hdmiOut;

        // 参数（可根据机器调小/调大）
        int _bufMsSpdif = 8;     // SPDIF 主路：小缓冲（低延迟）
        int _bufMsHdmi  = 120;   // HDMI 从路：大缓冲（稳定推流）
        bool _running = false;

        public TrayApp()
        {
            _tray.Icon = SystemIcons.Application;
            _tray.Visible = true;
            _tray.Text = "Mirror2Out（SPDIF低延迟 / HDMI高缓冲）";

            // —— 菜单 —— //
            var startItem = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var stopItem  = new ToolStripMenuItem("停止(&T)", null, (s, e) => Stop());
            var devItem   = new ToolStripMenuItem("选择设备(&D)");
            var exitItem  = new ToolStripMenuItem("退出(&X)", null, (s, e) => { Stop(); Application.Exit(); });

            // 设备子菜单（在打开菜单时动态重建）
            _menu.Opening += (s, e) => BuildDeviceMenus(devItem);

            _menu.Items.Add(startItem);
            _menu.Items.Add(stopItem);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(devItem);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(new ToolStripLabel("提示：建议 VB-CABLE 设为 24-bit / 192 kHz"));
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(exitItem);
            _tray.ContextMenuStrip = _menu;

            Info("右键托盘图标设置输入/输出设备；建议系统默认输出先指到 VB-CABLE，并把其格式设为 24-bit / 192 kHz。");
            StartOrRestart();
        }

        // 动态构建设备菜单
        void BuildDeviceMenus(ToolStripMenuItem root)
        {
            root.DropDownItems.Clear();

            // 输入（录音/渲染）
            var inMenu = new ToolStripMenuItem("输入设备");
            var mm = new MMDeviceEnumerator();

            // 常用快捷项
            inMenu.DropDownItems.Add("使用默认渲染设备（环回）", null, (s, e) => { _inDev = null; Info("输入→默认渲染设备（Loopback）"); });
            var vb = mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                       .FirstOrDefault(d => d.FriendlyName.IndexOf("CABLE Output", StringComparison.OrdinalIgnoreCase) >= 0);
            if (vb != null)
                inMenu.DropDownItems.Add("使用 VB-CABLE Output（录音）", null, (s, e) => { _inDev = vb; Info("输入→VB-CABLE Output（录音捕获）"); });

            inMenu.DropDownItems.Add(new ToolStripSeparator());
            inMenu.DropDownItems.Add("（展开选择任一录音设备）").Enabled = false;
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
            {
                var item = new ToolStripMenuItem("录音: " + d.FriendlyName);
                item.Click += (s, e) => { _inDev = d; Info("输入→" + d.FriendlyName + "（录音捕获）"); };
                inMenu.DropDownItems.Add(item);
            }
            inMenu.DropDownItems.Add(new ToolStripSeparator());
            inMenu.DropDownItems.Add("（展开选择任一渲染设备环回）").Enabled = false;
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                var item = new ToolStripMenuItem("环回: " + d.FriendlyName);
                item.Click += (s, e) => { _inDev = d; Info("输入→" + d.FriendlyName + "（渲染环回）"); };
                inMenu.DropDownItems.Add(item);
            }

            // SPDIF 输出
            var spdifMenu = new ToolStripMenuItem("SPDIF 输出（主路，低延迟）");
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                                .OrderByDescending(d => ScoreSpdif(d.FriendlyName)))
            {
                var item = new ToolStripMenuItem(d.FriendlyName + (d.ID == _outSpdif?.ID ? "  ✓" : ""));
                item.Click += (s, e) => { _outSpdif = d; Info("SPDIF 输出→" + d.FriendlyName); };
                spdifMenu.DropDownItems.Add(item);
            }

            // HDMI 输出
            var hdmiMenu = new ToolStripMenuItem("HDMI 输出（从路，高缓冲）");
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                                .OrderByDescending(d => ScoreHdmi(d.FriendlyName)))
            {
                var item = new ToolStripMenuItem(d.FriendlyName + (d.ID == _outHdmi?.ID ? "  ✓" : ""));
                item.Click += (s, e) => { _outHdmi = d; Info("HDMI 输出→" + d.FriendlyName); };
                hdmiMenu.DropDownItems.Add(item);
            }

            root.DropDownItems.Add(inMenu);
            root.DropDownItems.Add(spdifMenu);
            root.DropDownItems.Add(hdmiMenu);
        }

        // 简单打分：优先把含 SPDIF/数字/光纤的放前面
        static int ScoreSpdif(string name)
        {
            name = name.ToLowerInvariant();
            int s = 0;
            if (name.Contains("spdif") || name.Contains("s/pdif")) s += 10;
            if (name.Contains("digital")) s += 5;
            if (name.Contains("optical")) s += 3;
            if (name.Contains("coax")) s += 3;
            return s;
        }
        static int ScoreHdmi(string name)
        {
            name = name.ToLowerInvariant();
            int s = 0;
            if (name.Contains("hdmi")) s += 10;
            if (name.Contains("nvidia") || name.Contains("amd") || name.Contains("intel")) s += 2;
            return s;
        }

        void StartOrRestart()
        {
            Stop();

            var mm = new MMDeviceEnumerator();

            // 选择输入：未手动选则优先 VB-CABLE Output（录音），否则默认渲染环回
            if (_inDev == null)
            {
                _inDev = mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                           .FirstOrDefault(d => d.FriendlyName.IndexOf("CABLE Output", StringComparison.OrdinalIgnoreCase) >= 0)
                        ?? mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            }

            // 选择输出：尽力猜测，仍允许手动改
            _outSpdif ??= mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                            .OrderByDescending(d => ScoreSpdif(d.FriendlyName)).FirstOrDefault();
            _outHdmi  ??= mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                            .OrderByDescending(d => ScoreHdmi(d.FriendlyName)).FirstOrDefault();

            if (_outSpdif == null || _outHdmi == null)
            {
                Info("请先在菜单里设置 SPDIF 与 HDMI 输出设备。");
                return;
            }

            // —— 构建输入捕获 —— //
            if (_inDev.DataFlow == DataFlow.Capture)
            {
                // 录音设备捕获（适用于 VB-CABLE Output）
                var cap = new WasapiCapture(_inDev) { ShareMode = AudioClientShareMode.Shared };
                _capture = cap;
            }
            else
            {
                // 渲染设备环回（扬声器/HDMI）
                _capture = new WasapiLoopbackCapture(_inDev);
            }

            var fmt = (_capture as WasapiCaptureBase)?.WaveFormat ?? new WaveFormat(48000, 16, 2);

            _bufSpdif = new BufferedWaveProvider(fmt)
            {
                DiscardOnBufferOverflow = true,
                BufferDuration = TimeSpan.FromMilliseconds(_bufMsSpdif * 8)
            };
            _bufHdmi = new BufferedWaveProvider(fmt)
            {
                DiscardOnBufferOverflow = true,
                BufferDuration = TimeSpan.FromMilliseconds(_bufMsHdmi * 8)
            };

            // —— 输出：SPDIF 优先独占小缓冲；失败则回退共享 —— //
            _spdifOut = TryCreateWasapiOut(_outSpdif, AudioClientShareMode.Exclusive, _bufMsSpdif, _bufSpdif)
                        ?? TryCreateWasapiOut(_outSpdif, AudioClientShareMode.Shared,    Math.Max(_bufMsSpdif, 10), _bufSpdif);

            if (_spdifOut == null)
            {
                Info("SPDIF 输出初始化失败（独占与共享均未成功）。请检查格式/设备是否被占用。");
                return;
            }

            // —— HDMI：共享 + 大缓冲（直播不敏感延迟） —— //
            _hdmiOut = TryCreateWasapiOut(_outHdmi, AudioClientShareMode.Shared, _bufMsHdmi, _bufHdmi);
            if (_hdmiOut == null)
            {
                Info("HDMI 输出初始化失败。");
                _spdifOut.Dispose(); _spdifOut = null;
                return;
            }

            // 绑定数据
            _capture.DataAvailable += (s, e) =>
            {
                _bufSpdif.AddSamples(e.Buffer, 0, e.BytesRecorded);
                _bufHdmi.AddSamples(e.Buffer, 0, e.BytesRecorded);
            };
            _capture.RecordingStopped += (s, e) =>
            {
                _bufSpdif?.ClearBuffer();
                _bufHdmi?.ClearBuffer();
            };

            try
            {
                _spdifOut.Play();
                _hdmiOut.Play();
                _capture.StartRecording();
                _running = true;

                var inName = _inDev.FriendlyName;
                var sName  = _outSpdif.FriendlyName;
                var hName  = _outHdmi.FriendlyName;
                Info($"运行中：\n输入《{inName}》→ SPDIF(优先独占，{_bufMsSpdif}ms) + HDMI(共享，{_bufMsHdmi}ms)\n" +
                     $"若要 24/192，请在“声音”中把 VB-CABLE 与 SPDIF 都设为 24-bit / 192 kHz。");
            }
            catch (Exception ex)
            {
                Info("启动失败：" + ex.Message);
                Stop();
            }
        }

        // 创建 WasapiOut（带异常保护）
        WasapiOut TryCreateWasapiOut(MMDevice dev, AudioClientShareMode mode, int bufMs, IWaveProvider src)
        {
            try
            {
                var w = new WasapiOut(dev, mode, true, bufMs);
                w.Init(src);
                return w;
            }
            catch
            {
                try { src = src; } catch { }
                return null;
            }
        }

        public void Stop()
        {
            if (!_running) return;
            try { _capture?.StopRecording(); } catch { }
            try { _spdifOut?.Stop(); } catch { }
            try { _hdmiOut?.Stop(); } catch { }
            Thread.Sleep(50);

            _capture?.Dispose(); _capture = null;
            _spdifOut?.Dispose(); _spdifOut = null;
            _hdmiOut?.Dispose();  _hdmiOut  = null;
            _bufSpdif = null; _bufHdmi = null;

            _running = false;
            _tray.ShowBalloonTip(1200, "Mirror2Out", "已停止", ToolTipIcon.Info);
        }

        public void Dispose()
        {
            Stop();
            _tray.Visible = false;
            _tray.Dispose();
            _menu.Dispose();
        }

        static void Info(string t)
            => MessageBox.Show(t, "Mirror2Out", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
