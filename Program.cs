using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.MediaFoundation;
using NAudio.Wave;

namespace Mirror2Out
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            MediaFoundationApi.Startup(); // 为重采样做准备（仅在需要时用）
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp())
            {
                Application.Run();
            }
            MediaFoundationApi.Shutdown();
        }
    }

    /// <summary>
    /// 托盘常驻：SPDIF 主路（独占优先、低延迟） + HDMI 从路（共享、高缓冲）
    /// 输入：若选“录音设备”（如 VB-CABLE Output）→ WasapiCapture；否则“渲染环回”→ WasapiLoopbackCapture
    /// </summary>
    class TrayApp : IDisposable
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();

        MMDevice _inDev;              // 可能是 Capture 或 Render
        MMDevice _outSpdif, _outHdmi; // 两个渲染设备
        IWaveIn _capture;             // WasapiCapture 或 WasapiLoopbackCapture（均实现 IWaveIn）
        BufferedWaveProvider _bufSpdif, _bufHdmi;
        IWaveProvider _srcSpdif, _srcHdmi; // 最终喂给输出的源（可能包了重采样）
        WasapiOut _spdifOut, _hdmiOut;
        MediaFoundationResampler _resSpdif; // SPDIF 主路重采样器（需要时）

        int _bufMsSpdif = 8;    // SPDIF 小缓冲（低延迟）
        int _bufMsHdmi  = 120;  // HDMI 大缓冲（稳定推流）
        bool _running = false;

        readonly string _log = Path.Combine(Path.GetTempPath(), "Mirror2Out.log");
        void Log(string msg)
        {
            try { File.AppendAllText(_log, $"[{DateTime.Now:HH:mm:ss}] {msg}\r\n"); } catch { }
        }

        public TrayApp()
        {
            _tray.Icon = SystemIcons.Application;
            _tray.Visible = true;
            _tray.Text = "Mirror2Out（SPDIF低延迟 / HDMI高缓冲）";

            var startItem = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var stopItem  = new ToolStripMenuItem("停止(&T)", null, (s, e) => Stop());
            var devItem   = new ToolStripMenuItem("选择设备(&D)");
            var exitItem  = new ToolStripMenuItem("退出(&X)", null, (s, e) => { Stop(); Application.Exit(); });

            _menu.Opening += (s, e) => BuildDeviceMenus(devItem);

            _menu.Items.Add(startItem);
            _menu.Items.Add(stopItem);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(devItem);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(new ToolStripLabel("建议：VB-CABLE 设为 24-bit / 192 kHz"));
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(new ToolStripMenuItem("打开日志目录", null, (s,e)=> Process.Start("explorer.exe", Path.GetTempPath())));
            _menu.Items.Add(exitItem);
            _tray.ContextMenuStrip = _menu;

            Info("右键托盘图标设置输入/输出设备；建议系统默认输出设为 VB-CABLE，并在其属性中设 24-bit/192 kHz。");
            StartOrRestart();
        }

        void BuildDeviceMenus(ToolStripMenuItem root)
        {
            root.DropDownItems.Clear();
            var mm = new MMDeviceEnumerator();

            // —— 输入 —— //
            var inMenu = new ToolStripMenuItem("输入设备");
            inMenu.DropDownItems.Add("使用默认渲染设备（环回）", null, (s, e) => { _inDev = null; Info("输入→默认渲染环回"); });
            var vb = mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                       .FirstOrDefault(d => d.FriendlyName.IndexOf("CABLE Output", StringComparison.OrdinalIgnoreCase) >= 0);
            if (vb != null)
            {
                inMenu.DropDownItems.Add("使用 VB-CABLE Output（录音）", null, (s, e) => { _inDev = vb; Info("输入→VB-CABLE Output（录音捕获）"); });
            }

            inMenu.DropDownItems.Add(new ToolStripSeparator());
            inMenu.DropDownItems.Add("（选择任一录音设备）").Enabled = false;
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
            {
                var item = new ToolStripMenuItem("录音: " + d.FriendlyName);
                item.Click += (s, e) => { _inDev = d; Info("输入→" + d.FriendlyName + "（录音捕获）"); };
                inMenu.DropDownItems.Add(item);
            }
            inMenu.DropDownItems.Add(new ToolStripSeparator());
            inMenu.DropDownItems.Add("（选择任一渲染设备环回）").Enabled = false;
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                var item = new ToolStripMenuItem("环回: " + d.FriendlyName);
                item.Click += (s, e) => { _inDev = d; Info("输入→" + d.FriendlyName + "（渲染环回）"); };
                inMenu.DropDownItems.Add(item);
            }

            // —— SPDIF 输出 —— //
            var spdifMenu = new ToolStripMenuItem("SPDIF 输出（主路，低延迟）");
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                                .OrderByDescending(d => ScoreSpdif(d.FriendlyName)))
            {
                var txt = d.FriendlyName + ( (_outSpdif != null && d.ID == _outSpdif.ID) ? "  ✓" : "");
                var item = new ToolStripMenuItem(txt);
                item.Click += (s, e) => { _outSpdif = d; Info("SPDIF 输出→" + d.FriendlyName); };
                spdifMenu.DropDownItems.Add(item);
            }

            // —— HDMI 输出 —— //
            var hdmiMenu = new ToolStripMenuItem("HDMI 输出（从路，高缓冲）");
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                                .OrderByDescending(d => ScoreHdmi(d.FriendlyName)))
            {
                var txt = d.FriendlyName + ( (_outHdmi != null && d.ID == _outHdmi.ID) ? "  ✓" : "");
                var item = new ToolStripMenuItem(txt);
                item.Click += (s, e) => { _outHdmi = d; Info("HDMI 输出→" + d.FriendlyName); };
                hdmiMenu.DropDownItems.Add(item);
            }

            root.DropDownItems.Add(inMenu);
            root.DropDownItems.Add(spdifMenu);
            root.DropDownItems.Add(hdmiMenu);
        }

        static int ScoreSpdif(string name)
        {
            name = name.ToLowerInvariant();
            int s = 0;
            if (name.Contains("spdif") || name.Contains("s/pdif")) s += 10;
            if (name.Contains("digital")) s += 5;
            if (name.Contains("optical") || name.Contains("coax")) s += 3;
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

            // 选择输入（优先 VB-CABLE Output 录音；否则默认渲染环回）
            if (_inDev == null)
            {
                var vb = mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                           .FirstOrDefault(d => d.FriendlyName.IndexOf("CABLE Output", StringComparison.OrdinalIgnoreCase) >= 0);
                if (vb != null) _inDev = vb;
                if (_inDev == null)
                {
                    _inDev = mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                }
            }

            // 选择输出（如未设定则猜测一个）
            if (_outSpdif == null)
            {
                _outSpdif = mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                              .OrderByDescending(d => ScoreSpdif(d.FriendlyName)).FirstOrDefault();
            }
            if (_outHdmi == null)
            {
                _outHdmi = mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                             .OrderByDescending(d => ScoreHdmi(d.FriendlyName)).FirstOrDefault();
            }

            if (_outSpdif == null || _outHdmi == null)
            {
                Info("请在菜单里设置 SPDIF 与 HDMI 输出设备。");
                return;
            }

            // —— 构建输入捕获，并获得 inFmt —— //
            WaveFormat inFmt;
            if (_inDev.DataFlow == DataFlow.Capture)
            {
                var cap = new WasapiCapture(_inDev) { ShareMode = AudioClientShareMode.Shared };
                _capture = cap;
                inFmt = cap.WaveFormat;
            }
            else
            {
                var cap = new WasapiLoopbackCapture(_inDev);
                _capture = cap;
                inFmt = cap.WaveFormat;
            }

            Log($"Input format: {inFmt.SampleRate} Hz, {inFmt.BitsPerSample} bit, ch={inFmt.Channels}, from={_inDev.FriendlyName}");

            _bufSpdif = new BufferedWaveProvider(inFmt)
            {
                DiscardOnBufferOverflow = true,
                BufferDuration = TimeSpan.FromMilliseconds(_bufMsSpdif * 8)
            };
            _bufHdmi = new BufferedWaveProvider(inFmt)
            {
                DiscardOnBufferOverflow = true,
                BufferDuration = TimeSpan.FromMilliseconds(_bufMsHdmi * 8)
            };

            // —— SPDIF：独占格式协商 + 必要时重采样 —— //
            WaveFormat spdifFmtExclusive = PickSpdifExclusiveFormat(_outSpdif, true);
            bool useExclusive = (spdifFmtExclusive != null);
            if (useExclusive)
            {
                Log($"SPDIF exclusive target: {spdifFmtExclusive.SampleRate} Hz, {spdifFmtExclusive.BitsPerSample} bit, ch={spdifFmtExclusive.Channels}");
                _srcSpdif = _bufSpdif;
                if (!WaveFormatsEqual(inFmt, spdifFmtExclusive))
                {
                    // 仅主路重采样到独占格式
                    _resSpdif = new MediaFoundationResampler(_bufSpdif, spdifFmtExclusive);
                    _resSpdif.ResamplerQuality = 60; // 0-60
                    _srcSpdif = _resSpdif;
                }
                _spdifOut = TryCreateWasapiOut(_outSpdif, AudioClientShareMode.Exclusive, _bufMsSpdif, _srcSpdif);
                if (_spdifOut == null)
                {
                    Log("Exclusive init failed, fallback to shared.");
                    useExclusive = false;
                }
            }
            if (!useExclusive)
            {
                _srcSpdif = _bufSpdif; // 交给系统混音
                _spdifOut = TryCreateWasapiOut(_outSpdif, AudioClientShareMode.Shared, Math.Max(_bufMsSpdif, 10), _srcSpdif);
                if (_spdifOut == null)
                {
                    Info("SPDIF 输出初始化失败（独占/共享均不可用）。");
                    return;
                }
            }

            // —— HDMI：共享 + 大缓冲 —— //
            _srcHdmi = _bufHdmi;
            _hdmiOut = TryCreateWasapiOut(_outHdmi, AudioClientShareMode.Shared, _bufMsHdmi, _srcHdmi);
            if (_hdmiOut == null)
            {
                Info("HDMI 输出初始化失败。");
                _spdifOut.Dispose(); _spdifOut = null;
                if (_resSpdif != null) { _resSpdif.Dispose(); _resSpdif = null; }
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
                if (_bufSpdif != null) _bufSpdif.ClearBuffer();
                if (_bufHdmi  != null) _bufHdmi.ClearBuffer();
            };

            try
            {
                _spdifOut.Play();
                _hdmiOut.Play();
                _capture.StartRecording();
                _running = true;

                var inName = _inDev.FriendlyName;
                var sName  = _outSpdif.FriendlyName + (useExclusive ? " [独占]" : " [共享]");
                var hName  = _outHdmi.FriendlyName + " [共享]";
                Info($"运行中：\n输入《{inName}》→ SPDIF：{sName}（{_bufMsSpdif}ms） + HDMI：{hName}（{_bufMsHdmi}ms）\n" +
                     $"若需 24/192，请在“声音”中把 VB-CABLE 与 SPDIF 都设为 24-bit / 192 kHz。");
                Log($"Started. SPDIF {(useExclusive ? "Exclusive" : "Shared")} buf={_bufMsSpdif}ms, HDMI Shared buf={_bufMsHdmi}ms");
            }
            catch (Exception ex)
            {
                Log("Start failed: " + ex);
                Info("启动失败：" + ex.Message);
                Stop();
            }
        }

        // 选 SPDIF 独占格式：优先 192/24，然后 96/24、48/24、96/16、48/16（双声道）
        WaveFormat PickSpdifExclusiveFormat(MMDevice dev, bool prefer192)
        {
            try
            {
                var ac = dev.AudioClient;
                int[] rates = prefer192 ? new[] { 192000, 176400, 96000, 88200, 48000, 44100 }
                                        : new[] { 96000, 88200, 48000, 44100 };
                int[] bits = new[] { 24, 16 };
                foreach (var sr in rates)
                {
                    foreach (var b in bits)
                    {
                        var fmt = new WaveFormat(sr, b, 2);
                        if (ac.IsFormatSupported(AudioClientShareMode.Exclusive, fmt))
                            return fmt;
                    }
                }
            }
            catch { }
            return null;
        }

        static bool WaveFormatsEqual(WaveFormat a, WaveFormat b)
        {
            if (a == null || b == null) return false;
            return a.SampleRate == b.SampleRate && a.BitsPerSample == b.BitsPerSample && a.Channels == b.Channels;
        }

        WasapiOut TryCreateWasapiOut(MMDevice dev, AudioClientShareMode mode, int bufMs, IWaveProvider src)
        {
            try
            {
                var w = new WasapiOut(dev, mode, true, bufMs);
                w.Init(src);
                return w;
            }
            catch (Exception ex)
            {
                Log($"WasapiOut init failed: {dev.FriendlyName}, mode={mode}, buf={bufMs}ms, err={ex.Message}");
                return null;
            }
        }

        public void Stop()
        {
            if (!_running)
            {
                // 也清一遍资源，防止半初始化残留
                if (_capture != null) { _capture.Dispose(); _capture = null; }
                if (_spdifOut != null) { _spdifOut.Dispose(); _spdifOut = null; }
                if (_hdmiOut  != null) { _hdmiOut.Dispose();  _hdmiOut  = null; }
                if (_resSpdif != null) { _resSpdif.Dispose(); _resSpdif = null; }
                _bufSpdif = null; _bufHdmi = null;
                return;
            }
            try { if (_capture != null) _capture.StopRecording(); } catch { }
            try { if (_spdifOut != null) _spdifOut.Stop(); } catch { }
            try { if (_hdmiOut  != null) _hdmiOut.Stop();  } catch { }
            Thread.Sleep(50);

            if (_capture != null) { _capture.Dispose(); _capture = null; }
            if (_spdifOut != null) { _spdifOut.Dispose(); _spdifOut = null; }
            if (_hdmiOut  != null) { _hdmiOut.Dispose();  _hdmiOut  = null; }
            if (_resSpdif != null) { _resSpdif.Dispose(); _resSpdif = null; }
            _bufSpdif = null; _bufHdmi = null;

            _running = false;
            _tray.ShowBalloonTip(1000, "Mirror2Out", "已停止", ToolTipIcon.Info);
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
