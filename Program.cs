using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
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
            try { MediaFoundationApi.Startup(); } catch { } // Windows N 可忽略
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new TrayApp()) { Application.Run(); }
            try { MediaFoundationApi.Shutdown(); } catch { }
        }
    }

    [DataContract]
    class AppSettings
    {
        [DataMember] public string InputDeviceId;
        [DataMember] public string MainDeviceId;
        [DataMember] public string AuxDeviceId;

        [DataMember] public bool   MainExclusive = true;  // 主通道优先独占
        [DataMember] public int    MainRate      = 192000;
        [DataMember] public int    MainBits      = 24;
        [DataMember] public int    MainBufMs     = 8;     // 低延迟
        [DataMember] public int    AuxBufMs      = 120;   // 高缓冲更稳更省
    }

    class Config
    {
        static string Dir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Mirror2Out");
        static string FilePath = Path.Combine(Dir, "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (!System.IO.File.Exists(FilePath)) return new AppSettings();
                using (var fs = System.IO.File.OpenRead(FilePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(AppSettings));
                    return (AppSettings)ser.ReadObject(fs);
                }
            }
            catch { return new AppSettings(); }
        }
        public static void Save(AppSettings s)
        {
            try
            {
                if (!Directory.Exists(Dir)) Directory.CreateDirectory(Dir);
                using (var fs = System.IO.File.Create(FilePath))
                {
                    var ser = new DataContractJsonSerializer(typeof(AppSettings));
                    ser.WriteObject(fs, s);
                }
            }
            catch { }
        }
    }

    class TrayApp : IDisposable
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly string _log = Path.Combine(Path.GetTempPath(), "Mirror2Out.log");

        AppSettings _cfg = Config.Load();

        // 设备与音频对象
        MMDevice _inDev;              // 数据流输入（可为 Capture 或 Render）
        MMDevice _outMain;            // 主通道输出（低延迟/高音质）
        MMDevice _outAux;             // 副通道输出（普通延迟/低占用）

        IWaveIn _capture;             // WasapiCapture 或 WasapiLoopbackCapture
        BufferedWaveProvider _bufMain, _bufAux;
        IWaveProvider _srcMain, _srcAux; // 最终喂给输出（主通道可能包重采样）
        WasapiOut _mainOut, _auxOut;
        MediaFoundationResampler _resMain;

        bool _running = false;

        public TrayApp()
        {
            _tray.Icon = SystemIcons.Application;
            _tray.Visible = true;
            _tray.Text = "Mirror2Out（三通道：输入/主/副）";

            var startItem = new ToolStripMenuItem("启动/重启(&S)", null, (s,e)=> StartOrRestart());
            var stopItem  = new ToolStripMenuItem("停止(&T)", null, (s,e)=> Stop());
            var setItem   = new ToolStripMenuItem("设置(&G)...", null, (s,e)=> OpenSettings());
            var openLog   = new ToolStripMenuItem("打开日志目录", null, (s,e)=> Process.Start("explorer.exe", Path.GetTempPath()));
            var exitItem  = new ToolStripMenuItem("退出(&X)", null, (s,e)=> { Stop(); Application.Exit(); });

            _menu.Items.Add(startItem);
            _menu.Items.Add(stopItem);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(setItem);
            _menu.Items.Add(openLog);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(exitItem);
            _tray.ContextMenuStrip = _menu;

            Info("右键→设置：可手选通道1输入、通道2主、通道3副；主通道默认独占192/24，副通道共享高缓冲。");
            StartOrRestart();
        }

        void OpenSettings()
        {
            using (var f = new SettingsForm(_cfg))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    _cfg = f.Result;
                    Config.Save(_cfg);
                    Info("设置已保存，正在重启音频...");
                    StartOrRestart();
                }
            }
        }

        void Log(string msg) { try { File.AppendAllText(_log, $"[{DateTime.Now:HH:mm:ss}] {msg}\r\n"); } catch { } }
        static void Info(string t) => MessageBox.Show(t, "Mirror2Out", MessageBoxButtons.OK, MessageBoxIcon.Information);

        void StartOrRestart()
        {
            Stop();
            var mm = new MMDeviceEnumerator();

            // 绑定设备（允许任意选择；找不到则给出提示）
            _inDev   = FindById(mm, _cfg.InputDeviceId, DataFlow.Capture) ?? FindById(mm, _cfg.InputDeviceId, DataFlow.Render);
            _outMain = FindById(mm, _cfg.MainDeviceId,  DataFlow.Render);
            _outAux  = FindById(mm, _cfg.AuxDeviceId,   DataFlow.Render);

            if (_inDev == null)
                _inDev = mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia); // 兜底：默认渲染环回
            if (_outMain == null || _outAux == null)
            {
                Info("请先在“设置”里选择主/副输出设备。");
                return;
            }

            // 构建输入：录音设备→WasapiCapture；渲染设备→Loopback
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
            Log("Input: " + _inDev.FriendlyName + $" | {inFmt.SampleRate}Hz/{inFmt.BitsPerSample}bit/{inFmt.Channels}ch");

            // 缓冲
            _bufMain = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, BufferDuration = TimeSpan.FromMilliseconds(_cfg.MainBufMs * 8) };
            _bufAux  = new BufferedWaveProvider(inFmt) { DiscardOnBufferOverflow = true, BufferDuration = TimeSpan.FromMilliseconds(_cfg.AuxBufMs  * 8) };

            // 主通道：优先独占到指定格式（必要时重采样），失败则回退共享
            bool mainExclusive = false;
            _srcMain = _bufMain;
            _resMain = null;

            WaveFormat desired = new WaveFormat(_cfg.MainRate, _cfg.MainBits, 2);
            try
            {
                if (_cfg.MainExclusive && _outMain.AudioClient.IsFormatSupported(AudioClientShareMode.Exclusive, desired))
                {
                    if (!WaveFormatsEqual(inFmt, desired))
                    {
                        _resMain = new MediaFoundationResampler(_bufMain, desired);
                        _resMain.ResamplerQuality = 60;
                        _srcMain = _resMain;
                    }
                    _mainOut = TryCreateWasapiOut(_outMain, AudioClientShareMode.Exclusive, _cfg.MainBufMs, _srcMain);
                    mainExclusive = (_mainOut != null);
                }
            }
            catch { }

            if (!mainExclusive)
            {
                // 共享模式，省资源（不必手动重采样）
                _srcMain = _bufMain;
                _mainOut = TryCreateWasapiOut(_outMain, AudioClientShareMode.Shared, Math.Max(_cfg.MainBufMs, 10), _srcMain);
                if (_mainOut == null)
                {
                    Info("主通道初始化失败（独占/共享均不可用）。");
                    CleanupCreated();
                    return;
                }
            }

            // 副通道：始终共享 + 高缓冲（最省心）
            _srcAux = _bufAux;
            _auxOut = TryCreateWasapiOut(_outAux, AudioClientShareMode.Shared, _cfg.AuxBufMs, _srcAux);
            if (_auxOut == null)
            {
                Info("副通道初始化失败（共享不可用）。");
                CleanupCreated();
                return;
            }

            // 绑定数据
            _capture.DataAvailable += (s, e) =>
            {
                _bufMain.AddSamples(e.Buffer, 0, e.BytesRecorded);
                _bufAux.AddSamples(e.Buffer, 0, e.BytesRecorded);
            };
            _capture.RecordingStopped += (s, e) =>
            {
                if (_bufMain != null) _bufMain.ClearBuffer();
                if (_bufAux  != null) _bufAux.ClearBuffer();
            };

            try
            {
                _mainOut.Play();
                _auxOut.Play();
                _capture.StartRecording();
                _running = true;

                Info("运行中：\n" +
                     $"通道1 输入：{_inDev.FriendlyName}\n" +
                     $"通道2 主通道：{_outMain.FriendlyName}（{(mainExclusive ? "独占" : "共享")}，{_cfg.MainBufMs}ms）\n" +
                     $"通道3 副通道：{_outAux.FriendlyName}（共享，{_cfg.AuxBufMs}ms）");
                Log($"Started: Main={(mainExclusive?"Exclusive":"Shared")} buf={_cfg.MainBufMs}ms, Aux=Shared buf={_cfg.AuxBufMs}ms");
            }
            catch (Exception ex)
            {
                Log("Start failed: " + ex);
                Info("启动失败：" + ex.Message);
                Stop();
            }
        }

        void CleanupCreated()
        {
            try { if (_resMain != null) _resMain.Dispose(); } catch { }
            try { if (_mainOut != null) _mainOut.Dispose(); } catch { }
            try { if (_auxOut  != null) _auxOut.Dispose();  } catch { }
            _resMain = null; _mainOut = null; _auxOut = null;
        }

        public void Stop()
        {
            if (!_running)
            {
                if (_capture != null) { _capture.Dispose(); _capture = null; }
                if (_mainOut != null) { _mainOut.Dispose(); _mainOut = null; }
                if (_auxOut  != null) { _auxOut.Dispose();  _auxOut  = null; }
                if (_resMain != null) { _resMain.Dispose(); _resMain = null; }
                _bufMain = null; _bufAux = null;
                return;
            }
            try { if (_capture != null) _capture.StopRecording(); } catch { }
            try { if (_mainOut != null) _mainOut.Stop(); } catch { }
            try { if (_auxOut  != null) _auxOut.Stop();  } catch { }
            Thread.Sleep(50);

            if (_capture != null) { _capture.Dispose(); _capture = null; }
            if (_mainOut != null) { _mainOut.Dispose(); _mainOut = null; }
            if (_auxOut  != null) { _auxOut.Dispose();  _auxOut  = null; }
            if (_resMain != null) { _resMain.Dispose(); _resMain = null; }
            _bufMain = null; _bufAux = null;
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

        // —— 工具方法 —— //
        static bool WaveFormatsEqual(WaveFormat a, WaveFormat b)
        {
            if (a == null || b == null) return false;
            return a.SampleRate == b.SampleRate && a.BitsPerSample == b.BitsPerSample && a.Channels == b.Channels;
        }
        static MMDevice FindById(MMDeviceEnumerator mm, string id, DataFlow flow)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return mm.EnumerateAudioEndPoints(flow, DeviceState.Active).FirstOrDefault(d => d.ID == id);
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
                Log("WasapiOut init failed: " + dev.FriendlyName + $" | mode={mode} buf={bufMs}ms | " + ex.Message);
                return null;
            }
        }
    }
}
