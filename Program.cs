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
            using (var app = new TrayApp()) Application.Run();
        }
    }

    class TrayApp : IDisposable
    {
        readonly NotifyIcon _tray = new NotifyIcon();
        readonly ContextMenuStrip _menu = new ContextMenuStrip();

        MMDevice _inDev, _outSpdif, _outHdmi;
        WasapiLoopbackCapture _capture;
        BufferedWaveProvider _bufSpdif, _bufHdmi;
        WasapiOut _spdifOut, _hdmiOut;

        int _bufMsSpdif = 8;    // SPDIF 主路：小缓冲（低延迟）
        int _bufMsHdmi  = 120;  // HDMI 从路：大缓冲（稳定推流）
        bool _running = false;

        public TrayApp()
        {
            _tray.Icon = SystemIcons.Application;
            _tray.Visible = true;
            _tray.Text = "Mirror2Out（SPDIF低延迟 / HDMI高缓冲）";

            var startItem = new ToolStripMenuItem("启动/重启(&S)", null, (s, e) => StartOrRestart());
            var stopItem  = new ToolStripMenuItem("停止(&T)", null, (s, e) => Stop());
            var devItem   = new ToolStripMenuItem("选择设备(&D)");
            var exitItem  = new ToolStripMenuItem("退出(&X)", null, (s, e) => { Stop(); Application.Exit(); });

            devItem.DropDownItems.Add("刷新列表", null, (s,e)=> {
                var list = string.Join("\n", ListRenderDevices());
                _tray.ShowBalloonTip(2500, "可用输出设备", list, ToolTipIcon.Info);
            });
            devItem.DropDownItems.Add(new ToolStripSeparator());
            devItem.DropDownItems.Add("输入=默认渲染设备", null, (s,e)=> { _inDev = null; Info("输入改为：默认渲染设备（其混音的 Loopback）"); });
            devItem.DropDownItems.Add("输入=包含“CABLE Input”的设备", null, (s,e)=> { _inDev = FindRender("CABLE Input"); Info("输入改为：VB-CABLE（请将其设为 24/192）"); });
            devItem.DropDownItems.Add(new ToolStripSeparator());
            devItem.DropDownItems.Add("SPDIF=包含“SPDIF/S-PDIF/Digital Audio”", null, (s,e)=> { _outSpdif = FindRender("SPDIF") ?? FindRender("S/PDIF") ?? FindRender("Digital Audio"); Info("SPDIF 输出设备已设置"); });
            devItem.DropDownItems.Add("HDMI=包含“HDMI/NVIDIA High Definition”", null, (s,e)=> { _outHdmi = FindRender("HDMI") ?? FindRender("NVIDIA High Definition"); Info("HDMI 输出设备已设置"); });

            _menu.Items.Add(startItem);
            _menu.Items.Add(stopItem);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(devItem);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(new ToolStripLabel("提示：建议将 VB-CABLE 设为 24-bit / 192 kHz"));
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(exitItem);
            _tray.ContextMenuStrip = _menu;

            Info("右键托盘图标设置输入/输出设备；建议系统默认输出先指到 VB-CABLE，并把其格式设为 24-bit/192kHz。");
            StartOrRestart();
        }

        void StartOrRestart()
        {
            Stop();

            var mm = new MMDeviceEnumerator();
            var inDev = _inDev ?? mm.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            if (inDev == null) { Info("找不到输入（默认渲染设备）"); return; }

            _outSpdif ??= FindRender("SPDIF") ?? FindRender("S/PDIF") ?? FindRender("Digital Audio");
            _outHdmi  ??= FindRender("HDMI") ?? FindRender("NVIDIA High Definition");
            if (_outSpdif == null || _outHdmi == null) { Info("请在菜单里设置 SPDIF 与 HDMI 输出设备"); return; }

            _capture = new WasapiLoopbackCapture(inDev);
            var fmt = _capture.WaveFormat; // 尽量让输入设备（如 VB-CABLE）预先设为 192/24

            _bufSpdif = new BufferedWaveProvider(fmt) { DiscardOnBufferOverflow = true, BufferDuration = TimeSpan.FromMilliseconds(_bufMsSpdif * 8) };
            _bufHdmi  = new BufferedWaveProvider(fmt) { DiscardOnBufferOverflow = true, BufferDuration = TimeSpan.FromMilliseconds(_bufMsHdmi  * 8) };

            _spdifOut = new WasapiOut(_outSpdif, AudioClientShareMode.Exclusive, true, _bufMsSpdif);
            _hdmiOut  = new WasapiOut(_outHdmi,  AudioClientShareMode.Shared,    true, _bufMsHdmi);

            _spdifOut.Init(_bufSpdif);
            _hdmiOut.Init(_bufHdmi);

            _capture.DataAvailable += (s, e) =>
            {
                _bufSpdif.AddSamples(e.Buffer, 0, e.BytesRecorded);
                _bufHdmi.AddSamples(e.Buffer, 0, e.BytesRecorded);
            };
            _capture.RecordingStopped += (s, e) => { _bufSpdif?.ClearBuffer(); _bufHdmi?.ClearBuffer(); };

            _spdifOut.Play();
            _hdmiOut.Play();
            _capture.StartRecording();

            _running = true;
            Info($"运行中：输入《{inDev.FriendlyName}》→ SPDIF(独占 {_bufMsSpdif}ms) + HDMI(共享 {_bufMsHdmi}ms)。\n若要 24/192，请将输入设备（如 VB-CABLE）在系统里设为 24-bit / 192 kHz。");
        }

        static void Info(string t) => MessageBox.Show(t, "Mirror2Out", MessageBoxButtons.OK, MessageBoxIcon.Information);

        static string[] ListRenderDevices()
        {
            var mm = new MMDeviceEnumerator();
            return mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                     .Select(d => d.FriendlyName).ToArray();
        }
        static MMDevice FindRender(string keyword)
        {
            var mm = new MMDeviceEnumerator();
            return mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                     .FirstOrDefault(d => d.FriendlyName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
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
            _tray.ShowBalloonTip(1500, "Mirror2Out", "已停止", ToolTipIcon.Info);
        }

        public void Dispose()
        {
            Stop();
            _tray.Visible = false;
            _tray.Dispose();
            _menu.Dispose();
        }
    }
}
