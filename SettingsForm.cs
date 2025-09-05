using System;
using System.Linq;
using System.Windows.Forms;
using NAudio.CoreAudioApi;

namespace Mirror2Out
{
    public class SettingsForm : Form
    {
        ComboBox cmbInput = new ComboBox();
        ComboBox cmbMain  = new ComboBox();
        ComboBox cmbAux   = new ComboBox();

        NumericUpDown numRate   = new NumericUpDown();
        NumericUpDown numBits   = new NumericUpDown();
        NumericUpDown numBufMain= new NumericUpDown();
        NumericUpDown numBufAux = new NumericUpDown();

        CheckBox chkExclusive = new CheckBox();

        Button btnOk = new Button();
        Button btnCancel = new Button();

        class DevItem
        {
            public string Id;
            public string Name;
            public override string ToString() { return Name; }
        }

        public AppSettings Result { get; private set; }

        public SettingsForm(AppSettings current)
        {
            this.Text = "Mirror2Out 设置（通道1/2/3）";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false; this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Width = 580; this.Height = 360;

            var lbl1 = new Label(){ Text="通道1 输入（可选录音或渲染环回）", Left=20, Top=20, Width=260 };
            var lbl2 = new Label(){ Text="通道2 主通道（低延迟/高音质）", Left=20, Top=70, Width=260 };
            var lbl3 = new Label(){ Text="通道3 副通道（普通延迟/低占用）", Left=20, Top=120, Width=260 };

            cmbInput.Left=300; cmbInput.Top=20; cmbInput.Width=250; cmbInput.DropDownStyle=ComboBoxStyle.DropDownList;
            cmbMain.Left =300; cmbMain.Top =70; cmbMain.Width =250; cmbMain.DropDownStyle=ComboBoxStyle.DropDownList;
            cmbAux.Left  =300; cmbAux.Top  =120; cmbAux.Width  =250; cmbAux.DropDownStyle=ComboBoxStyle.DropDownList;

            var lblRate = new Label(){ Text="主通道目标采样率(Hz)", Left=20, Top=170, Width=200 };
            numRate.Left=300; numRate.Top=170; numRate.Width=120; numRate.Maximum=384000; numRate.Minimum=44100; numRate.Increment=1000;

            var lblBits = new Label(){ Text="主通道路位深(bit)", Left=20, Top=200, Width=200 };
            numBits.Left=300; numBits.Top=200; numBits.Width=120; numBits.Maximum=32; numBits.Minimum=16; numBits.Increment=8;

            chkExclusive.Left=20; chkExclusive.Top=230; chkExclusive.Width=530; chkExclusive.Text="主通道优先使用独占模式（失败自动回退共享）";

            var lblBufMain = new Label(){ Text="主通道缓冲(ms)", Left=20, Top=260, Width=200 };
            numBufMain.Left=300; numBufMain.Top=260; numBufMain.Width=120; numBufMain.Maximum=200; numBufMain.Minimum=2; numBufMain.Value=8;

            var lblBufAux = new Label(){ Text="副通道缓冲(ms)", Left=20, Top=290, Width=200 };
            numBufAux.Left=300; numBufAux.Top=290; numBufAux.Width=120; numBufAux.Maximum=400; numBufAux.Minimum=20; numBufAux.Value=120;

            btnOk.Text="保存"; btnOk.Left=440; btnOk.Top=290; btnOk.Width=110;
            btnCancel.Text="取消"; btnCancel.Left=440; btnCancel.Top=260; btnCancel.Width=110;

            btnOk.Click += (s,e)=> { SaveAndClose(); };
            btnCancel.Click += (s,e)=> { this.DialogResult = DialogResult.Cancel; this.Close(); };

            this.Controls.Add(lbl1); this.Controls.Add(lbl2); this.Controls.Add(lbl3);
            this.Controls.Add(cmbInput); this.Controls.Add(cmbMain); this.Controls.Add(cmbAux);
            this.Controls.Add(lblRate); this.Controls.Add(numRate);
            this.Controls.Add(lblBits); this.Controls.Add(numBits);
            this.Controls.Add(chkExclusive);
            this.Controls.Add(lblBufMain); this.Controls.Add(numBufMain);
            this.Controls.Add(lblBufAux); this.Controls.Add(numBufAux);
            this.Controls.Add(btnOk); this.Controls.Add(btnCancel);

            // 载入设备列表
            LoadDevices();

            // 填入当前配置
            Result = new AppSettings();
            Result.InputDeviceId = current.InputDeviceId;
            Result.MainDeviceId  = current.MainDeviceId;
            Result.AuxDeviceId   = current.AuxDeviceId;
            Result.MainExclusive = current.MainExclusive;
            Result.MainRate      = current.MainRate;
            Result.MainBits      = current.MainBits;
            Result.MainBufMs     = current.MainBufMs;
            Result.AuxBufMs      = current.AuxBufMs;

            SelectById(cmbInput, current.InputDeviceId);
            SelectById(cmbMain,  current.MainDeviceId);
            SelectById(cmbAux,   current.AuxDeviceId);

            numRate.Value = Math.Max(numRate.Minimum, Math.Min(numRate.Maximum, current.MainRate));
            numBits.Value = Math.Max(numBits.Minimum, Math.Min(numBits.Maximum, current.MainBits));
            numBufMain.Value = Math.Max(numBufMain.Minimum, Math.Min(numBufMain.Maximum, current.MainBufMs));
            numBufAux.Value  = Math.Max(numBufAux.Minimum,  Math.Min(numBufAux.Maximum,  current.AuxBufMs));
            chkExclusive.Checked = current.MainExclusive;
        }

        void LoadDevices()
        {
            var mm = new MMDeviceEnumerator();

            // 输入：录音 + 渲染（供选择环回）
            cmbInput.Items.Clear();
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
                cmbInput.Items.Add(new DevItem{ Id=d.ID, Name="录音: " + d.FriendlyName});
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
                cmbInput.Items.Add(new DevItem{ Id=d.ID, Name="环回: " + d.FriendlyName});

            // 输出（主/副）：渲染
            cmbMain.Items.Clear(); cmbAux.Items.Clear();
            foreach (var d in mm.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                var item = new DevItem{ Id=d.ID, Name=d.FriendlyName };
                cmbMain.Items.Add(item);
                cmbAux.Items.Add(new DevItem{ Id=d.ID, Name=d.FriendlyName });
            }
        }

        void SelectById(ComboBox cmb, string id)
        {
            if (string.IsNullOrEmpty(id) || cmb.Items.Count == 0) { if (cmb.Items.Count>0) cmb.SelectedIndex=0; return; }
            for (int i=0;i<cmb.Items.Count;i++)
            {
                var it = cmb.Items[i] as DevItem;
                if (it != null && it.Id == id) { cmb.SelectedIndex = i; return; }
            }
            cmb.SelectedIndex = 0;
        }

        void SaveAndClose()
        {
            var inSel  = cmbInput.SelectedItem as DevItem;
            var mainSel= cmbMain.SelectedItem  as DevItem;
            var auxSel = cmbAux.SelectedItem   as DevItem;
            if (mainSel == null || auxSel == null)
            {
                MessageBox.Show("请至少选择主通道与副通道输出设备。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Result = new AppSettings
            {
                InputDeviceId = inSel != null ? inSel.Id : null,
                MainDeviceId  = mainSel.Id,
                AuxDeviceId   = auxSel.Id,
                MainExclusive = chkExclusive.Checked,
                MainRate      = (int)numRate.Value,
                MainBits      = (int)numBits.Value,
                MainBufMs     = (int)numBufMain.Value,
                AuxBufMs      = (int)numBufAux.Value
            };
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
