using System.Diagnostics;

namespace ANB.RenameBatch.GUI;

public partial class Form1 : Form
{
    private static readonly Color Bg = Color.FromArgb(24, 24, 27);
    private static readonly Color PanelBg = Color.FromArgb(32, 33, 37);
    private static readonly Color InputBg = Color.FromArgb(20, 21, 24);
    private static readonly Color TextMain = Color.FromArgb(234, 236, 240);
    private static readonly Color TextMuted = Color.FromArgb(160, 164, 171);
    private static readonly Color Accent = Color.FromArgb(74, 130, 255);
    private static readonly Color AccentHover = Color.FromArgb(92, 148, 255);
    private static readonly Color Border = Color.FromArgb(55, 57, 62);

    private TextBox _pathText = null!;
    private TextBox _destText = null!;
    private TextBox _prefixText = null!;
    private TextBox _includeExtText = null!;
    private TextBox _excludeExtText = null!;
    private ComboBox _orderCombo = null!;
    private ComboBox _conflictCombo = null!;
    private CheckBox _streamCheck = null!;
    private CheckBox _recurseCheck = null!;
    private CheckBox _perExtCheck = null!;
    private CheckBox _includeHiddenCheck = null!;
    private CheckBox _dryRunCheck = null!;
    private CheckBox _preserveCheck = null!;
    private CheckBox _noLogCheck = null!;
    private CheckBox _noHashCheck = null!;
    private NumericUpDown _startNum = null!;
    private NumericUpDown _padNum = null!;
    private NumericUpDown _minPadNum = null!;
    private NumericUpDown _progressEveryNum = null!;
    private NumericUpDown _confirmThresholdNum = null!;
    private TextBox _outputBox = null!;
    private Button _runButton = null!;
    private Label _scriptLabel = null!;
    private Process? _process;
    private bool _suppressOutput;

    public Form1()
    {
        InitializeComponent();
        BuildUi();
        WireEvents();
    }

    private void BuildUi()
    {
        Text = "ANB Rename Batch - Developed by Lahiru Sanjika";
        var appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        if (appIcon is not null)
        {
            Icon = appIcon;
        }
        MinimumSize = new Size(980, 720);
        Size = new Size(1040, 780);
        BackColor = Bg;
        ForeColor = TextMain;
        Font = new Font("Segoe UI", 9.5f);

        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = 400,
            FixedPanel = FixedPanel.Panel1
        };
        split.Panel1.BackColor = Bg;
        split.Panel2.BackColor = Bg;

        var controlsPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            RowCount = 9,
            Padding = new Padding(12),
        };

        layout.Controls.Add(BuildHeaderRow(), 0, 0);
        layout.Controls.Add(BuildPathRow(), 0, 1);
        layout.Controls.Add(BuildDestinationRow(), 0, 2);
        layout.Controls.Add(BuildOptionsRow(), 0, 3);
        layout.Controls.Add(BuildNumbersRow(), 0, 4);
        layout.Controls.Add(BuildExtensionsRow(), 0, 5);
        layout.Controls.Add(BuildAdvancedRow(), 0, 6);
        layout.Controls.Add(BuildRunRow(), 0, 7);
        layout.Controls.Add(BuildScriptRow(), 0, 8);

        controlsPanel.Controls.Add(layout);
        split.Panel1.Controls.Add(controlsPanel);

        _outputBox = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            Font = new Font("Consolas", 10f),
            BackColor = Color.FromArgb(18, 18, 20),
            ForeColor = Color.FromArgb(208, 212, 219),
            BorderStyle = BorderStyle.FixedSingle,
            WordWrap = false
        };
        split.Panel2.Controls.Add(_outputBox);

        Controls.Add(split);

        ApplyTheme(this);
        _scriptLabel.ForeColor = TextMuted;
        _runButton.BackColor = Accent;
        _runButton.ForeColor = Color.White;
        _runButton.FlatAppearance.BorderSize = 0;
        _runButton.FlatAppearance.MouseOverBackColor = AccentHover;

        Shown += (_, _) => AdjustLayout(split, controlsPanel, layout);
        Resize += (_, _) => AdjustLayout(split, controlsPanel, layout);
    }

    private Control BuildHeaderRow()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            Height = 52,
            BackColor = PanelBg,
            Padding = new Padding(10)
        };

        var title = new Label
        {
            Text = "ANB Rename Batch",
            AutoSize = true,
            Font = new Font("Segoe UI Semibold", 14f),
            ForeColor = TextMain
        };
        var subtitle = new Label
        {
            Text = "Batch rename, move, and log files safely",
            AutoSize = true,
            Top = 26,
            Left = 2,
            ForeColor = TextMuted
        };

        panel.Controls.Add(title);
        panel.Controls.Add(subtitle);
        return panel;
    }

    private Control BuildPathRow()
    {
        var panel = BuildRowPanel();
        panel.Controls.Add(new Label { Text = "Source Folder:", AutoSize = true, Margin = new Padding(0, 8, 8, 0) });
        _pathText = new TextBox { Width = 520, PlaceholderText = "Select a folder..." };
        var browse = new Button { Text = "Browse..." };
        browse.Click += (_, _) => BrowseFolder(_pathText);
        panel.Controls.Add(_pathText);
        panel.Controls.Add(browse);
        StyleRow(panel);
        return panel;
    }

    private Control BuildDestinationRow()
    {
        var panel = BuildRowPanel();
        panel.Controls.Add(new Label { Text = "Destination (optional):", AutoSize = true, Margin = new Padding(0, 8, 8, 0) });
        _destText = new TextBox { Width = 430, PlaceholderText = "Leave empty to rename in place" };
        var browse = new Button { Text = "Browse..." };
        browse.Click += (_, _) => BrowseFolder(_destText);
        _preserveCheck = new CheckBox { Text = "Preserve folders", AutoSize = true, Enabled = false };
        panel.Controls.Add(_destText);
        panel.Controls.Add(browse);
        panel.Controls.Add(_preserveCheck);
        StyleRow(panel);
        return panel;
    }

    private Control BuildOptionsRow()
    {
        var panel = BuildRowPanel();
        panel.Controls.Add(new Label { Text = "Order:", AutoSize = true, Margin = new Padding(0, 8, 4, 0) });
        _orderCombo = new ComboBox { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
        _orderCombo.Items.AddRange(new[] { "Time", "Name", "Size", "None" });
        _orderCombo.SelectedIndex = 0;

        _streamCheck = new CheckBox { Text = "Stream", AutoSize = true, Margin = new Padding(12, 6, 4, 0) };
        _recurseCheck = new CheckBox { Text = "Recurse", AutoSize = true, Margin = new Padding(8, 6, 4, 0) };
        _perExtCheck = new CheckBox { Text = "Per extension", AutoSize = true, Margin = new Padding(8, 6, 4, 0) };
        _includeHiddenCheck = new CheckBox { Text = "Include hidden", AutoSize = true, Margin = new Padding(8, 6, 4, 0) };
        _dryRunCheck = new CheckBox { Text = "Dry run", AutoSize = true, Margin = new Padding(8, 6, 4, 0) };

        panel.Controls.Add(_orderCombo);
        panel.Controls.Add(_streamCheck);
        panel.Controls.Add(_recurseCheck);
        panel.Controls.Add(_perExtCheck);
        panel.Controls.Add(_includeHiddenCheck);
        panel.Controls.Add(_dryRunCheck);

        panel.Controls.Add(new Label { Text = "OnConflict:", AutoSize = true, Margin = new Padding(12, 8, 4, 0) });
        _conflictCombo = new ComboBox { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
        _conflictCombo.Items.AddRange(new[] { "Error", "Skip", "Overwrite" });
        _conflictCombo.SelectedIndex = 0;
        panel.Controls.Add(_conflictCombo);
        StyleRow(panel);
        return panel;
    }

    private Control BuildNumbersRow()
    {
        var panel = BuildRowPanel();
        panel.Controls.Add(new Label { Text = "Prefix:", AutoSize = true, Margin = new Padding(0, 8, 4, 0) });
        _prefixText = new TextBox { Width = 120 };

        panel.Controls.Add(new Label { Text = "Start:", AutoSize = true, Margin = new Padding(12, 8, 4, 0) });
        _startNum = new NumericUpDown { Minimum = 0, Maximum = 100000000, Value = 1, Width = 80 };

        panel.Controls.Add(new Label { Text = "Pad:", AutoSize = true, Margin = new Padding(12, 8, 4, 0) });
        _padNum = new NumericUpDown { Minimum = 0, Maximum = 20, Value = 0, Width = 60 };

        panel.Controls.Add(new Label { Text = "MinPad:", AutoSize = true, Margin = new Padding(12, 8, 4, 0) });
        _minPadNum = new NumericUpDown { Minimum = 1, Maximum = 20, Value = 4, Width = 60 };

        panel.Controls.Add(_prefixText);
        panel.Controls.Add(_startNum);
        panel.Controls.Add(_padNum);
        panel.Controls.Add(_minPadNum);
        StyleRow(panel);
        return panel;
    }

    private Control BuildExtensionsRow()
    {
        var panel = BuildRowPanel();
        panel.Controls.Add(new Label { Text = "Include ext:", AutoSize = true, Margin = new Padding(0, 8, 4, 0) });
        _includeExtText = new TextBox { Width = 180, PlaceholderText = "jpg,mp4" };
        panel.Controls.Add(_includeExtText);
        panel.Controls.Add(new Label { Text = "Exclude ext:", AutoSize = true, Margin = new Padding(12, 8, 4, 0) });
        _excludeExtText = new TextBox { Width = 180, PlaceholderText = "tmp,ps1" };
        panel.Controls.Add(_excludeExtText);
        StyleRow(panel);
        return panel;
    }

    private Control BuildAdvancedRow()
    {
        var panel = BuildRowPanel();
        _noLogCheck = new CheckBox { Text = "No log", AutoSize = true };
        _noHashCheck = new CheckBox { Text = "No hash", AutoSize = true, Margin = new Padding(8, 6, 4, 0) };

        panel.Controls.Add(_noLogCheck);
        panel.Controls.Add(_noHashCheck);

        panel.Controls.Add(new Label { Text = "Progress every:", AutoSize = true, Margin = new Padding(12, 8, 4, 0) });
        _progressEveryNum = new NumericUpDown { Minimum = 0, Maximum = 1000000, Value = 1000, Width = 80 };
        panel.Controls.Add(_progressEveryNum);

        panel.Controls.Add(new Label { Text = "Confirm threshold:", AutoSize = true, Margin = new Padding(12, 8, 4, 0) });
        _confirmThresholdNum = new NumericUpDown { Minimum = 0, Maximum = 100000000, Value = 10000, Width = 100 };
        panel.Controls.Add(_confirmThresholdNum);

        StyleRow(panel);
        return panel;
    }

    private Control BuildRunRow()
    {
        var panel = BuildRowPanel();
        _runButton = new Button { Text = "Run", Width = 140, Height = 34 };
        var clearButton = new Button { Text = "Clear Output", Width = 140, Height = 34 };
        clearButton.Click += (_, _) => _outputBox.Clear();
        panel.Controls.Add(_runButton);
        panel.Controls.Add(clearButton);
        StyleRow(panel);
        return panel;
    }

    private Control BuildScriptRow()
    {
        var panel = BuildRowPanel();
        _scriptLabel = new Label { AutoSize = true, ForeColor = Color.DimGray };
        panel.Controls.Add(_scriptLabel);
        StyleRow(panel);
        return panel;
    }

    private FlowLayoutPanel BuildRowPanel()
    {
        return new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true,
            Font = Font,
            BackColor = PanelBg,
            Padding = new Padding(8, 6, 8, 6),
            Margin = new Padding(0, 0, 0, 8)
        };
    }

    private void StyleRow(FlowLayoutPanel panel)
    {
        var rowHeight = GetRowHeight(panel);
        foreach (Control c in panel.Controls)
        {
            StyleControl(c, rowHeight);
        }
    }

    private int GetRowHeight(FlowLayoutPanel panel)
    {
        var max = 0;
        foreach (Control c in panel.Controls)
        {
            var height = GetPreferredHeight(c);
            if (height > max) { max = height; }
        }
        var rowHeight = max + 2;
        return Math.Max(26, rowHeight);
    }

    private int GetPreferredHeight(Control c)
    {
        switch (c)
        {
            case TextBox tb:
                return tb.PreferredHeight;
            case ComboBox cb:
                return cb.PreferredHeight;
            case NumericUpDown nud:
                return nud.PreferredHeight;
            case Button btn:
                return btn.PreferredSize.Height;
            case Label lbl:
                lbl.UseCompatibleTextRendering = true;
                return lbl.PreferredSize.Height;
            case CheckBox chk:
                chk.UseCompatibleTextRendering = true;
                return chk.PreferredSize.Height;
            default:
                return c.PreferredSize.Height;
        }
    }

    private void StyleControl(Control c, int rowHeight)
    {
        switch (c)
        {
            case Label lbl:
                lbl.TextAlign = ContentAlignment.MiddleLeft;
                lbl.UseCompatibleTextRendering = true;
                var labelTop = Math.Max(0, (rowHeight - lbl.PreferredSize.Height) / 2);
                lbl.Margin = new Padding(lbl.Margin.Left, labelTop, lbl.Margin.Right, 0);
                break;
            case TextBox tb:
                tb.AutoSize = false;
                tb.Height = rowHeight;
                tb.Margin = new Padding(tb.Margin.Left, 0, tb.Margin.Right, 0);
                break;
            case ComboBox cb:
                cb.IntegralHeight = false;
                cb.Height = rowHeight;
                cb.Margin = new Padding(cb.Margin.Left, 0, cb.Margin.Right, 0);
                break;
            case NumericUpDown nud:
                nud.Height = rowHeight;
                nud.Margin = new Padding(nud.Margin.Left, 0, nud.Margin.Right, 0);
                break;
            case Button btn:
                btn.AutoSize = false;
                btn.Height = rowHeight;
                btn.TextAlign = ContentAlignment.MiddleCenter;
                btn.UseCompatibleTextRendering = true;
                btn.Padding = new Padding(0);
                btn.Margin = new Padding(btn.Margin.Left, 0, btn.Margin.Right, 0);
                break;
            case CheckBox chk:
                chk.UseCompatibleTextRendering = true;
                var checkTop = Math.Max(0, (rowHeight - chk.PreferredSize.Height) / 2);
                chk.Margin = new Padding(chk.Margin.Left, checkTop, chk.Margin.Right, 0);
                break;
        }
    }

    private void AdjustLayout(SplitContainer split, Panel controlsPanel, TableLayoutPanel layout)
    {
        layout.PerformLayout();
        var desired = layout.PreferredSize.Height + 20;
        var minOutput = 180;
        var maxTop = Math.Max(260, split.Height - minOutput);
        if (desired <= maxTop)
        {
            controlsPanel.AutoScroll = false;
            split.SplitterDistance = desired;
        }
        else
        {
            controlsPanel.AutoScroll = true;
            split.SplitterDistance = maxTop;
        }
    }

    private void ApplyTheme(Control root)
    {
        foreach (Control c in root.Controls)
        {
            switch (c)
            {
                case TextBox tb:
                    tb.BackColor = InputBg;
                    tb.ForeColor = TextMain;
                    tb.BorderStyle = BorderStyle.FixedSingle;
                    break;
                case ComboBox cb:
                    cb.BackColor = InputBg;
                    cb.ForeColor = TextMain;
                    cb.FlatStyle = FlatStyle.Flat;
                    break;
                case NumericUpDown nud:
                    nud.BackColor = InputBg;
                    nud.ForeColor = TextMain;
                    break;
                case Button btn:
                    btn.BackColor = PanelBg;
                    btn.ForeColor = TextMain;
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderColor = Border;
                    btn.FlatAppearance.BorderSize = 1;
                    break;
                case Label lbl:
                    if (lbl != _scriptLabel)
                    {
                        lbl.ForeColor = TextMain;
                    }
                    break;
                case CheckBox chk:
                    chk.ForeColor = TextMain;
                    break;
                case FlowLayoutPanel flp:
                    flp.BackColor = PanelBg;
                    break;
                case TableLayoutPanel tlp:
                    tlp.BackColor = Bg;
                    break;
                case Panel panel when panel is not FlowLayoutPanel && panel is not TableLayoutPanel:
                    if (panel != root) { panel.BackColor = PanelBg; }
                    break;
            }

            if (c.HasChildren)
            {
                ApplyTheme(c);
            }
        }
    }

    private void WireEvents()
    {
        _streamCheck.CheckedChanged += (_, _) =>
        {
            if (_streamCheck.Checked)
            {
                _orderCombo.SelectedItem = "None";
                _orderCombo.Enabled = false;
                _perExtCheck.Checked = false;
                _perExtCheck.Enabled = false;
            }
            else
            {
                _orderCombo.Enabled = true;
                _perExtCheck.Enabled = true;
            }
        };

        _destText.TextChanged += (_, _) =>
        {
            _preserveCheck.Enabled = !string.IsNullOrWhiteSpace(_destText.Text);
            if (!_preserveCheck.Enabled) { _preserveCheck.Checked = false; }
        };

        _runButton.Click += (_, _) => RunTool();
        FormClosing += OnFormClosing;
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_process is { HasExited: false })
        {
            var result = MessageBox.Show(
                this,
                "A rename is still running. Do you want to cancel it?",
                "ANB Rename Batch",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }

            if (result == DialogResult.Yes)
            {
                try
                {
                    _process.Kill(true);
                    _process.WaitForExit(2000);
                }
                catch
                {
                    // Ignore kill failures on exit.
                }
            }
            else
            {
                _suppressOutput = true;
            }
        }
    }

    private void BrowseFolder(TextBox target)
    {
        using var dialog = new FolderBrowserDialog();
        if (Directory.Exists(target.Text)) { dialog.SelectedPath = target.Text; }
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            target.Text = dialog.SelectedPath;
        }
    }

    private void RunTool()
    {
        if (_process is { HasExited: false })
        {
            MessageBox.Show(this, "A run is already in progress.", "ANB Rename Batch");
            return;
        }

        var path = _pathText.Text.Trim();
        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
        {
            MessageBox.Show(this, "Please select a valid source folder.", "ANB Rename Batch");
            return;
        }

        var confirm = MessageBox.Show(this, "This will rename files in the selected folder. Continue?", "ANB Rename Batch", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (confirm != DialogResult.Yes)
        {
            return;
        }

        var scriptPath = Path.Combine(AppContext.BaseDirectory, "ANB-RenameBatch.ps1");
        _scriptLabel.Text = $"Script: {scriptPath}";
        if (!File.Exists(scriptPath))
        {
            MessageBox.Show(this, "ANB-RenameBatch.ps1 was not found next to this EXE.", "ANB Rename Batch");
            return;
        }

        var args = new List<string>
        {
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", scriptPath,
            "-Path", path,
            "-Order", _orderCombo.SelectedItem?.ToString() ?? "Time",
            "-Start", _startNum.Value.ToString(),
            "-Pad", _padNum.Value.ToString(),
            "-MinPad", _minPadNum.Value.ToString(),
            "-OnConflict", _conflictCombo.SelectedItem?.ToString() ?? "Error",
            "-ProgressEvery", _progressEveryNum.Value.ToString(),
            "-ConfirmThreshold", _confirmThresholdNum.Value.ToString(),
            "-Force"
        };

        if (!string.IsNullOrWhiteSpace(_prefixText.Text))
        {
            args.Add("-Prefix");
            args.Add(_prefixText.Text.Trim());
        }

        if (!string.IsNullOrWhiteSpace(_destText.Text))
        {
            args.Add("-Destination");
            args.Add(_destText.Text.Trim());
            if (_preserveCheck.Checked) { args.Add("-PreserveFolders"); }
        }

        if (_streamCheck.Checked) { args.Add("-Stream"); }
        if (_recurseCheck.Checked) { args.Add("-Recurse"); }
        if (_perExtCheck.Checked) { args.Add("-PerExtension"); }
        if (_includeHiddenCheck.Checked) { args.Add("-IncludeHidden"); }
        if (_dryRunCheck.Checked) { args.Add("-DryRun"); }
        if (_noLogCheck.Checked) { args.Add("-NoLog"); }
        if (_noHashCheck.Checked) { args.Add("-NoHash"); }

        AddExtensionsArg(args, "-Extensions", _includeExtText.Text);
        AddExtensionsArg(args, "-ExcludeExtensions", _excludeExtText.Text);

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        foreach (var arg in args)
        {
            psi.ArgumentList.Add(arg);
        }

        _outputBox.AppendText($"> powershell {string.Join(" ", args)}{Environment.NewLine}");

        _process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        _process.OutputDataReceived += ProcessOutput;
        _process.ErrorDataReceived += ProcessOutput;
        _process.Exited += ProcessExited;

        _runButton.Enabled = false;
        _process.Start();
        _process.BeginOutputReadLine();
        _process.BeginErrorReadLine();
    }

    private void AddExtensionsArg(List<string> args, string flag, string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) { return; }
        var parts = raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0) { return; }
        args.Add(flag);
        args.AddRange(parts);
    }

    private void AppendOutput(string? text)
    {
        if (_suppressOutput || IsDisposed || Disposing) { return; }
        if (string.IsNullOrWhiteSpace(text)) { return; }
        if (InvokeRequired)
        {
            try
            {
                Invoke(() => AppendOutput(text));
            }
            catch
            {
                // Ignore output failures during shutdown.
            }
            return;
        }
        _outputBox.AppendText(text + Environment.NewLine);
    }

    private void ProcessOutput(object? sender, DataReceivedEventArgs e)
    {
        AppendOutput(e.Data);
    }

    private void ProcessExited(object? sender, EventArgs e)
    {
        if (_suppressOutput || IsDisposed || Disposing) { return; }
        try
        {
            Invoke(() =>
            {
                AppendOutput($"[exit {_process?.ExitCode}]");
                _runButton.Enabled = true;
            });
        }
        catch
        {
            // Ignore exit handler failures during shutdown.
        }
    }
}
