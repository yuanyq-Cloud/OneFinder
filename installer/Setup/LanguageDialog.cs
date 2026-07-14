using System;
using System.Drawing;
using System.Windows.Forms;

namespace OneFinder.Setup
{
    /// <summary>
    /// Language selection dialog for the OneFinder Setup bootstrapper.
    /// </summary>
    public partial class LanguageDialog : Form
    {
        public string SelectedLanguage { get; private set; } = "zh-CN";

        private RadioButton _rbChinese = null!;
        private RadioButton _rbEnglish = null!;
        private Button _installButton = null!;

        public LanguageDialog()
        {
            BuildUI();
        }

        private void BuildUI()
        {
            Text = "OneFinder Setup";
            Size = new Size(440, 300);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;
            Icon = null;

            // --- Header banner ---
            var banner = new Panel
            {
                BackColor = Color.FromArgb(128, 57, 123),
                Height = 56,
                Dock = DockStyle.Top,
            };

            var bannerLabel = new Label
            {
                Text = "  OneFinder Setup",
                Font = new Font("Segoe UI", 16f, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(14, 10),
                BackColor = Color.Transparent,
            };
            banner.Controls.Add(bannerLabel);

            // --- Content area ---
            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(24, 16, 24, 16),
                BackColor = Color.White,
            };

            var descLabel = new Label
            {
                Text = "Select your display language / 请选择显示语言",
                Font = new Font("Segoe UI", 11f),
                ForeColor = Color.FromArgb(50, 50, 50),
                AutoSize = true,
                Location = new Point(0, 4),
            };

            _rbChinese = new RadioButton
            {
                Text = "  中文 (Chinese)",
                AutoSize = true,
                Font = new Font("Microsoft YaHei", 11f),
                Location = new Point(16, 46),
                Checked = true,
                Cursor = Cursors.Hand,
            };
            _rbChinese.CheckedChanged += (s, e) =>
            {
                if (_rbChinese.Checked) SelectedLanguage = "zh-CN";
            };

            _rbEnglish = new RadioButton
            {
                Text = "  English",
                AutoSize = true,
                Font = new Font("Segoe UI", 11f),
                Location = new Point(16, 76),
                Cursor = Cursors.Hand,
            };
            _rbEnglish.CheckedChanged += (s, e) =>
            {
                if (_rbEnglish.Checked) SelectedLanguage = "en-US";
            };

            var hintLabel = new Label
            {
                Text = "You can change this later via the registry.\n以后可以通过注册表更改。",
                Font = new Font("Segoe UI", 8f),
                ForeColor = Color.FromArgb(140, 140, 140),
                AutoSize = true,
                Location = new Point(16, 110),
            };

            // --- Install button ---
            _installButton = new Button
            {
                Text = "  Install  /  安装  ",
                Size = new Size(170, 38),
                BackColor = Color.FromArgb(128, 57, 123),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                Cursor = Cursors.Hand,
            };
            _installButton.FlatAppearance.BorderSize = 0;
            _installButton.Click += (s, e) =>
            {
                DialogResult = DialogResult.OK;
                Close();
            };

            // Center the install button
            _installButton.Location = new Point(
                (contentPanel.Width - _installButton.Width) / 2 - contentPanel.Padding.Left,
                155);

            contentPanel.Controls.Add(descLabel);
            contentPanel.Controls.Add(_rbChinese);
            contentPanel.Controls.Add(_rbEnglish);
            contentPanel.Controls.Add(hintLabel);
            contentPanel.Controls.Add(_installButton);

            Controls.Add(contentPanel);
            Controls.Add(banner);
        }
    }
}
