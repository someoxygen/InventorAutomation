namespace InventorAutomation
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.FlowLayoutPanel panelButtons;
        private System.Windows.Forms.Button btnStartInventor;
        private System.Windows.Forms.Button btnCreatePart;
        private System.Windows.Forms.Button btnRectangle;
        private System.Windows.Forms.Button btnTriangle;
        private System.Windows.Forms.Button btnCircle;
        private System.Windows.Forms.Button btnPentagon;
        private System.Windows.Forms.Button btnHexagon;
        private System.Windows.Forms.Button btnSaveAndClose;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.Text = "Inventor Automation";
            this.Size = new System.Drawing.Size(650, 850);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.Color.LightGray;

            panelHeader = new System.Windows.Forms.Panel()
            {
                Dock = DockStyle.Top,
                BackColor = System.Drawing.Color.DarkGray,
                Padding = new Padding(0, 20, 0, 20),
                AutoSize = true
            };
            this.Controls.Add(panelHeader);

            lblTitle = new System.Windows.Forms.Label()
            {
                Text = "Autodesk Inventor Kontrol Paneli",
                Font = new System.Drawing.Font("Arial", 16F, System.Drawing.FontStyle.Bold),
                AutoSize = false,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                ForeColor = System.Drawing.Color.White
            };
            panelHeader.Controls.Add(lblTitle);

            panelButtons = new System.Windows.Forms.FlowLayoutPanel()
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                Padding = new Padding(50, 100, 50, 20),
                BackColor = System.Drawing.Color.WhiteSmoke,
                AutoSize = false,
                WrapContents = false
            };
            this.Controls.Add(panelButtons);

            btnStartInventor = AddButton("Autodesk Inventor Başlat", System.Drawing.Color.LightBlue, btnStartInventor_Click);
            btnCreatePart = AddButton("Yeni Parça Oluştur", System.Drawing.Color.LightGreen, btnCreatePart_Click);
            btnRectangle = AddButton("Dikdörtgen Çiz", System.Drawing.Color.LightGoldenrodYellow, btnRectangle_Click);
            btnTriangle = AddButton("Üçgen Çiz", System.Drawing.Color.LightGoldenrodYellow, btnTriangle_Click);
            btnCircle = AddButton("Daire Çiz", System.Drawing.Color.LightGoldenrodYellow, btnCircle_Click);
            btnPentagon = AddButton("Beşgen Çiz", System.Drawing.Color.LightGoldenrodYellow, btnPentagon_Click);
            btnHexagon = AddButton("Altıgen Çiz", System.Drawing.Color.LightGoldenrodYellow, btnHexagon_Click);
            btnSaveAndClose = AddButton("Kaydet ve Kapat", System.Drawing.Color.LightCoral, btnSaveAndClose_Click);

            SetInitialButtonStates();
        }

        private Button AddButton(string text, System.Drawing.Color color, EventHandler clickEvent)
        {
            Button btn = new Button()
            {
                Text = text,
                AutoSize = false,
                Size = new System.Drawing.Size(500, 60),
                BackColor = color,
                Font = new System.Drawing.Font("Arial", 14F, System.Drawing.FontStyle.Bold),
                Margin = new Padding(10),
                FlatStyle = FlatStyle.Flat,
                Enabled = false 
            };

            btn.FlatAppearance.BorderSize = 0;
            btn.Click += clickEvent;
            panelButtons.Controls.Add(btn);
            return btn;
        }

        private void SetInitialButtonStates()
        {
            btnStartInventor.Enabled = true;
        }
    }
}
