namespace InventorAutomation
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnStartInventor;
        private System.Windows.Forms.Button btnCreatePart;
        private System.Windows.Forms.Button btnCreateSketch;
        private System.Windows.Forms.Button btnSaveAndClose;
        private System.Windows.Forms.Label lblTitle;

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
            this.Size = new System.Drawing.Size(350, 350);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.Color.LightGray;

            lblTitle = new System.Windows.Forms.Label()
            {
                Text = "Autodesk Inventor Kontrol Paneli",
                Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold),
                AutoSize = true,
                Location = new System.Drawing.Point(50, 20)
            };
            this.Controls.Add(lblTitle);

            btnStartInventor = new System.Windows.Forms.Button()
            {
                Text = "Autodesk Inventor Başlat",
                Size = new System.Drawing.Size(250, 40),
                Location = new System.Drawing.Point(50, 60),
                BackColor = System.Drawing.Color.LightBlue,
                Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold)
            };
            btnStartInventor.Click += new System.EventHandler(this.btnStartInventor_Click);
            this.Controls.Add(btnStartInventor);

            btnCreatePart = new System.Windows.Forms.Button()
            {
                Text = "Yeni Parça Oluştur",
                Size = new System.Drawing.Size(250, 40),
                Location = new System.Drawing.Point(50, 110),
                BackColor = System.Drawing.Color.LightGreen,
                Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold)
            };
            btnCreatePart.Click += new System.EventHandler(this.btnCreatePart_Click);
            this.Controls.Add(btnCreatePart);

            btnCreateSketch = new System.Windows.Forms.Button()
            {
                Text = "Sketch Ekle",
                Size = new System.Drawing.Size(250, 40),
                Location = new System.Drawing.Point(50, 160),
                BackColor = System.Drawing.Color.LightGoldenrodYellow,
                Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold)
            };
            btnCreateSketch.Click += new System.EventHandler(this.btnCreateSketch_Click);
            this.Controls.Add(btnCreateSketch);

            btnSaveAndClose = new System.Windows.Forms.Button()
            {
                Text = "Kaydet ve Kapat",
                Size = new System.Drawing.Size(250, 40),
                Location = new System.Drawing.Point(50, 210),
                BackColor = System.Drawing.Color.LightCoral,
                Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold)
            };
            btnSaveAndClose.Click += new System.EventHandler(this.btnSaveAndClose_Click);
            this.Controls.Add(btnSaveAndClose);
        }
    }
}
