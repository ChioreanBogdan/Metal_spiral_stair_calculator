namespace Calculator_spirala
{
    partial class Meniu_principal
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Meniu_principal));
            Poza_stalp = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)Poza_stalp).BeginInit();
            SuspendLayout();
            // 
            // Poza_stalp
            // 
            Poza_stalp.Image = (Image)resources.GetObject("Poza_stalp.Image");
            Poza_stalp.Location = new Point(12, 12);
            Poza_stalp.Name = "Poza_stalp";
            Poza_stalp.Size = new Size(118, 501);
            Poza_stalp.SizeMode = PictureBoxSizeMode.Zoom;
            Poza_stalp.TabIndex = 0;
            Poza_stalp.TabStop = false;
            Poza_stalp.Click += Poza_stalp_Click;
            // 
            // Meniu_principal
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(878, 513);
            Controls.Add(Poza_stalp);
            Name = "Meniu_principal";
            Text = "Calculator spirala";
            ((System.ComponentModel.ISupportInitialize)Poza_stalp).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private PictureBox Poza_stalp;
    }
}