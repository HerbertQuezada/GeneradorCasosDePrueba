
namespace CasosDePruebaECF
{
    partial class Main
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            xlsx = new Button();
            groupBox1 = new GroupBox();
            copy = new Button();
            casos = new ComboBox();
            path = new TextBox();
            groupBox2 = new GroupBox();
            documento = new TextBox();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            SuspendLayout();
            // 
            // xlsx
            // 
            xlsx.Image = (Image)resources.GetObject("xlsx.Image");
            xlsx.Location = new Point(657, 12);
            xlsx.Name = "xlsx";
            xlsx.Size = new Size(45, 43);
            xlsx.TabIndex = 0;
            xlsx.UseVisualStyleBackColor = true;
            xlsx.Click += xlsx_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(copy);
            groupBox1.Controls.Add(casos);
            groupBox1.Controls.Add(path);
            groupBox1.Controls.Add(xlsx);
            groupBox1.Location = new Point(12, 6);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(764, 61);
            groupBox1.TabIndex = 1;
            groupBox1.TabStop = false;
            // 
            // copy
            // 
            copy.Image = (Image)resources.GetObject("copy.Image");
            copy.Location = new Point(708, 12);
            copy.Name = "copy";
            copy.Size = new Size(45, 43);
            copy.TabIndex = 2;
            copy.UseVisualStyleBackColor = true;
            copy.Click += copy_Click;
            // 
            // casos
            // 
            casos.DropDownStyle = ComboBoxStyle.DropDownList;
            casos.FormattingEnabled = true;
            casos.Location = new Point(495, 21);
            casos.Name = "casos";
            casos.Size = new Size(156, 23);
            casos.TabIndex = 1;
            casos.SelectedIndexChanged += casos_SelectedIndexChanged;
            // 
            // path
            // 
            path.Location = new Point(6, 22);
            path.Name = "path";
            path.ReadOnly = true;
            path.Size = new Size(483, 23);
            path.TabIndex = 3;
            path.Text = " ";
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(documento);
            groupBox2.Location = new Point(12, 73);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(764, 472);
            groupBox2.TabIndex = 2;
            groupBox2.TabStop = false;
            // 
            // documento
            // 
            documento.Dock = DockStyle.Fill;
            documento.Location = new Point(3, 19);
            documento.Multiline = true;
            documento.Name = "documento";
            documento.ReadOnly = true;
            documento.ScrollBars = ScrollBars.Both;
            documento.Size = new Size(758, 450);
            documento.TabIndex = 0;
            // 
            // Main
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(789, 557);
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Main";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Main";
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        private GroupBox groupBox1;
        private Button xlsx;
        private TextBox path;
        private ComboBox casos;
        private Button copy;
        private GroupBox groupBox2;
        private TextBox documento;
    }
}
