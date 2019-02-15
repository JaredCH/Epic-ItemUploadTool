namespace ItemUploadTool
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.glgrid3 = new System.Windows.Forms.DataGridView();
            this.sagrid4 = new System.Windows.Forms.DataGridView();
            this.fourthchar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.szes = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gldonebtn = new System.Windows.Forms.Button();
            this.thirdchar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cat = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.grps = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.glgrid3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sagrid4)).BeginInit();
            this.SuspendLayout();
            // 
            // glgrid3
            // 
            this.glgrid3.AllowUserToAddRows = false;
            this.glgrid3.AllowUserToDeleteRows = false;
            this.glgrid3.AllowUserToResizeColumns = false;
            this.glgrid3.AllowUserToResizeRows = false;
            this.glgrid3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.glgrid3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.thirdchar,
            this.cat,
            this.grps});
            this.glgrid3.Location = new System.Drawing.Point(12, 12);
            this.glgrid3.MultiSelect = false;
            this.glgrid3.Name = "glgrid3";
            this.glgrid3.ReadOnly = true;
            this.glgrid3.Size = new System.Drawing.Size(475, 212);
            this.glgrid3.TabIndex = 0;
            this.glgrid3.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.glgrid3_CellContentClick);
            // 
            // sagrid4
            // 
            this.sagrid4.AllowUserToAddRows = false;
            this.sagrid4.AllowUserToDeleteRows = false;
            this.sagrid4.AllowUserToResizeColumns = false;
            this.sagrid4.AllowUserToResizeRows = false;
            this.sagrid4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.sagrid4.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.fourthchar,
            this.szes});
            this.sagrid4.Location = new System.Drawing.Point(78, 230);
            this.sagrid4.MultiSelect = false;
            this.sagrid4.Name = "sagrid4";
            this.sagrid4.ReadOnly = true;
            this.sagrid4.Size = new System.Drawing.Size(345, 191);
            this.sagrid4.TabIndex = 3;
            this.sagrid4.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.sagrid4_CellContentClick);
            // 
            // fourthchar
            // 
            this.fourthchar.HeaderText = "4th Character Selection";
            this.fourthchar.Name = "fourthchar";
            // 
            // szes
            // 
            this.szes.HeaderText = "Sizes";
            this.szes.Name = "szes";
            this.szes.Width = 200;
            // 
            // gldonebtn
            // 
            this.gldonebtn.Location = new System.Drawing.Point(177, 438);
            this.gldonebtn.Name = "gldonebtn";
            this.gldonebtn.Size = new System.Drawing.Size(133, 31);
            this.gldonebtn.TabIndex = 4;
            this.gldonebtn.Text = "Done";
            this.gldonebtn.UseVisualStyleBackColor = true;
            this.gldonebtn.Click += new System.EventHandler(this.gldonebtn_Click);
            // 
            // thirdchar
            // 
            this.thirdchar.HeaderText = "3rd Character Selection";
            this.thirdchar.Name = "thirdchar";
            this.thirdchar.ReadOnly = true;
            // 
            // cat
            // 
            this.cat.HeaderText = "Category";
            this.cat.Name = "cat";
            this.cat.ReadOnly = true;
            this.cat.Width = 230;
            // 
            // grps
            // 
            this.grps.HeaderText = "Groups";
            this.grps.Name = "grps";
            this.grps.ReadOnly = true;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(499, 481);
            this.Controls.Add(this.gldonebtn);
            this.Controls.Add(this.sagrid4);
            this.Controls.Add(this.glgrid3);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form2";
            this.Text = "GL CLass Code Table";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.glgrid3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sagrid4)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView glgrid3;
        private System.Windows.Forms.DataGridView sagrid4;
        private System.Windows.Forms.DataGridViewTextBoxColumn fourthchar;
        private System.Windows.Forms.DataGridViewTextBoxColumn szes;
        private System.Windows.Forms.Button gldonebtn;
        private System.Windows.Forms.DataGridViewTextBoxColumn thirdchar;
        private System.Windows.Forms.DataGridViewTextBoxColumn cat;
        private System.Windows.Forms.DataGridViewTextBoxColumn grps;
    }
}