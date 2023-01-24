namespace SelectAndCalculate
{
    partial class frmMessageCheckbox
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMessageCheckbox));
            this.chkOption = new System.Windows.Forms.CheckBox();
            this.lblOption = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // chkOption
            // 
            resources.ApplyResources(this.chkOption, "chkOption");
            this.chkOption.BackColor = System.Drawing.Color.Transparent;
            this.chkOption.Name = "chkOption";
            this.chkOption.UseVisualStyleBackColor = false;
            // 
            // lblOption
            // 
            this.lblOption.BackColor = System.Drawing.Color.Transparent;
            resources.ApplyResources(this.lblOption, "lblOption");
            this.lblOption.ForeColor = System.Drawing.Color.DarkBlue;
            this.lblOption.Name = "lblOption";
            // 
            // btnOK
            // 
            this.btnOK.Image = global::SelectAndCalculate.Properties.Resources.check;
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.Name = "btnOK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // frmMessageCheckbox
            // 
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblOption);
            this.Controls.Add(this.chkOption);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmMessageCheckbox";
            this.Load += new System.EventHandler(this.frmMessageCheckbox_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkOption;
        private System.Windows.Forms.Label lblOption;
        private System.Windows.Forms.Button btnOK;
    }
}
