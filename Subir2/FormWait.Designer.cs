
namespace Subir2
{
    partial class FormWait
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormWait));
            this.progressBarWait = new System.Windows.Forms.ProgressBar();
            this.lblwait = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBarWait
            // 
            this.progressBarWait.Location = new System.Drawing.Point(12, 29);
            this.progressBarWait.Name = "progressBarWait";
            this.progressBarWait.Size = new System.Drawing.Size(476, 23);
            this.progressBarWait.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBarWait.TabIndex = 0;
            this.progressBarWait.UseWaitCursor = true;
            // 
            // lblwait
            // 
            this.lblwait.AutoSize = true;
            this.lblwait.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblwait.Location = new System.Drawing.Point(213, 9);
            this.lblwait.Name = "lblwait";
            this.lblwait.Size = new System.Drawing.Size(72, 17);
            this.lblwait.TabIndex = 1;
            this.lblwait.Text = "Cargando...";
            this.lblwait.UseWaitCursor = true;
            // 
            // FormWait
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(500, 68);
            this.Controls.Add(this.lblwait);
            this.Controls.Add(this.progressBarWait);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormWait";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Cargando...";
            this.UseWaitCursor = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBarWait;
        private System.Windows.Forms.Label lblwait;
    }
}