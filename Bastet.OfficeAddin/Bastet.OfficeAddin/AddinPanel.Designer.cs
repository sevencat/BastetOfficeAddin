﻿namespace Bastet.OfficeAddin
{
    partial class AddinPanel
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.elementHostMain = new System.Windows.Forms.Integration.ElementHost();
            this.SuspendLayout();
            // 
            // elementHostMain
            // 
            this.elementHostMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHostMain.Location = new System.Drawing.Point(0, 0);
            this.elementHostMain.Name = "elementHostMain";
            this.elementHostMain.Size = new System.Drawing.Size(341, 334);
            this.elementHostMain.TabIndex = 0;
            this.elementHostMain.Text = "elementHost1";
            this.elementHostMain.Child = null;
            // 
            // AddinPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.elementHostMain);
            this.Name = "AddinPanel";
            this.Size = new System.Drawing.Size(341, 334);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHostMain;
    }
}
