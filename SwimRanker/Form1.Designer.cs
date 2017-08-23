namespace SwimRanker
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.runningButton = new System.Windows.Forms.Button();
            this.progressBarHTML = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.exWomen = new System.Windows.Forms.CheckBox();
            this.exMen = new System.Windows.Forms.CheckBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.ファイルToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.オプションToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.終了ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.大会データ変更ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.標準記録変更ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.yearSelect = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // runningButton
            // 
            this.runningButton.Font = new System.Drawing.Font("MS UI Gothic", 35F);
            this.runningButton.Location = new System.Drawing.Point(16, 43);
            this.runningButton.Margin = new System.Windows.Forms.Padding(4);
            this.runningButton.Name = "runningButton";
            this.runningButton.Size = new System.Drawing.Size(182, 112);
            this.runningButton.TabIndex = 0;
            this.runningButton.Text = "実行";
            this.runningButton.UseVisualStyleBackColor = true;
            this.runningButton.Click += new System.EventHandler(this.runningButton_Click);
            // 
            // progressBarHTML
            // 
            this.progressBarHTML.Location = new System.Drawing.Point(13, 291);
            this.progressBarHTML.Name = "progressBarHTML";
            this.progressBarHTML.Size = new System.Drawing.Size(354, 23);
            this.progressBarHTML.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 245);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 15);
            this.label1.TabIndex = 3;
            // 
            // exWomen
            // 
            this.exWomen.AutoSize = true;
            this.exWomen.Location = new System.Drawing.Point(218, 131);
            this.exWomen.Name = "exWomen";
            this.exWomen.Size = new System.Drawing.Size(89, 19);
            this.exWomen.TabIndex = 4;
            this.exWomen.Text = "女子除外";
            this.exWomen.UseVisualStyleBackColor = true;
            // 
            // exMen
            // 
            this.exMen.AutoSize = true;
            this.exMen.Location = new System.Drawing.Point(218, 106);
            this.exMen.Name = "exMen";
            this.exMen.Size = new System.Drawing.Size(89, 19);
            this.exMen.TabIndex = 5;
            this.exMen.Text = "男子除外";
            this.exMen.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ファイルToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(392, 28);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // ファイルToolStripMenuItem
            // 
            this.ファイルToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.オプションToolStripMenuItem,
            this.終了ToolStripMenuItem});
            this.ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem";
            this.ファイルToolStripMenuItem.Size = new System.Drawing.Size(63, 24);
            this.ファイルToolStripMenuItem.Text = "ファイル";
            // 
            // オプションToolStripMenuItem
            // 
            this.オプションToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.大会データ変更ToolStripMenuItem,
            this.標準記録変更ToolStripMenuItem});
            this.オプションToolStripMenuItem.Name = "オプションToolStripMenuItem";
            this.オプションToolStripMenuItem.Size = new System.Drawing.Size(181, 26);
            this.オプションToolStripMenuItem.Text = "オプション";
            // 
            // 終了ToolStripMenuItem
            // 
            this.終了ToolStripMenuItem.Name = "終了ToolStripMenuItem";
            this.終了ToolStripMenuItem.Size = new System.Drawing.Size(181, 26);
            this.終了ToolStripMenuItem.Text = "終了";
            this.終了ToolStripMenuItem.Click += new System.EventHandler(this.終了ToolStripMenuItem_Click);
            // 
            // 大会データ変更ToolStripMenuItem
            // 
            this.大会データ変更ToolStripMenuItem.Name = "大会データ変更ToolStripMenuItem";
            this.大会データ変更ToolStripMenuItem.Size = new System.Drawing.Size(181, 26);
            this.大会データ変更ToolStripMenuItem.Text = "大会データ変更";
            this.大会データ変更ToolStripMenuItem.Click += new System.EventHandler(this.大会データ変更ToolStripMenuItem_Click);
            // 
            // 標準記録変更ToolStripMenuItem
            // 
            this.標準記録変更ToolStripMenuItem.Name = "標準記録変更ToolStripMenuItem";
            this.標準記録変更ToolStripMenuItem.Size = new System.Drawing.Size(181, 26);
            this.標準記録変更ToolStripMenuItem.Text = "標準記録変更";
            this.標準記録変更ToolStripMenuItem.Click += new System.EventHandler(this.標準記録変更ToolStripMenuItem_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // yearSelect
            // 
            this.yearSelect.FormattingEnabled = true;
            this.yearSelect.Location = new System.Drawing.Point(298, 43);
            this.yearSelect.Name = "yearSelect";
            this.yearSelect.Size = new System.Drawing.Size(85, 23);
            this.yearSelect.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(215, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 15);
            this.label2.TabIndex = 8;
            this.label2.Text = "年度選択";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(392, 326);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.yearSelect);
            this.Controls.Add(this.exMen);
            this.Controls.Add(this.exWomen);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBarHTML);
            this.Controls.Add(this.runningButton);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "都道府県大会リザルト合体マシーン";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button runningButton;
        private System.Windows.Forms.ProgressBar progressBarHTML;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox exWomen;
        private System.Windows.Forms.CheckBox exMen;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ファイルToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem オプションToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 大会データ変更ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 標準記録変更ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 終了ToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.ComboBox yearSelect;
        private System.Windows.Forms.Label label2;
    }
}

