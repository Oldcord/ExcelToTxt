namespace From_xls_to_txt {
    partial class Excel {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            this.buttonSave = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.cboSheet = new System.Windows.Forms.ComboBox();
            this.textPathFile = new System.Windows.Forms.RichTextBox();
            this.btnOpen = new System.Windows.Forms.Button();
            this.textReadFile = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // buttonSave
            // 
            this.buttonSave.Location = new System.Drawing.Point(12, 68);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(100, 23);
            this.buttonSave.TabIndex = 1;
            this.buttonSave.Text = "Сохранить";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.ButtonSave_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // cboSheet
            // 
            this.cboSheet.AccessibleDescription = "";
            this.cboSheet.AccessibleName = "";
            this.cboSheet.AllowDrop = true;
            this.cboSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSheet.FormattingEnabled = true;
            this.cboSheet.Location = new System.Drawing.Point(12, 41);
            this.cboSheet.Name = "cboSheet";
            this.cboSheet.Size = new System.Drawing.Size(100, 21);
            this.cboSheet.TabIndex = 5;
            this.cboSheet.Tag = "";
            this.cboSheet.SelectedIndexChanged += new System.EventHandler(this.CboSheet_SelectedIndexChanged);
            // 
            // textPathFile
            // 
            this.textPathFile.Location = new System.Drawing.Point(122, 12);
            this.textPathFile.Name = "textPathFile";
            this.textPathFile.Size = new System.Drawing.Size(500, 23);
            this.textPathFile.TabIndex = 6;
            this.textPathFile.Text = "Путь файла";
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(12, 12);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(100, 23);
            this.btnOpen.TabIndex = 7;
            this.btnOpen.Text = "Открыть";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.BtnOpen_Click);
            // 
            // textReadFile
            // 
            this.textReadFile.Location = new System.Drawing.Point(122, 41);
            this.textReadFile.Name = "textReadFile";
            this.textReadFile.Size = new System.Drawing.Size(500, 308);
            this.textReadFile.TabIndex = 8;
            this.textReadFile.Text = "";
            // 
            // Excel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 361);
            this.Controls.Add(this.textReadFile);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.textPathFile);
            this.Controls.Add(this.cboSheet);
            this.Controls.Add(this.buttonSave);
            this.Name = "Excel";
            this.Text = "From xls to txt";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ComboBox cboSheet;
        private System.Windows.Forms.RichTextBox textPathFile;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.RichTextBox textReadFile;
    }
}

