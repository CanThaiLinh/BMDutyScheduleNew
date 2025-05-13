namespace ScheduleGeneratorApp
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnLoadData;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DateTimePicker dtpMonthYear;
        private System.Windows.Forms.Button btnGenerateCalendar;

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
            btnLoadData = new System.Windows.Forms.Button();
            btnGenerate = new System.Windows.Forms.Button();
            btnExport = new System.Windows.Forms.Button();
            dataGridView1 = new System.Windows.Forms.DataGridView();
            dtpMonthYear = new System.Windows.Forms.DateTimePicker();
            btnGenerateCalendar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // btnLoadData
            // 
            btnLoadData.Location = new System.Drawing.Point(12, 12);
            btnLoadData.Name = "btnLoadData";
            btnLoadData.Size = new System.Drawing.Size(100, 30);
            btnLoadData.TabIndex = 0;
            btnLoadData.Text = "Load Excel";
            btnLoadData.Click += btnLoadData_Click;
            // 
            // btnGenerate
            // 
            btnGenerate.Location = new System.Drawing.Point(743, 12);
            btnGenerate.Name = "btnGenerate";
            btnGenerate.Size = new System.Drawing.Size(100, 30);
            btnGenerate.TabIndex = 1;
            btnGenerate.Text = "Generate";
            btnGenerate.Click += btnGenerate_Click;
            // 
            // btnExport
            // 
            btnExport.Location = new System.Drawing.Point(889, 12);
            btnExport.Name = "btnExport";
            btnExport.Size = new System.Drawing.Size(100, 30);
            btnExport.TabIndex = 2;
            btnExport.Text = "Export";
            btnExport.Click += btnExport_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.Location = new System.Drawing.Point(12, 50);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new System.Drawing.Size(1000, 800);
            dataGridView1.TabIndex = 3;
            // 
            // dtpMonthYear
            // 
            dtpMonthYear.CustomFormat = "MM/yyyy";
            dtpMonthYear.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            dtpMonthYear.Location = new System.Drawing.Point(147, 14);
            dtpMonthYear.Name = "dtpMonthYear";
            dtpMonthYear.ShowUpDown = true;
            dtpMonthYear.Size = new System.Drawing.Size(150, 23);
            dtpMonthYear.TabIndex = 4;
            // 
            // btnGenerateCalendar
            // 
            btnGenerateCalendar.Location = new System.Drawing.Point(583, 12);
            btnGenerateCalendar.Name = "btnGenerateCalendar";
            btnGenerateCalendar.Size = new System.Drawing.Size(120, 30);
            btnGenerateCalendar.TabIndex = 5;
            btnGenerateCalendar.Text = "Kiểm tra";
            btnGenerateCalendar.UseVisualStyleBackColor = true;
            btnGenerateCalendar.Click += btnGenerateCalendar_Click;
            // 
            // Form1
            // 
            ClientSize = new System.Drawing.Size(1024, 861);
            Controls.Add(btnLoadData);
            Controls.Add(btnGenerate);
            Controls.Add(btnExport);
            Controls.Add(dataGridView1);
            Controls.Add(dtpMonthYear);
            Controls.Add(btnGenerateCalendar);
            Name = "Form1";
            Text = "Lịch trực auto";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
        }
    }
}