using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static OfficeOpenXml.ExcelErrorValue;

namespace ScheduleGeneratorApp
{
    public partial class Form1 : Form
    {
        private List<Employee> employees = new List<Employee>();
        private DataTable oldDataTable;
        private DataTable newScheduleTable;

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Chọn file Employee.xlsx";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    LoadEmployees(ofd.FileName);
                }
            }

            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Chọn file LichTrucOld.xlsx";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    LoadOldSchedule(ofd.FileName);
                }
            }
        }

        private void LoadEmployees(string path)
        {
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var ws = package.Workbook.Worksheets["Sheet1"];
                int row = 2;
                while (!string.IsNullOrEmpty(ws.Cells[row, 1].Text))
                {
                    employees.Add(new Employee
                    {
                        Id = int.Parse(ws.Cells[row, 1].Text),
                        FullName = ws.Cells[row, 2].Text,
                        IdPosition = int.Parse(ws.Cells[row, 3].Text),
                        IdDepartment = int.Parse(ws.Cells[row, 4].Text),
                        NameOfDepartment = ws.Cells[row, 5].Text
                    });
                    row++;
                }
            }
        }

        private void LoadOldSchedule(string path)
        {
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var ws = package.Workbook.Worksheets["Sheet1"];
                oldDataTable = new DataTable();
                for (int col = 1; col <= 12; col++)
                {
                    oldDataTable.Columns.Add(ws.Cells[2, col].Text);
                }

                int row = 3;
                while (!string.IsNullOrEmpty(ws.Cells[row, 1].Text))
                {
                    var dataRow = oldDataTable.NewRow();
                    for (int col = 1; col <= 12; col++)
                    {
                        dataRow[col - 1] = ws.Cells[row, col].Text;
                    }
                    oldDataTable.Rows.Add(dataRow);
                    row++;
                }
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            // Lấy năm và tháng được chọn từ DateTimePicker
            int year = dtpMonthYear.Value.Year;
            int month = dtpMonthYear.Value.Month;

            // Tính số ngày của tháng đó (ví dụ: tháng 2 năm nhuận = 29 ngày)
            int days = DateTime.DaysInMonth(year, month);

            // Tổng số vị trí công việc cần phân công (giả định là 10 vị trí)
            int positions = 10;

            // Tạo DataTable mới để lưu kết quả phân công (clone từ bảng cũ)
            newScheduleTable = oldDataTable.Clone();

            // Tạo bản đồ ánh xạ giữa từng vị trí công việc và danh sách nhân viên tương ứng (hàng đợi - Queue)
            Dictionary<int, Queue<Employee>> posEmpMap = new Dictionary<int, Queue<Employee>>();
            foreach (var emp in employees)
            {
                if (!posEmpMap.ContainsKey(emp.IdPosition))
                    posEmpMap[emp.IdPosition] = new Queue<Employee>();

                // Đưa từng nhân viên vào hàng đợi theo vị trí công việc
                posEmpMap[emp.IdPosition].Enqueue(emp);
            }

            // Quy định giới hạn số lượng nhân viên của mỗi phòng ban được phân công trong 1 ngày
            Dictionary<int, int> departmentLimits = new Dictionary<int, int>
    {
        {1, 3}, {2, 2}, {3, 1}, {4, 2}, {5, 2}, {6, 2}, {7, 3}, {8, 1}
    };

            // Tạo dictionary lưu các phân công gần đây của mỗi nhân viên
            // Dùng để hạn chế việc phân công nhân viên lặp lại quá gần nhau
            Dictionary<int, Queue<int>> recentAssignments = new Dictionary<int, Queue<int>>();

            // Vòng lặp qua từng ngày trong tháng
            for (int i = 0; i < days; i++)
            {
                int newi = i + 1;

                // Tạo dòng mới tương ứng với ngày làm việc
                var row = newScheduleTable.NewRow();
                row[0] = newi; // Cột 'Day'

                // Tính ngày trong tuần (1 = Thứ 2, 7 = Chủ Nhật)
                DateTime dateValue = new DateTime(year, month, newi);
                int dateOfWeek = (int)dateValue.DayOfWeek + 1;
                row[1] = dateOfWeek == 1 ? "CN" : dateOfWeek.ToString(); // Cột 'DayOfWeek'

                // Tạo bản đồ đếm số nhân viên mỗi phòng ban được phân công hôm nay
                Dictionary<int, int> departmentCount = new Dictionary<int, int>();

                // Duyệt từng vị trí công việc (1 đến 10)
                for (int pos = 1; pos <= positions; pos++)
                {
                    int posId = pos;
                   
                    // Nếu không có nhân viên nào thuộc vị trí này → để trống
                    if (!posEmpMap.ContainsKey(posId) || posEmpMap[posId].Count == 0)
                    {
                        row[pos + 1] = "";
                        continue;
                    }

                    Employee chosen = null;
                    int maxTries = posEmpMap[posId].Count;

                    // Thử chọn nhân viên phù hợp từ hàng đợi
                    for (int t = 0; t < maxTries; t++)
                    {
                        var candidate = posEmpMap[posId].Dequeue();

                        // Khởi tạo danh sách lịch sử nếu chưa có
                        if (!recentAssignments.ContainsKey(candidate.Id))
                            recentAssignments[candidate.Id] = new Queue<int>();

                        var recent = recentAssignments[candidate.Id];

                        // 1 nhân viên trực ko đc chọn 2 ngày trong khoảng numberCheck 
                        int numberCheck = (employees.Count(e => e.IdPosition == posId) - 2);

                        // Giới hạn số lần phân công gần nhất được lưu để tránh trùng lặp
                        if (recent.Count >= numberCheck)
                            recent.Dequeue();

                        // Nếu nhân viên đã được phân công vào ngày này → bỏ qua
                        if (recent.Contains(i))
                        {
                            posEmpMap[posId].Enqueue(candidate);
                            continue;
                        }

                        // Nếu nhân viên đã được phân công vào ngày trong khoảng count - 2 từ tháng trước thì bỏ qua
                        if (!checkOldData(posId, numberCheck, newi, candidate.FullName))
                        {
                            posEmpMap[posId].Enqueue(candidate);
                            continue;
                        }

                        int deptId = candidate.IdDepartment;
                        departmentCount.TryGetValue(deptId, out int deptCount);

                        // Nếu phòng ban đã đủ số lượng nhân viên → bỏ qua nhân viên này
                        if (deptCount >= departmentLimits[deptId])
                        {
                            posEmpMap[posId].Enqueue(candidate);
                            continue;
                        }

                        // Nhân viên hợp lệ → chọn và cập nhật các danh sách
                        chosen = candidate;
                        departmentCount[deptId] = deptCount + 1;
                        recent.Enqueue(i);
                        break;
                    }

                    // Ghi tên nhân viên vào cột tương ứng của vị trí
                    row[pos + 1] = chosen?.FullName ?? "";

                    // Đưa lại nhân viên vào hàng đợi nếu đã được chọn (xoay vòng)
                    if (chosen != null)
                        posEmpMap[posId].Enqueue(chosen);
                }

                // Thêm dòng phân công vào bảng kết quả
                newScheduleTable.Rows.Add(row);
            }

            // Hiển thị bảng phân công ra DataGridView
            dataGridView1.DataSource = newScheduleTable;

            // Gọi hàm đổi màu hàng theo phòng ban
            ColorizeRowsByDepartment();
        }

        //ham nay check xem 1 người ở lịch mới ko xuất hiện trong numberCheck ngày ở cả lịch cũ
        private bool checkOldData(int posID, int numberCheck, int currentDay, string nameOfEmployee)
        {
            if (currentDay >= numberCheck)
            {
                return true;
            }
            // đầu tiên lấy numberCheck ngày cuối cùng của dataOld
            List<String> oldEmployeeListCheck = new List<String>();

            for (int index = oldDataTable.Rows.Count - numberCheck + 1; index < oldDataTable.Rows.Count; index++)
            {
                oldEmployeeListCheck.Add(oldDataTable.Rows[index][posID + 1].ToString());
            }

            for (int i = currentDay - 1; i < numberCheck - 1; i++)
            {
                if (oldEmployeeListCheck[i] == nameOfEmployee)
                {
                    return false;
                }
            }

            return true;
        }
        
        private void btnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Files|*.xlsx";
                sfd.Title = "Lưu lịch trực mới";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using (var package = new ExcelPackage())
                    {
                        var ws = package.Workbook.Worksheets.Add("NewSchedule");
                        for (int i = 0; i < newScheduleTable.Columns.Count; i++)
                            ws.Cells[1, i + 1].Value = newScheduleTable.Columns[i].ColumnName;

                        for (int r = 0; r < newScheduleTable.Rows.Count; r++)
                        {
                            for (int c = 0; c < newScheduleTable.Columns.Count; c++)
                            {
                                ws.Cells[r + 2, c + 1].Value = newScheduleTable.Rows[r][c];
                            }
                        }

                        package.SaveAs(new FileInfo(sfd.FileName));
                    }

                    MessageBox.Show("Xuất file thành công!");
                }
            }
        }

        private System.Drawing.Color GetColorForDepartment(int idDept)
        {
            return idDept switch
            {
                1 => System.Drawing.Color.Red,
                2 => System.Drawing.Color.Orange,
                3 => System.Drawing.Color.Green,
                4 => System.Drawing.Color.Blue,
                5 => System.Drawing.Color.Purple,
                6 => System.Drawing.Color.LightGreen,
                7 => System.Drawing.Color.Yellow,
                8 => System.Drawing.Color.Gray,
                _ => System.Drawing.Color.White
            };
        }
        private void ColorizeRowsByDepartment()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                //if (row.Cells["IdDepartment"].Value == null) continue;

                //int idDept = Convert.ToInt32(row.Cells["IdDepartment"].Value);
                //row.DefaultCellStyle.BackColor = GetColorForDepartment(idDept);
                for (int col = 2; col < row.Cells.Count; col++)
                {
                    if (row.Cells[col].Value == null) continue;
                    string value = row.Cells[col].Value.ToString();
                    foreach (var emp in employees)
                    {
                        if (emp.FullName == value)
                        {
                            row.Cells[col].Style.BackColor = GetColorForDepartment(emp.IdDepartment);
                            break;
                        }
                    }
                }
            }
        }

        private void ExportToExcelWithColors()
        {
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Employees");

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    var row = dataGridView1.Rows[i];
                    if (row.IsNewRow) continue;

                    int idDept = 0;
                    if (row.Cells["IdDepartment"].Value != null)
                    {
                        idDept = Convert.ToInt32(row.Cells["IdDepartment"].Value);
                    }

                    var bgColor = GetColorForDepartment(idDept);
                    var excelColor = System.Drawing.ColorTranslator.ToHtml(bgColor);

                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        var cell = worksheet.Cells[i + 2, j + 1];
                        cell.Value = row.Cells[j].Value;
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(bgColor);
                    }
                }

                var dialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Save an Excel File"
                };
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    package.SaveAs(new FileInfo(dialog.FileName));
                    MessageBox.Show("Export thành công!");
                }
            }
        }
        private void btnGenerateCalendar_Click(object sender, EventArgs e)
        {
            Dictionary<string, ScheduleModel> dicNgayDiLam = new Dictionary<string, ScheduleModel>();
            // Kiem tra so ngay
            for (int i = 0; i < newScheduleTable.Rows.Count; i++)
            {
                for (int j = 2; j < newScheduleTable.Columns.Count; j++)
                {
                    string fullName = newScheduleTable.Rows[i][j].ToString();
                    if (!dicNgayDiLam.ContainsKey(fullName))
                    {
                        ScheduleModel model = new ScheduleModel(0,0,0);
                        dicNgayDiLam[fullName] = model;
                    }
                    dicNgayDiLam[fullName].numberDay += 1;
                    if (newScheduleTable.Rows[i][1].ToString() == "7")
                    {
                        dicNgayDiLam[fullName].numberSaturDay += 1;
                    }
                    if (newScheduleTable.Rows[i][1].ToString() == "CN")
                    {
                        dicNgayDiLam[fullName].numberSunDay += 1;
                    }
                }
            }

            foreach (string key in dicNgayDiLam.Keys)
            {
                ScheduleModel model = dicNgayDiLam[key];
                Console.WriteLine("name " + key
                    + " \t\t" + model.numberDay
                    + " |\t" + model.numberSaturDay
                    + " |\t" + model.numberSunDay
                    );
            }
        }

    }

    public class Employee
    {
        public int Id { get; set; }
        public string FullName { get; set; }
        public int IdPosition { get; set; }
        public int IdDepartment { get; set; }
        public string NameOfDepartment { get; set; }
    }

    public class ScheduleModel
    {
        public int numberDay { get; set; }        
        public int numberSaturDay { get; set; }
        public int numberSunDay { get; set; }
        
        public ScheduleModel(int numberDay, int numberSaturday, int numberSunDay)
        {
            this.numberDay = numberDay;
            this.numberSaturDay = numberSaturday;
            this.numberSunDay = numberSunDay;
        }
    }
}






