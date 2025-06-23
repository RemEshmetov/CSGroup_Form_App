using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSGroup_Form_App
{
    public partial class Form1 : Form
    {
        private Panel panel;
        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            
            this.Text = "CS Group";
            this.BackColor = Color.White;
            this.Size = new Size(500, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.CenterToScreen();
            this.Resize += MainForm_Resize;
            
            Panel panel = new Panel
            {
                BackColor = Color.White,
                
                Size = new Size(450, 500),
                Dock = DockStyle.Fill

            };

            panel.Location = new Point(
            (this.ClientSize.Width - panel.Width) / 2,
            (this.ClientSize.Height - panel.Height) / 2
        );
            





            
            Label titleLabel = new Label
            {
                Text = "CS Group",
                Font = new Font("Arial", 16, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 57, 18),
                Location = new Point(180, 20),
                AutoSize = true
            };
            
            
            string[] labels = { "Фамилия Имя Отчество", "Название организации", "Номер телефона", "Почта", "Должность", "Что вас интересует?" };

            textBoxName = CreateTextBoxAndLabel(panel, labels[0], 70);
            textBoxOrganization = CreateTextBoxAndLabel(panel, labels[1], 120);
            textBoxPhone = CreateTextBoxAndLabel(panel, labels[2], 170);
            textBoxEmail = CreateTextBoxAndLabel(panel, labels[3], 220);
            textBoxPosition = CreateTextBoxAndLabel(panel, labels[4], 270);
            textBoxMessage = CreateTextBoxAndLabel(panel, labels[5], 320, 60); 
           
            Button buttonSend = new Button
            {
                Text = "Отправить",
                Location = new Point(50, 420), 
                Width = 350,
                Height = 40,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Font = new Font("Arial", 10, FontStyle.Bold)
            };

            buttonSend.Click += ButtonSend_Click;

            panel.Controls.Add(titleLabel);
            panel.Controls.Add(buttonSend);
            this.Controls.Add(panel);
        }
        private void MainForm_Resize(object sender, EventArgs e)
        {
            
            if (panel != null)
            {
                panel.Location = new Point(
                    (this.ClientSize.Width - panel.Width) / 2,
                    (this.ClientSize.Height - panel.Height) / 2
                );
            }
        }

        private TextBox CreateTextBoxAndLabel(Panel panel, string labelText, int yPosition, int height = 30)
        {
            Label label = new Label
            {
                Text = labelText,
                Location = new Point(50, yPosition),
                AutoSize = true,
                Font = new Font("Arial", 10, FontStyle.Regular)
            };

            TextBox textBox = new TextBox
            {
                Location = new Point(50, yPosition + 20),
                Width = 350,
                Height = height,
                Font = new Font("Arial", 10)
            };

            panel.Controls.Add(label);
            panel.Controls.Add(textBox);

            return textBox;
        }


        private void ButtonSend_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(textBoxName.Text) ||
                string.IsNullOrEmpty(textBoxOrganization.Text) ||
                string.IsNullOrEmpty(textBoxPhone.Text) ||
                string.IsNullOrEmpty(textBoxEmail.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все обязательные поля!");
                return;
            }

            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Path.Combine(documentsPath, "CS Clients.xlsx");
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();

                
                if (!File.Exists(filePath))
                {
                    workbook = excelApp.Workbooks.Add();
                    worksheet = workbook.Sheets[1];

                    // Заголовки столбцов
                    worksheet.Cells[1, 1].Value = "ФИО";
                    worksheet.Cells[1, 2].Value = "Название организации";
                    worksheet.Cells[1, 3].Value = "Номер телефона";
                    worksheet.Cells[1, 4].Value = "Почта";
                    worksheet.Cells[1, 5].Value = "Должность";
                    worksheet.Cells[1, 6].Value = "Что вас интересует?";
                }
                else
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                    worksheet = workbook.Sheets[1];
                }

                
                int row = worksheet.Cells.Find(
                    What: "*",
                    After: worksheet.Cells[1, 1],
                    LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlPrevious
                ).Row + 1;
               
                worksheet.Cells[row, 1].Value = textBoxName.Text;
                worksheet.Cells[row, 2].Value = textBoxOrganization.Text;
                worksheet.Cells[row, 3].Value = textBoxPhone.Text;
                worksheet.Cells[row, 4].Value = textBoxEmail.Text;
                worksheet.Cells[row, 5].Value = textBoxPosition.Text;
                worksheet.Cells[row, 6].Value = textBoxMessage.Text;

                
                worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 6]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 6]].Interior.Color = Color.LightYellow;

                textBoxName.Clear();
                textBoxOrganization.Clear();
                textBoxPhone.Clear();
                textBoxEmail.Clear();
                textBoxPosition.Clear();
                textBoxMessage.Clear();
                
                workbook.Save();
                MessageBox.Show("Данные сохранены успешно!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}");
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }

                
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                GC.Collect();
            }
        }

        
    }
}

