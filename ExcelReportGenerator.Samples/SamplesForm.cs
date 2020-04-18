using ClosedXML.Excel;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Samples.Customizations;
using ExcelReportGenerator.Samples.Reports;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReportGenerator.Samples
{
    public partial class SamplesForm : Form
    {
        public SamplesForm()
        {
            InitializeComponent();
        }

        private void SamplesForm_Load(object sender, EventArgs e)
        {
            cmbReports.DataSource = typeof(ReportBase).Assembly.GetTypes()
                .Where(t => typeof(ReportBase).IsAssignableFrom(t) && !t.IsAbstract).ToArray();
            txtOutputFolder.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private async void btnRun_Click(object sender, EventArgs e)
        {
            var reportType = GetSelectedReport();
            var reportGenerator = GetReportGenerator(reportType);

            ToggleControlEnabled(true);

            try
            {
                await Task.Factory.StartNew(() =>
                {
                    var result = reportGenerator.Render(GetReportTemplateWorkbook(reportType));
                    result.SaveAs(Path.Combine(txtOutputFolder.Text, $"{reportType.Name}_Result.xlsx"));
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while running report: {ex.GetBaseException().Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ToggleControlEnabled(false);
        }

        private void ToggleControlEnabled(bool reportRunning)
        {
            btnRun.Enabled = !reportRunning;
            progressBar.Visible = reportRunning;
        }

        private static ReportBase GetReportInstance(Type reportType)
        {
            return (ReportBase)Activator.CreateInstance(reportType);
        }

        private Type GetSelectedReport()
        {
            return (Type)cmbReports.SelectedItem;
        }

        private static XLWorkbook GetReportTemplateWorkbook(MemberInfo reportType)
        {
            return new XLWorkbook(Path.Combine("Reports", "Templates", $"{reportType.Name}.xlsx"));
        }

        private static DefaultReportGenerator GetReportGenerator(Type reportType)
        {
            var report = GetReportInstance(reportType);
            return reportType == typeof(CustomReportGeneratorSample) ? new CustomReportGenerator(report) : new DefaultReportGenerator(report);
        }
    }
}