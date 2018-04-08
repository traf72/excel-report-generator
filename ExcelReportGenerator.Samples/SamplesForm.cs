using ClosedXML.Excel;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Samples.Customizations;
using ExcelReportGenerator.Samples.Reports;
using System;
using System.IO;
using System.Linq;
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
            Type reportType = GetSelectedReport();
            DefaultReportGenerator reportGenerator = GetReportGenerator(reportType);

            ToggleControlEnabled(true);

            try
            {
                await Task.Factory.StartNew(() =>
                {
                    XLWorkbook result = reportGenerator.Render(GetReportTemplateWorkbook(reportType));
                    result.SaveAs(Path.Combine(txtOutputFolder.Text, string.Format("{0}_Result.xlsx", reportType.Name)));
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("An error occurred while running report: {0}", ex.GetBaseException().Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ToggleControlEnabled(false);
        }

        private void ToggleControlEnabled(bool reportRunning)
        {
            btnRun.Enabled = !reportRunning;
            progressBar.Visible = reportRunning;
        }

        private ReportBase GetReportInstance(Type reportType)
        {
            return (ReportBase)Activator.CreateInstance(reportType);
        }

        private Type GetSelectedReport()
        {
            return (Type)cmbReports.SelectedItem;
        }

        private XLWorkbook GetReportTemplateWorkbook(Type reportType)
        {
            return new XLWorkbook(Path.Combine("Reports", "Templates", string.Format("{0}.xlsx", reportType.Name)));
        }

        private DefaultReportGenerator GetReportGenerator(Type reportType)
        {
            ReportBase report = GetReportInstance(reportType);
            return reportType == typeof(CustomReportGeneratorSample) ? new CustomReportGenerator(report) : new DefaultReportGenerator(report);
        }
    }
}