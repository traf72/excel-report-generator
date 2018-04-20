using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace ExcelReportGenerator.LicenseManager
{
    public partial class MainForm : Form
    {
        private const int LicenseExpirationDateByteNumber = 217;
        private const string EncryptionKey = "lColuccimTNERPEULLPARSIstanRTTAtalpmednotfoertyvcnuFecxEelblttempplatrPecnatMethodCtPropertdataittcnuElenaPcERTEMtemPanelEvePLATEimanyDecrFmetemvalueptyValuePallValuePtsnItleProclateproairaVditageralueprundexctllacdEULAVceProvNGSETAVMETIVELEanydlmnAttr";
        private const string LicenseFileName = "ExcelReportGenerator.lic";

        public MainForm()
        {
            InitializeComponent();
            dtpExpirationDate.Value = DateTime.Now.AddYears(2);
            //dtpExpirationDate.Value = new DateTime(2200, 1, 1);
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            var rnd = new Random((int)DateTime.Now.Ticks & 0x0000FFFF);
            byte[] fictitiousBytes = new byte[369];
            rnd.NextBytes(fictitiousBytes);

            DateTime expirationDate = EndOfDay(dtpExpirationDate.Value);
            byte[] ticksBytes = BitConverter.GetBytes(expirationDate.Ticks);

            byte[] payload = fictitiousBytes.Take(LicenseExpirationDateByteNumber).Concat(ticksBytes).Concat(fictitiousBytes.Skip(LicenseExpirationDateByteNumber).Take(fictitiousBytes.Length)).ToArray();
            byte[] hash;
            using (MD5 hashAlg = MD5.Create())
            {
                hash = hashAlg.ComputeHash(payload);
            }

            byte[] allBytes = payload.Concat(hash).ToArray();
            byte[] encryptedBytes = Encryptor.Encrypt(allBytes, EncryptionKey);
            File.WriteAllBytes(LicenseFileName, encryptedBytes);

            MessageBox.Show(this, "Success", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private DateTime EndOfDay(DateTime date)
        {
            return date.Date.AddDays(1).AddTicks(-1);
        }
    }
}