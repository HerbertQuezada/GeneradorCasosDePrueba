using System.Data;

namespace CasosDePruebaECF
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void xlsx_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel Files |*.xlsx";
                openFile.Title = "Seleccione el archivo de Excel";
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    if (openFile.FileName.Equals("") == false)
                    {
                        path.Text = openFile.FileName;
                        Services.dt = Services.ConvertExcelToDataTable(openFile.FileName);
                        casos.DataSource = new BindingSource(Services.CasoDePrueba, null);
                        casos.DisplayMember = "Value";
                        casos.ValueMember = "Key";
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void casos_SelectedIndexChanged(object sender, EventArgs e)
        {
            documento.Text = Services.loadJson(casos.SelectedIndex);
        }

        private void copy_Click(object sender, EventArgs e) => Clipboard.SetText(documento.Text);
    }
}
