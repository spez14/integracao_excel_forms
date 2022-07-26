using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace IntegracaoExcel
{
    public partial class Form1 : Form
    {
        Excel.Application app = new Excel.Application();
        Workbook pasta;
        Worksheet plan;
        string path = @"c:\dados\resultado.xlsx";

        public Form1()
        {
            InitializeComponent();
            CarregarPlanilha();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void btnAbrir_Click(object sender, EventArgs e)
        {
            CarregarPlanilha();
        }

        private void CarregarPlanilha()
        {
            lblStatus.Text = "Abrindo planilha de resultado";

            try
            {
                pasta = app.Workbooks.Open(path);
                plan = pasta.Worksheets["Plan1"];

                lblStatus.Text = "Caregando receitas";
                // ------ Receitas ------------------
                txtFaturamento.Text = plan.Cells[5, 3].Value.ToString("N2");
                lblDevolucoes.Text = plan.Cells[6, 3].Value.ToString("N2");
                lblTotalReceitas.Text = plan.Cells[7, 3].Value.ToString("N2");

                lblStatus.Text = "Carregando despesas";
                //------ Despesas ---------------------
                lblComissoes.Text = plan.Cells[10, 3].Value.ToString("N2");
                lblCustosProdutos.Text = plan.Cells[11, 3].Value.ToString("N2");
                lblImpostos.Text = plan.Cells[12, 3].Value.ToString("N2");
                lblDespesasAdministrativas.Text = plan.Cells[13, 3].Value.ToString("N2");
                lblTotalDespesas.Text = plan.Cells[14, 3].Value.ToString("N2");

                lblStatus.Text = "Carregando resultado";
                //------ Resultado ---------------------
                lblResultado.Text = Convert.ToDecimal(plan.Cells[16, 3].Value).ToString("N2");

                if (pasta.ReadOnly)
                {
                    txtFaturamento.Enabled = false;
                    btnSalvar.Enabled = false;
                    lblStatus.Text = "Pronto, somente para leitura";
                }
                else
                {
                    txtFaturamento.Enabled = true;
                    btnSalvar.Enabled = true;
                    txtFaturamento.Focus();
                    lblStatus.Text = "Pronto";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha: " + ex.Message);
            }

           
        }

        private void lblResultado_Click(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (pasta != null)
                pasta.Close(true);

            app.Quit();

            plan = null;
            pasta = null;
            app = null;
            
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            SalvarPlanilha();
        }

        private void SalvarPlanilha()
        {
            lblStatus.Text = "Salvando a planilha de resultado, aguarde...";

            try
            {
                //------ Receitas ---------------------
                plan.Cells[5, 3].Value = Convert.ToDecimal("0" + txtFaturamento.Text);
                pasta.Save();

                MessageBox.Show("A Planilha foi salva!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha: " + ex.Message);
            }

            lblStatus.Text = "Pronto";
        }
    }
}