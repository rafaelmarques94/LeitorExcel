using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace HelperExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private DataSet ConexaoPlanilha(string caminho)
        {
            DataSet dsPlanilhaImportada = new DataSet();

            OleDbConnection conexao = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + caminho + ";Extended Properties='Excel 12.0;'");
            conexao.Open();

            DataTable dt = conexao.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            List<string> listaGuias = new List<string>();

            foreach (DataRow row in dt.Rows)
            {
                if (row["TABLE_NAME"].ToString().Contains("$"))
                {
                    listaGuias.Add(row["TABLE_NAME"].ToString());
                }
            }

            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + listaGuias[0] + "]", conexao);
            adapter.Fill(dsPlanilhaImportada);

            return dsPlanilhaImportada;
        }

        private void Bnt_UploadExcel_Click(object sender, EventArgs e)
        {
            // Limpa os itens do listBox
            listBox1.Items.Clear();

            OpenFileDialog ofdSelecionarArquivo = new OpenFileDialog();

            if (ofdSelecionarArquivo.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var filePath = ofdSelecionarArquivo.FileName;                
                    DataSet dsPlanilhaImportada = ConexaoPlanilha(filePath);

                    MessageBox.Show("A planilha foi importada com Sucesso!" + "\n" + "as celulas serão exibidas na lista abaixo", "Planilha Importada", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    for (int i = 0; i < dsPlanilhaImportada.Tables[0].Rows.Count; i++)
                    {
                        string ValorCelula = dsPlanilhaImportada.Tables[0].Rows[i][0].ToString();

                        listBox1.Items.Add(ValorCelula);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro.\n\nMensagem de Erro: {ex.Message}\n\n");

                }
            }
        }
    }
}
