using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Planilhas
{
    public class PlanilhaHandler
    {
        public PlanilhaHandler()
        {

        }
        public void CriarPlanilha(string caminhoPlanilha)
        {
            try
            {
                var Vendas = new[]
                {
                new {Id = "Dados", Filial = "01", Vendas = 1},
                new {Id = "Dados", Filial = "02", Vendas = 2},
                new {Id = "Dados", Filial = "03", Vendas = 3},
                new {Id = "Dados", Filial = "04", Vendas = 4},
                new {Id = "Dados", Filial = "05", Vendas = 5},
                new {Id = "Dados", Filial = "06", Vendas = 6},
                new {Id = "Dados", Filial = "07", Vendas = 7}
            };

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excel = new ExcelPackage())
                {
                    var workSheet = excel.Workbook.Worksheets.Add("Planilha Vendas");
                    workSheet.TabColor = Color.Black;
                    workSheet.DefaultRowHeight = 12;

                    workSheet.Row(1).Height = 20;
                    workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Row(1).Style.Font.Bold = true;

                    workSheet.Cells[1, 1].Value = "Cod.";
                    workSheet.Cells[1, 2].Value = "Filial";
                    workSheet.Cells[1, 3].Value = "Vendas/mil";
                    workSheet.Cells["A1:C1"].Style.Font.Italic = true;

                    int indice = 2;
                    foreach (var venda in Vendas)
                    {
                        workSheet.Cells[indice, 1].Value = venda.Id;
                        workSheet.Cells[indice, 2].Value = venda.Filial;
                        workSheet.Cells[indice, 3].Value = venda.Vendas;
                        indice++;
                    }

                    workSheet.Column(1).AutoFit();
                    workSheet.Column(2).AutoFit();
                    workSheet.Column(3).AutoFit();

                    string arquivoPlanilha = $"{caminhoPlanilha}\\PlanilhaVendas.xls";
                    if (File.Exists(arquivoPlanilha))
                        File.Delete(arquivoPlanilha);

                    File.WriteAllBytes(arquivoPlanilha, excel.GetAsByteArray());
                }

                Console.WriteLine($"Planilha Planilha.xls criada com sucesso em: {caminhoPlanilha}!\n");
            }
            catch (Exception ex)
            {
                var mensagem = $"Caminho informado {caminhoPlanilha} sem permissao de acesso. Informe outro caminho para criar a planilha!\n";
                Console.WriteLine(mensagem);
            }
        }

        public void AbrirPlanilha(string caminhoPlanilha)
        {
            try
            {
                using (var arquivoExcel = new ExcelPackage(new FileInfo(caminhoPlanilha)))
                {
                    ExcelWorksheet planilhaVendas = arquivoExcel.Workbook.Worksheets.FirstOrDefault();
                    int linhas = planilhaVendas.Dimension.Rows;
                    int colunas = planilhaVendas.Dimension.Columns;

                    for (int i = 1; i <= linhas; i++)
                    {
                        for (int j = 1; j <= colunas; j++)
                        {
                            string conteudo = planilhaVendas.Cells[i, j].Value?.ToString();
                            Console.WriteLine(conteudo);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var mensagem = $"Erro ao tentar criar o arquivo: {caminhoPlanilha}!\n";
                Console.WriteLine(mensagem);
            }
        }
    }
}
