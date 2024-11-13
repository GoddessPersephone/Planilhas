namespace Planilhas
{
    public class Program
    {
        private static void Main(string[] args)
        {
            // Solicita ao usuário o caminho da pasta para salvar a planilha
            Console.WriteLine("Por favor, selecione o diretório para salvar a planilha.");

            // Instância da classe PlanilhaHandler
            PlanilhaHandler planilhaHandler = new PlanilhaHandler();

            // Solicita o caminho da planilha do usuário
            Console.WriteLine("Digite o caminho da pasta onde deseja salvar a planilha:");
            string caminhoPlanilha = Console.ReadLine();

            // Cria e abre a planilha
            planilhaHandler.CriarPlanilha(caminhoPlanilha);
            planilhaHandler.AbrirPlanilha($"{caminhoPlanilha}\\Planilha.xls");

            Console.WriteLine("Pressione qualquer tecla para encerrar.");
            Console.ReadKey();
        }
    }
}
