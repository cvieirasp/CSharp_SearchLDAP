using System;
using System.DirectoryServices;

namespace SearchLDAP
{
    class Program
    {
        private const int TOTAL_REG = 5;
        private static DirectoryEntry rootEntry = null;

        static void Main(string[] args)
        {
            char exit = 'N';
            while(exit != 'S')
            {
                Console.WriteLine("#####################");
                Console.WriteLine("### PESQUISA LDAP ###");
                Console.WriteLine("#####################\n");

                Console.WriteLine("[U] Pesquisar por Usuário");
                Console.WriteLine("[G] Pesquisar por Grupo");
                Console.WriteLine("[S] Sair");

                Console.WriteLine("Selecione uma das opções.\n");
                ConsoleKeyInfo keyPressed = Console.ReadKey(true);

                string parameter = string.Empty;
                SelectDirectory();

                switch (keyPressed.Key)
                {
                    case ConsoleKey.U:

                        do
                        {
                            Console.WriteLine("Insira o nome do usuário: ");
                            parameter = Console.ReadLine();
                            if (string.IsNullOrWhiteSpace(parameter))
                            {
                                Console.WriteLine("Nome inválido!\n");
                            }
                            else
                            {
                                SearchUser(parameter);
                                Console.WriteLine("Deseja realizar nova pesquisa?");
                                Console.WriteLine("(Clique 'S' para continuar ou qualquer outra tecla para cancelar)\n");
                                keyPressed = Console.ReadKey(true);
                                if (keyPressed.Key == ConsoleKey.S)
                                    parameter = string.Empty;
                            }

                        } while (string.IsNullOrWhiteSpace(parameter));

                        break;
                    case ConsoleKey.G:

                        do
                        {
                            Console.WriteLine("Insira o nome do grupo de usuário: ");
                            parameter = Console.ReadLine();
                            if (string.IsNullOrWhiteSpace(parameter))
                            {
                                Console.WriteLine("Nome inválido!\n");
                            }
                            else
                            {
                                SearchGroup(parameter);
                                Console.WriteLine("Deseja realizar nova pesquisa?");
                                Console.WriteLine("(Clique 'S' para continuar ou qualquer outra tecla para cancelar)\n");
                                keyPressed = Console.ReadKey(true);
                                if (keyPressed.Key == ConsoleKey.S)
                                    parameter = string.Empty;
                            }

                        } while (string.IsNullOrWhiteSpace(parameter));

                        break;
                    case ConsoleKey.S:
                        exit = 'S';
                        break;
                      default:
                        break;
                }
            }

            Console.WriteLine("\nFIM!!!");
            ExitSearch();
            Console.ReadKey();
        }

        private static void SelectDirectory()
        {
            while (rootEntry == null)
            {
                // LDAP://empresa.com.br
                //Console.WriteLine("Insira o diretório LDAP: ");
                //string directory = Console.ReadLine();
                string directory = "LDAP://unimedbh.com.br";
                try
                {
                    rootEntry = new DirectoryEntry(directory);
                    rootEntry.AuthenticationType = AuthenticationTypes.None;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Console.WriteLine("Erro ao conectar no diretório LDAP.\nDeseja finalizar?");
                    Console.WriteLine("(Clique 'S' para finalizar ou qualquer outra tecla para continuar)");
                    rootEntry = null;

                    if (Console.ReadKey(true).Key == ConsoleKey.S)
                        break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ocorreu um erro inesperado.");
                    Console.WriteLine("({0})\n", ex.Message);
                    ConsoleKeyInfo keyPressed = Console.ReadKey(true);
                    rootEntry = null;
                }
            }
        }

        private static void SearchUser(string user)
        {
            int count = 0;
            using (DirectorySearcher searcher = new DirectorySearcher(rootEntry))
            {
                var queryFormat = "(&(objectClass=user)(objectCategory=person)(cn={0}*)(!userAccountControl:1.2.840.113556.1.4.803:=2))";
                searcher.Filter = string.Format(queryFormat, user);
                searcher.SizeLimit = TOTAL_REG;
                Console.WriteLine("--------------------");
                foreach (SearchResult result in searcher.FindAll())
                {
                    Console.WriteLine("Nome: {0}", result.Properties["cn"].Count > 0 ? result.Properties["cn"][0] : string.Empty);
                    Console.WriteLine("Login: {0}", result.Properties["samAccountName"].Count > 0 ? result.Properties["samAccountName"][0] : string.Empty);
                    Console.WriteLine("Email: {0}", result.Properties["mail"].Count > 0 ? result.Properties["mail"][0] : string.Empty);
                    Console.WriteLine("--------------------");
                    count++;
                }
            }
            Console.WriteLine("Usuários encontrados: {0}\n", count);
        }

        private static void SearchGroup(string group)
        {
            int count = 0;
            using (DirectorySearcher searcher = new DirectorySearcher(rootEntry))
            {
                var queryFormat = "(&(objectClass=group)(cn={0}*))";
                searcher.Filter = string.Format(queryFormat, group);
                searcher.SizeLimit = TOTAL_REG;
                Console.WriteLine("--------------------");
                foreach (SearchResult result in searcher.FindAll())
                {
                    Console.WriteLine("Nome: {0}", result.Properties["cn"].Count > 0 ? result.Properties["cn"][0] : string.Empty);
                    Console.WriteLine("--------------------");
                    count++;
                }
            }
            Console.WriteLine("Grupos encontrados: {0}\n", count);
        }

        private static void ExitSearch()
        {
            if (rootEntry != null)
            {
                rootEntry.Dispose();
                rootEntry = null;
            }
        }
    }
}
