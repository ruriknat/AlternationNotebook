//using Microsoft.Win32;
//using Preactor;
//using Preactor.Interop.PreactorObject;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Runtime.InteropServices;
//using System.Security.Cryptography;
//using System.Xml.Linq;

//namespace NativeRules
//{
//    [Guid("01215791-0677-4f55-9803-635ab5694427")]
//    [ComVisible(true)]
//    public interface IAlternationOperation
//    {
//        int UnallocateAlternationOperation(ref PreactorObj preactorComObject, ref object pespComObject);
//        int SelectedResources(ref PreactorObj preactorComObject, ref object pespComObject);
//    }

//    [ComVisible(true)]
//    [ClassInterface(ClassInterfaceType.None)]
//    [Guid("3d2e0d4e-f45e-46da-ba1e-2568920ee92f")]
//    public class AlternationOperation : IAlternationOperation
//    {
//        public int UnallocateAlternationOperation(ref PreactorObj preactorComObject, ref object pespComObject)
//        {
//            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

//            // ok // 1 - Sequenciar operações para frente
//            // nk // // 2 - Sequenciar operações SOLDAR ROBO
//            // -- // ok // 2.1 - Selecioanr Operações SOLDAR ROBO 
//            // -- // ok // 2.2 - Desprogramar as operações SOLDAR ROBO e operações subsequentes
//            // -- // nk // 2.3 - Programar operações SOLDAR ROBO de forma alternada
//            // -- // nk // 2.4 - Corrigir tempo de operação para que seja possivel realizar 1 operacao de forma alternada no recurso Robo N
//            // -- // nk // 2.5 - Consolidar/Agrupar operações SOLDAR ROBO de uma mesma ordem
//            // nk // 3 - Sequenciar as operações posteriores a solda robo

//            // Abre as listas de Ordens e Recursos
//            IList<Orders> listaOrders = new List<Orders>();
//            IList<Resources> listaRecursos = new List<Resources>();

//            GetOrders(preactor, listaOrders);

//            GetResources(preactor, listaRecursos);

//            // Sequencia todas as operações para frente
//            preactor.PlanningBoard.SequenceAll(SequenceAllDirection.Forwards, SequencePriority.DueDate);

//            // Filtrando as operacoes de "SOLDAR ROBO" e ordenando por SetupStart
//            var listaOrdemSoldarRobo = listaOrders
//                .Where(s => s.OperationName == "SOLDAR ROBO")
//                .ToList();

//            // Iterar sobre as operacoes de "SOLDAR ROBO"
//            foreach (var registoOrdem in listaOrdemSoldarRobo)
//            {
//                // busca a operacao anterior para que seja desprograma todas as operacoes subsequentes
//                int previousRecord = preactor.PlanningBoard.GetPreviousOperation(registoOrdem.Record, 1);

//                // Verificacao se existe uma operacao subsequente (se (PreviusRecord < 0 nao existe operacao antecessora)
//                if (previousRecord > 0)
//                {
//                    // Desprograma as operações subsequentes
//                    preactor.PlanningBoard.UnallocateOperation(previousRecord, OperationSelection.SubsequentOperations);
//                    // Desprograma a operação de SOLDAR ROBO
//                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.ThisOperation);
//                }
//                else
//                {
//                    // Se não houver operação anterior, desprogramar a operação atual
//                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.BiDirectionalOperations);

//                    // Desprograma a operação de SOLDAR ROBO
//                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.ThisOperation);
//                }
//            }

//            // --------------------------------------------- Excluir daqui

//            return 0;

//        }

//        public int SelectedResources(ref PreactorObj preactorComObject, ref object pespComObject)
//        {
//            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

//            // Lista em que sera salvo os dados das ordens/operações
//            IList<Orders> listaOrders = new List<Orders>();

//            GetOrders(preactor, listaOrders);


//            // == Dados Recursos ==

//            // Lista em que será salvo os dados dos recursos
//            IList<Resources> listaRecursos = new List<Resources>();

//            GetResources(preactor, listaRecursos);

//            // Filtrando as operacoes de "SOLDAR ROBO" e ordenando por SetupStart
//            var listaOrdemSoldarRobo = listaOrders
//                .Where(s => s.OperationName == "SOLDAR ROBO")
//                .OrderBy(x => x.DueDate)
//                .ToList();

//            // --------------------------------------------- Até aqui excluir daqui



//            // Cria uma lista com os dados filtrando os recursos de solda robo
//            var listaRecursoRoboSolda = listaRecursos
//                .Where(r => r.ResourceName != null &&
//                            r.ResourceName.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0 &&
//                            r.Attribute4.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0)
//                .OrderBy(r => r.ResourceName)
//                .ToList();

//            // Cria uma lista com os agrupamentos de recursos de solda robo
//            var listaGrupoRecursoRoboSolda = listaRecursoRoboSolda
//                .GroupBy(r => r.Attribute4)
//                .Select(s => new { Attribute4 = s.Key })
//                .ToList();


//            // == Dados Recursos ==

//            // == Inicio Dados Ordens Soldar Robo ==


//            // Agrupa as ordens pelo campo OrderNo
//            var listaOrdensAgrupadasPorOrderNo = listaOrdemSoldarRobo
//                .GroupBy(o => o.OrderNo)
//                .ToList();

//            // Cria um valor de ordenação alternado dentro de cada grupo de ordens
//            int valorOrdenacaoCounter = 0;
//            foreach (var grupo in listaOrdensAgrupadasPorOrderNo)
//            {
//                // Ordena as ordens dentro do grupo por DueDate
//                var listaOrdensDoGrupoOrdenadas = grupo.OrderBy(o => o.DueDate).ToList();

//                // Atribui um valor alternado de ordenação para as ordens dentro do grupo
//                foreach (var ordem in listaOrdensDoGrupoOrdenadas)
//                {
//                    // Atualiza o valor de ordenação na lista original para todos os registros com o mesmo OrderNo
//                    var ordensParaAtualizar = listaOrdemSoldarRobo
//                        .Where(o => o.OrderNo == ordem.OrderNo)
//                        .ToList();

//                    foreach (var ordemOriginal in ordensParaAtualizar)
//                    {
//                        ordemOriginal.ValorOrdenacao = valorOrdenacaoCounter;
//                    }
//                    valorOrdenacaoCounter++; // Alterna entre 1 e 2
//                }
//            }

//            // Ordena a lista final com base no novo valor de ordenação
//            var ListaOrdemSoldarRoboOrdenada = listaOrdemSoldarRobo
//                .OrderBy(x => x.OrdenacaoPeca)  // Depois ordena por OrdenacaoPeca
//                .ThenBy(x => x.ValorOrdenacao) // Ordena pelo valor sequêncial das ordens
//                .ThenBy(x => x.DueDate)       // Por fim, ordena por DueDate
//                .ToList();


//            // == Fim Dados Ordens Soldar Robo ==

//            // == Interações Solda Robo ==

//            var queueOrdens = new Queue<Orders>(ListaOrdemSoldarRoboOrdenada);

//            DateTime dimc = preactor.PlanningBoard.TerminatorTime;

//            // Lista para armazenar os recursos que já foram programados
//            List<int> recursosProgramados = new List<int>();

//            // Tabela para armazenar os recursos e as ordens alocadas
//            List<(int recursoId, string OrderNo, int ordemId, DateTime changeStart)> tabelaOrdensRecurso = new List<(int, string, int, DateTime)>();

//            while (recursosProgramados.Count < listaRecursoRoboSolda.Count)
//            {
//                // Seleciona a primeira ordem
//                var primeiraOrdem = queueOrdens.Dequeue();


//                // Lista para armazenar os resultados dos testes
//                List<(int recursoId, DateTime changeStart)> listaResultadosTestes = new List<(int recursoId, DateTime changeStart)>();

//                // Itera sobre os recursos disponíveis (excluindo os já programados)
//                foreach (var recurso in listaRecursoRoboSolda.Where(r => !recursosProgramados.Contains(r.ResourceId)))
//                {
//                    // Realiza o teste de operação para o recurso atual e a ordem
//                    var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(primeiraOrdem.Record, recurso.ResourceId, dimc);

//                    if (resultadoTeste.HasValue)
//                    {
//                        // Armazena o recurso e o tempo de início do teste
//                        listaResultadosTestes.Add((recurso.ResourceId, resultadoTeste.Value.ChangeStart));
//                    }
//                }

//                // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
//                if (listaResultadosTestes.Count > 0)
//                {
//                    // Ordena os resultados pelo tempo de início (ChangeStart)
//                    var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

//                    // Verifica as condições para o PutOp
//                    if (primeiraOrdem.RecursoRequerido == -1)
//                    {
//                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
//                        preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);

//                        // Marca a ordem como programada
//                        primeiraOrdem.Programada = true;

//                        // Adiciona o recurso à lista de recursos programados
//                        recursosProgramados.Add(resultadoMinimo.recursoId);

//                        // Registra a ordem e o recurso na tabela
//                        tabelaOrdensRecurso.Add((resultadoMinimo.recursoId, primeiraOrdem.OrderNo, primeiraOrdem.Record, resultadoMinimo.changeStart)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
//                    }
//                }
//                else if (listaResultadosTestes.Count <= 0)
//                {
//                    queueOrdens.Enqueue(primeiraOrdem);
//                }
//            }

//            // Lista para armazenar as operações sequenciadas
//            List<(int ordemId, string OrderNo, int recursoId, DateTime startTime, DateTime endTime)> sequenciamentoOperacoes = new List<(int, string, int, DateTime, DateTime)>();

//            // Considerando que você tenha uma lista de todas as operações por ordem
//            foreach (var ordem in ListaOrdemSoldarRoboOrdenada)
//            {
//                // Filtra as ordens que já estão programadas
//                var ordensComRecursoAlocado = tabelaOrdensRecurso.Where(t => t.OrderNo == ordem.OrderNo).OrderBy(t => t.changeStart).ToList();

//                foreach (var item in ordensComRecursoAlocado)
//                {
//                    if (ordem.OrdenacaoPeca == 1)  // Se a ordem for a primeira, pega o tempo de início do recurso alocado
//                    {
//                        // Definindo o tempo de início e fim
//                        DateTime tempoInicio = item.changeStart;
//                        DateTime tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", ordem.Record);

//                        // Adiciona ao sequenciamento
//                        sequenciamentoOperacoes.Add((ordem.Record, ordem.OrderNo, item.recursoId, tempoInicio, tempoFim));

//                        // Atualiza o dimc para o tempo de fim da operação
//                        dimc = tempoFim;
//                    }
//                    else
//                    {
//                        // Para ordens subsequentes, usa o dimc do ciclo anterior
//                        DateTime tempoInicio = dimc;
//                        preactor.PlanningBoard.PutOperationOnResource(ordem.Record, item.recursoId, dimc); // Aloca a operação no recurso

//                        DateTime tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", ordem.Record);

//                        // Adiciona ao sequenciamento
//                        sequenciamentoOperacoes.Add((ordem.Record, ordem.OrderNo, item.recursoId, tempoInicio, tempoFim));

//                        // Atualiza o dimc para o tempo de fim da operação
//                        dimc = tempoFim;
//                    }
//                }
//            }

//            return 0;
//        }

//        private static void GetResources(IPreactor preactor, IList<Resources> ListaRecursos)
//        {
//            // Coleta de dados das operações para cada uma dos OrdersId
//            for (int resourceRecord = 1; resourceRecord <= preactor.RecordCount("Resources"); resourceRecord++)
//            {
//                Resources Rec = new Resources();

//                Rec.ResourceName = preactor.ReadFieldString("Resources", "Name", resourceRecord);
//                Rec.Attribute1 = preactor.ReadFieldString("Resources", "Attribute 1", resourceRecord);
//                Rec.Attribute4 = preactor.ReadFieldString("Resources", "Attribute 4", resourceRecord);
//                Rec.ResourceId = resourceRecord;

//                ListaRecursos.Add(Rec);
//            }
//        }

//        private static void GetOrders(IPreactor preactor, IList<Orders> ListaOrders)
//        {
//            // Coleta de dados das operações para cada uma dos OrdersId
//            for (int OrdersRecord = 1; OrdersRecord <= preactor.RecordCount("Orders"); OrdersRecord++)
//            {

//                Orders Ord = new Orders();

//                Ord.Record = OrdersRecord;
//                Ord.OrderNo = preactor.ReadFieldString("Orders", "Order No.", OrdersRecord);
//                Ord.PartNo = preactor.ReadFieldString("Orders", "Part No.", OrdersRecord);
//                Ord.OpNo = preactor.ReadFieldString("Orders", "Op. No.", OrdersRecord);
//                Ord.OperationName = preactor.ReadFieldString("Orders", "Operation Name", OrdersRecord);
//                Ord.SetupStart = preactor.ReadFieldDateTime("Orders", "Setup Start", OrdersRecord);
//                Ord.StartTime = preactor.ReadFieldDateTime("Orders", "Start Time", OrdersRecord);
//                Ord.EndTime = preactor.ReadFieldDateTime("Orders", "End Time", OrdersRecord);
//                Ord.DueDate = preactor.ReadFieldDateTime("Orders", "Due Date", OrdersRecord);
//                Ord.Programada = false; // Inicializa a variável de controle
//                Ord.RecursoRequerido = -1; // Inicializa a variável de controle
//                Ord.OrdenacaoPeca = preactor.ReadFieldInt("Orders", "Numerical Attribute 1", OrdersRecord); // Inicializa a variável de controle para ordenação das peças
//                Ord.ValorOrdenacao = 10000; // Inicializa a variável de controle para ordenação das ordens de produção

//                ListaOrders.Add(Ord);

//            }
//        }

//    }
//}
