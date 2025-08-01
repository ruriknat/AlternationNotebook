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

//            // Lista em que sera salvo os dados das ordens/operações
//            IList<Orders> ListaOrders = new List<Orders>();

//            // Inicia o contador para salvar os dados de ordens/operações
//            int OrdersRecord = 1;

//            // Coleta de dados das operações para cada uma dos OrdersId
//            for (OrdersRecord = 1; OrdersRecord <= preactor.RecordCount("Orders"); OrdersRecord++)
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

//            // ==============================================================
//            // INICIO Ação 1

//            // Sequencia todas as operações para frente
//            preactor.PlanningBoard.SequenceAll(SequenceAllDirection.Forwards, SequencePriority.DueDate);

//            // FIM Ação 1
//            // ==============================================================


//            // ==============================================================
//            // Ação 2

//            // Filtrando as operacoes de "SOLDAR ROBO" e ordenando por SetupStart
//            var ListaOrdemSoldarRobo = ListaOrders.Where(s => s.OperationName == "SOLDAR ROBO" && s.SetupStart != null).ToList();

//            // Iterar sobre as operacoes de "SOLDAR ROBO"
//            foreach (var RegistoOrdem in ListaOrdemSoldarRobo)
//            {
//                // busca a operacao anterior para que seja desprograma todas as operacoes subsequentes
//                int PreviusRecord = preactor.PlanningBoard.GetPreviousOperation(RegistoOrdem.Record, 1);

//                // Verificacao se existe uma operacao subsequente (se (PreviusRecord < 0 nao existe operacao antecessora)
//                if (PreviusRecord > 0)
//                {
//                    // Desprograma as operacoes subsequentes
//                    preactor.PlanningBoard.UnallocateOperation(PreviusRecord, OperationSelection.SubsequentOperations);
//                    // Desprograma as operacoes de SOLDAR ROBO
//                    preactor.PlanningBoard.UnallocateOperation(RegistoOrdem.Record, OperationSelection.ThisOperation);
//                }
//                else
//                {
//                    // Se não houver operação anterior, desprogramar a operação atual
//                    preactor.PlanningBoard.UnallocateOperation(RegistoOrdem.Record, OperationSelection.BiDirectionalOperations);

//                    // Desprograma as operacoes de SOLDAR ROBO
//                    preactor.PlanningBoard.UnallocateOperation(RegistoOrdem.Record, OperationSelection.ThisOperation);
//                }
//            }

//            // ---------------------- Excluir daqui


//            return 0;
//        }

//        public int SelectedResources(ref PreactorObj preactorComObject, ref object pespComObject)
//        {
//            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

//            // Lista em que sera salvo os dados das ordens/operações
//            IList<Orders> ListaOrders = new List<Orders>();

//            // Inicia o contador para salvar os dados de ordens/operações
//            int OrdersRecord = 1;

//            // Coleta de dados das operações para cada uma dos OrdersId
//            for (OrdersRecord = 1; OrdersRecord <= preactor.RecordCount("Orders"); OrdersRecord++)
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


//            // ---------------------- Até aqui excluir daqui


//            // == Dados Recursos ==

//                // Lista em que sera salvo os dados dos recursso
//                IList<Resources> ListaRecursos = new List<Resources>();

//                // Inicia o contador de recursos
//                int resourceRecord = 1;

//                // Coleta de dados das operações para cada uma dos OrdersId
//                for (resourceRecord = 1; resourceRecord <= preactor.RecordCount("Resources"); resourceRecord++)
//                {
//                    Resources Rec = new Resources();

//                    Rec.ResourceName = preactor.ReadFieldString("Resources", "Name", resourceRecord);
//                    Rec.Attribute1 = preactor.ReadFieldString("Resources", "Attribute 1", resourceRecord);
//                    Rec.Attribute4 = preactor.ReadFieldString("Resources", "Attribute 4", resourceRecord);
//                    Rec.ResourceId = resourceRecord;

//                    ListaRecursos.Add(Rec);
//                }

//                // Cria uma lista com os dados filtrando os recursso de solda robo
//                var ListaRoboResources = ListaRecursos.Where(r => r.ResourceName != null && r.ResourceName.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0 && r.Attribute4.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(r => r.ResourceName).ToList();

//                // Cria uma lista com os agrupamentos de recursos de solda robo
//                var ListaGrupoRoboResources = ListaRoboResources.GroupBy(r => r.Attribute4).Select(s => new { Attribute4 = s.Key }).ToList();

//            // == Dados Recursos ==

//            // == Inicio Dados Ordens Soldar Robo ==

//            // Filtra e ordena as ordens "SOLDAR ROBO" por DueDate
//            var listaOrdenadaPorDueDate = ListaOrders
//                .Where(s => s.OperationName == "SOLDAR ROBO")
//                .OrderBy(x => x.DueDate)
//                .ToList();

//            // Agrupa as ordens pelo campo OrderNo
//            var listaOrdensAgrupadasPorOrderNo = listaOrdenadaPorDueDate
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
//                    var ordensParaAtualizar = listaOrdenadaPorDueDate
//                        .Where(o => o.OrderNo == ordem.OrderNo)
//                        .ToList();

//                    foreach (var ordemOriginal in ordensParaAtualizar)
//                    {
//                        ordemOriginal.ValorOrdenacao = valorOrdenacaoCounter;
//                    }

//                    valorOrdenacaoCounter++; // Alterna entre 1 e 2
//                }
//            }


//            //// Ordena a lista final com base no novo valor de ordenação
//            //var ListaOrdemSoldarRoboOrdenada = listaOrdenadaPorDueDate
//            //    .OrderBy(x => x.ValorOrdenacao) // Ordena pelo valor alternado
//            //    .ThenBy(x => x.OrdenacaoPeca)  // Depois ordena por OrdenacaoPeca
//            //    .ThenBy(x => x.DueDate)       // Por fim, ordena por DueDate
//            //    .ToList();

//            // Ordena a lista final com base no novo valor de ordenação
//            var ListaOrdemSoldarRoboOrdenada = listaOrdenadaPorDueDate
//                .OrderBy(x => x.OrdenacaoPeca)  // Depois ordena por OrdenacaoPeca
//                .ThenBy(x => x.ValorOrdenacao) // Ordena pelo valor sequêncial das ordens
//                .ThenBy(x => x.DueDate)       // Por fim, ordena por DueDate
//                .ToList();


//            // == Fim Dados Ordens Soldar Robo ==


//            // == Interações Solda Robo ==
            
//            List<int> listaRecursos = new List<int>();

//            foreach (var RegistroGrupoRecurso in ListaGrupoRoboResources)
//            {
//                DateTime dimc = DateTime.MinValue;
//                // Limpa a lista no início de cada interação
//                listaRecursos.Clear();

//                foreach (var RegistroRecurso in ListaRoboResources)
//                {
//                    if (RegistroRecurso.Attribute4 == RegistroGrupoRecurso.Attribute4)
//                    {
//                        // Verifica se o recurso já foi atribuído
//                        if (!listaRecursos.Contains(RegistroRecurso.ResourceId))
//                        {
//                            listaRecursos.Add(RegistroRecurso.ResourceId);
//                        }
//                    }
//                }
//                // Verifica se há recursos suficientes na lista antes de acessar os índices
//                if (listaRecursos.Count > 0)
//                {
//                    var queueOrdens = new Queue<Orders>(ListaOrdemSoldarRoboOrdenada);

//                    // Processa as ordens de solda robô
//                    int indiceRecurso = 0; // Começa do recurso 1 (índice 0)

//                    // Processa as ordens de solda robô
//                    while (queueOrdens.Count > 0)
//                    {
//                        var RegistoOrdem = queueOrdens.Dequeue();  // Remove o primeiro item da fila

//                        // Atualiza o DIMC (Data Inicial de Mais Cedo da Ordem)
//                        preactor.WriteField("Orders", "Earliest Start Date", RegistoOrdem.Record, dimc);

//                        // Criação de uma lista para armazenar os resultados dos testes (do tipo bool)
//                        List<Preactor.OperationTimes?> listaTesteop = new List<Preactor.OperationTimes?>();

//                        // Testa a operação nos recursos, verificando se a lista tem ao menos 1 recurso
//                        foreach (var recurso in listaRecursos)
//                        {
//                            // Chama o método que retorna um valor booleano e armazena o resultado na lista
//                            var testeop = preactor.PlanningBoard.TestOperationOnResource(RegistoOrdem.Record, recurso, dimc);

//                            // Adiciona o resultado do teste na lista de testes
//                            listaTesteop.Add(testeop);
//                        }

//                        // Verifica se o índice do recurso é válido na lista
//                        if (indiceRecurso >= listaRecursos.Count)
//                        {
//                            // Se o índice ultrapassar o número de recursos, reinicia o ciclo
//                            indiceRecurso = 0;
//                        }

//                        // Pega o recurso atual com base no índice
//                        var recursoSelecionado = listaRecursos[indiceRecurso];

//                        // Verifica se o teste de operação foi bem-sucedido para o recurso
//                        var testeopSelecionado = listaTesteop[indiceRecurso];

//                        if (testeopSelecionado.HasValue)
//                        {
//                            // Aloca o recurso com a data de início mais cedo
//                            preactor.PlanningBoard.PutOperationOnResource(RegistoOrdem.Record, recursoSelecionado, testeopSelecionado.Value.ChangeStart);
//                            RegistoOrdem.Programada = true;

//                            // Atualiza a lista de ordens com o recurso alocado
//                            foreach (var ordem in ListaOrders.Where(o => o.OrderNo == RegistoOrdem.OrderNo))
//                            {
//                                ordem.RecursoRequerido = recursoSelecionado;
//                            }

//                            // Atualiza a data de fim
//                            dimc = preactor.ReadFieldDateTime("Orders", "End Time", RegistoOrdem.Record);

//                            // Aqui é importante garantir que a ordem só sairá do ciclo quando estiver concluída
//                            // Em vez de continuar para o próximo recurso imediatamente, vamos verificar se a ordem terminou
//                            // Vamos continuar no mesmo recurso para a próxima iteração se a ordem não foi finalizada
//                            if (RegistoOrdem.Programada)
//                            {
//                                // A ordem foi programada, então avançamos para a próxima operação dentro dessa ordem
//                                // Aqui o código só deve sair dessa ordem quando a programação for completa
//                                continue; // Garante que a mesma ordem será processada até ser finalizada
//                            }
//                        }

//                        // Avança para o próximo recurso (ciclo alternado)
//                        indiceRecurso++;
//                    }
//                }
                
//            }

//            return 0;
//        }
//    }
//}
