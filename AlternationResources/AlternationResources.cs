using Microsoft.Win32;
using Preactor;
using Preactor.Interop.PreactorObject;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Xml.Linq;

namespace NativeRules
{
    [Guid("01215791-0677-4f55-9803-635ab5694427")]
    [ComVisible(true)]
    public interface IAlternationOperation
    {
        int UnallocateAlternationOperation(ref PreactorObj preactorComObject, ref object pespComObject);
        int SelectedResources(ref PreactorObj preactorComObject, ref object pespComObject);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("3d2e0d4e-f45e-46da-ba1e-2568920ee92f")]
    public class AlternationOperation : IAlternationOperation
    {

        public int UnallocateAlternationOperation(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

            // ok // 1 - Sequenciar operações para frente
            // nk // // 2 - Sequenciar operações SOLDAR ROBO
            // -- // ok // 2.1 - Selecioanr Operações SOLDAR ROBO 
            // -- // ok // 2.2 - Desprogramar as operações SOLDAR ROBO e operações subsequentes
            // -- // nk // 2.3 - Programar operações SOLDAR ROBO de forma alternada
            // -- // ok // 2.4 - Corrigir tempo de operação para que seja possivel realizar 1 operacao de forma alternada no recurso Robo N
            // -- // nk // 2.5 - Consolidar/Agrupar operações SOLDAR ROBO de uma mesma ordem
            // nk // 3 - Sequenciar as operações posteriores a solda robo

            // Abre as listas de Ordens e Recursos
            IList<Orders> listaOrders = new List<Orders>();
            IList<Resources> listaRecursos = new List<Resources>();

            GetOrders(preactor, listaOrders);

            GetResources(preactor, listaRecursos);

            // Sequencia todas as operações para frente
            preactor.PlanningBoard.SequenceAll(SequenceAllDirection.Forwards, SequencePriority.DueDate);

            // Filtrando as operacoes de "SOLDAR ROBO" e ordenando por SetupStart
            var listaOrdemSoldarRobo = listaOrders
                .Where(s => s.OperationName == "SOLDAR ROBO")
                .ToList();

            // Iterar sobre as operacoes de "SOLDAR ROBO"
            foreach (var registoOrdem in listaOrdemSoldarRobo)
            {
                // busca a operacao anterior para que seja desprograma todas as operacoes subsequentes
                int previousRecord = preactor.PlanningBoard.GetPreviousOperation(registoOrdem.Record, 1);

                // Verificacao se existe uma operacao subsequente (se (PreviusRecord < 0 nao existe operacao antecessora)
                if (previousRecord > 0)
                {
                    // Desprograma as operações subsequentes
                    preactor.PlanningBoard.UnallocateOperation(previousRecord, OperationSelection.SubsequentOperations);
                    // Desprograma a operação de SOLDAR ROBO
                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.ThisOperation);
                }
                else
                {
                    // Se não houver operação anterior, desprogramar a operação atual
                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.BiDirectionalOperations);

                    // Desprograma a operação de SOLDAR ROBO
                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.ThisOperation);
                }
            }

            // --------------------------------------------- Excluir daqui

            return 0;

        }

        public int SelectedResources(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

            // Lista em que sera salvo os dados das ordens/operações
            IList<Orders> listaOrders = new List<Orders>();

            GetOrders(preactor, listaOrders);


            // == Dados Recursos ==

            // Lista em que será salvo os dados dos recursos
            IList<Resources> listaRecursos = new List<Resources>();

            GetResources(preactor, listaRecursos);

            // Filtrando as operacoes de "SOLDAR ROBO" e ordenando por SetupStart
            var listaOrdemSoldarRobo = listaOrders
                .Where(s => s.OperationName == "SOLDAR ROBO")
                .OrderBy(x => x.DueDate)
                .ToList();

            // --------------------------------------------- Até aqui excluir daqui



            // Cria uma lista com os dados filtrando os recursos de solda robo
            var listaRecursoRoboSolda = listaRecursos
                .Where(r => r.ResourceName != null &&
                            r.ResourceName.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0 &&
                            r.Attribute4.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(r => r.ResourceName)
                .ToList();

            // Cria uma lista com os agrupamentos de recursos de solda robo
            var listaGrupoRecursoRoboSolda = listaRecursoRoboSolda
                .GroupBy(r => r.Attribute4)
                .Select(s => new { Attribute4 = s.Key })
                .ToList();


            // == Dados Recursos ==

            // == Inicio Dados Ordens Soldar Robo ==


            // Agrupa as ordens pelo campo OrderNo
            var listaOrdensAgrupadasPorOrderNo = listaOrdemSoldarRobo
                .GroupBy(o => o.OrderNo)
                .ToList();

            // Cria um valor de ordenação alternado dentro de cada grupo de ordens
            int valorOrdenacaoCounter = 0;
            foreach (var grupo in listaOrdensAgrupadasPorOrderNo)
            {
                // Ordena as ordens dentro do grupo por DueDate
                var listaOrdensDoGrupoOrdenadas = grupo.OrderBy(o => o.DueDate).ToList();

                // Atribui um valor alternado de ordenação para as ordens dentro do grupo
                foreach (var ordem in listaOrdensDoGrupoOrdenadas)
                {
                    // Atualiza o valor de ordenação na lista original para todos os registros com o mesmo OrderNo
                    var ordensParaAtualizar = listaOrdemSoldarRobo
                        .Where(o => o.OrderNo == ordem.OrderNo)
                        .ToList();

                    foreach (var ordemOriginal in ordensParaAtualizar)
                    {
                        ordemOriginal.ValorOrdenacao = valorOrdenacaoCounter;
                    }
                    valorOrdenacaoCounter++;
                }
            }

            // Ordena a lista final com base no novo valor de ordenação
            var ListaOrdemSoldarRoboOrdenada = listaOrdemSoldarRobo
                .OrderBy(x => x.OrdenacaoPeca)  // Depois ordena por OrdenacaoPeca
                .ThenBy(x => x.ValorOrdenacao) // Ordena pelo valor sequêncial das ordens
                .ThenBy(x => x.DueDate)       // Por fim, ordena por DueDate
                .ToList();


            // == Fim Dados Ordens Soldar Robo ==


            // ============================================================================================================
            // Parte um, inicia o sequenciamento das operações SOLDAR ROBO colocando uma operação de cada ordem em cada recurso de solda robo
            // ============================================================================================================



            var queueOrdens = new Queue<Orders>(ListaOrdemSoldarRoboOrdenada);

            DateTime dimc = preactor.PlanningBoard.TerminatorTime;

            bool robo1 = false;
            DateTime tempoRobo1 = preactor.PlanningBoard.TerminatorTime;
            int quantidadeRestanteRobo1Mesa1 = 10;
            int quantidadeRestanteRobo1Mesa2 = 10;
            bool robo2 = false;
            DateTime tempoRobo2 = preactor.PlanningBoard.TerminatorTime;
            int quantidadeRestanteRobo2Mesa1 = 10;
            int quantidadeRestanteRobo2Mesa2 = 10;
            bool robo3 = false;
            DateTime tempoRobo3 = preactor.PlanningBoard.TerminatorTime;
            int quantidadeRestanteRobo3Mesa1 = 10;
            int quantidadeRestanteRobo3Mesa2 = 10;
            bool robo4 = false;
            DateTime tempoRobo4 = preactor.PlanningBoard.TerminatorTime;
            int quantidadeRestanteRobo4Mesa1 = 10;
            int quantidadeRestanteRobo4Mesa2 = 10;

            // Lista para armazenar os recursos que já foram programados
            List<int> recursosProgramados = new List<int>();

            // Tabela para armazenar os recursos e as ordens alocadas
            List<(int recursoId, string OrderNo, int ordemId, DateTime changeStart)> tabelaOrdensRecurso = new List<(int, string, int, DateTime)>();

            while (recursosProgramados.Count < listaRecursoRoboSolda.Count)
            {
                // Seleciona a primeira ordem
                var primeiraOrdem = queueOrdens.Dequeue();


                // Lista para armazenar os resultados dos testes
                List<(int recursoId, string Attribute4, DateTime changeStart)> listaResultadosTestes = new List<(int recursoId, string Attribute4, DateTime changeStart)>();

                // Itera sobre os recursos disponíveis (excluindo os já programados)
                foreach (var recurso in listaRecursoRoboSolda.Where(r => !recursosProgramados.Contains(r.ResourceId)))
                {
                    // Realiza o teste de operação para o recurso atual e a ordem
                    var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(primeiraOrdem.Record, recurso.ResourceId, dimc);

                    if (resultadoTeste.HasValue)
                    {
                        // Armazena o recurso e o tempo de início do teste
                        listaResultadosTestes.Add((recurso.ResourceId, recurso.Attribute4, resultadoTeste.Value.ChangeStart));
                    }
                }

                // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                if (listaResultadosTestes.Count > 0)
                {
                    // Ordena os resultados pelo tempo de início (ChangeStart)
                    var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                    // Verifica as condições para o PutOp
                    if (primeiraOrdem.RecursoRequerido == -1)
                    {
                        if (resultadoMinimo.Attribute4 == "ROBO 1" && robo1 == true)
                        {
                            resultadoMinimo.changeStart = tempoRobo1;
                        }
                        else if (resultadoMinimo.Attribute4 == "ROBO 2" && robo2 == true)
                        {
                            resultadoMinimo.changeStart = tempoRobo2;
                        }
                        else if (resultadoMinimo.Attribute4 == "ROBO 3" && robo3 == true)
                        {
                            resultadoMinimo.changeStart = tempoRobo3;
                        }
                        else if (resultadoMinimo.Attribute4 == "ROBO 4" && robo4 == true)
                        {
                            resultadoMinimo.changeStart = tempoRobo4;
                        }
                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                        preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);

                        // Marca a ordem como programada
                        primeiraOrdem.Programada = true;

                        // Atualiza o valor de ordenação na lista original para todos os registros com o mesmo OrderNo
                        var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada
                            .Where(o => o.Record == primeiraOrdem.Record)
                            .ToList();

                            foreach (var ordemOriginal in ordensParaAtualizar)
                            {
                                ordemOriginal.Programada = primeiraOrdem.Programada;
                            }
                            valorOrdenacaoCounter++;
                        
                        // Adiciona o recurso à lista de recursos programados
                        recursosProgramados.Add(resultadoMinimo.recursoId);

                        // Registra a ordem e o recurso na tabela
                        tabelaOrdensRecurso.Add((resultadoMinimo.recursoId, primeiraOrdem.OrderNo, primeiraOrdem.Record, resultadoMinimo.changeStart)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem

                        if (resultadoMinimo.Attribute4 == "ROBO 1" && robo1 == false)
                        {
                            tempoRobo1 = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
                            robo1 = true;
                        }
                        else if (resultadoMinimo.Attribute4 == "ROBO 2" && robo2 == false)
                        {
                            tempoRobo2 = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
                            robo2 = true;
                        }
                        else if (resultadoMinimo.Attribute4 == "ROBO 3" && robo3 == false)
                        {
                            tempoRobo3 = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
                            robo3 = true;
                        }
                        else if (resultadoMinimo.Attribute4 == "ROBO 4" && robo4 == false)
                        {
                            tempoRobo4 = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
                            robo4 = true;
                        }

                    }
                }
                else if (listaResultadosTestes.Count <= 0)
                {
                    queueOrdens.Enqueue(primeiraOrdem);
                }
            }



            // ============================================================================================================
            // Parte dois, continuar o sequenciamento das operações SOLDAR ROBO com a mesma ordem já sequenciada
            // ============================================================================================================



            // Lista para armazenar as operações sequenciadas
            List<(int ordemId, string OrderNo, int recursoId, DateTime startTime, DateTime endTime)> sequenciamentoOperacoes = new List<(int, string, int, DateTime, DateTime)>();

            for (int i = 0; i < ListaOrdemSoldarRoboOrdenada.Count; i++)
            {
                foreach (var ordem in ListaOrdemSoldarRoboOrdenada)
                    {
                    // Filtra as ordens que não estão programadas (Programada == false)
                    if (ordem.Programada == false)  // Verifica se a ordem não está programada
                    {
                        // Filtra as ordens que já estão programadas
                        var ordensComRecursoAlocado = tabelaOrdensRecurso.Where(t => t.OrderNo == ordem.OrderNo).OrderBy(t => t.OrderNo).ThenBy(t => t.changeStart).ToList();

                        foreach (var item in ordensComRecursoAlocado)
                        {
                            if (item.recursoId == preactor.PlanningBoard.GetResourceNumber("ROBO 1 MESA 1"))
                            {
                                quantidadeRestanteRobo1Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false);
                            }
                            else if (item.recursoId == preactor.PlanningBoard.GetResourceNumber("ROBO 1 MESA 2"))
                            {
                                quantidadeRestanteRobo1Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false);
                            }
                            else if (item.recursoId == preactor.PlanningBoard.GetResourceNumber("ROBO 2 MESA 1"))
                            {
                                quantidadeRestanteRobo2Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false);
                            }
                            else if (item.recursoId == preactor.PlanningBoard.GetResourceNumber("ROBO 2 MESA 2"))
                            {
                                quantidadeRestanteRobo2Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false);
                            }
                            else if (item.recursoId == preactor.PlanningBoard.GetResourceNumber("ROBO 3 MESA 1"))
                            {
                                quantidadeRestanteRobo3Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false);
                            }
                            else if (item.recursoId == preactor.PlanningBoard.GetResourceNumber("ROBO 3 MESA 2"))
                            {
                                quantidadeRestanteRobo3Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false);
                            }
                            else if (item.recursoId == preactor.PlanningBoard.GetResourceNumber("ROBO 4 MESA 1"))
                            {
                                quantidadeRestanteRobo4Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false);
                            }
                            else if (item.recursoId == preactor.PlanningBoard.GetResourceNumber("ROBO 4 MESA 2"))
                            {
                                quantidadeRestanteRobo4Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false);
                            }


                            // Verifica se é a primeira operação ou não
                            if (ordem.OrdenacaoPeca == 1)  // Se a ordem for a primeira, pega o tempo de início do recurso alocado
                            {
                                // Define o tempo de início e fim para a primeira operação
                                DateTime tempoInicio = item.changeStart;
                                DateTime tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", ordem.Record);

                                // Adiciona ao sequenciamento
                                sequenciamentoOperacoes.Add((ordem.Record, ordem.OrderNo, item.recursoId, tempoInicio, tempoFim));
                            }
                            else
                            {
                                string grupoRecurso = preactor.ReadFieldString("Resources", "Attribute 4", item.recursoId);

                                var resourcesFound = listaRecursos
                                    .Where(r => r.Attribute4.Equals(grupoRecurso, StringComparison.OrdinalIgnoreCase))
                                    .ToList();

                                DateTime ultimoEndTimeProgramado = ordem.EndTime;

                                // Verificando se há recursos encontrados
                                if (resourcesFound.Count >= 2)  // Verificando se existem pelo menos 2 recursos para comparação
                                {
                                    // Fazendo chamadas para obter os resultados com a data ajustada para 120 dias
                                    var result1 = preactor.PlanningBoard.GetResourcesPreviousOperation(resourcesFound[0].ResourceId, DateTime.MaxValue);
                                    var result2 = preactor.PlanningBoard.GetResourcesPreviousOperation(resourcesFound[1].ResourceId, DateTime.MaxValue);

                                    // Obtendo os tempos de término
                                    DateTime endtime1 = preactor.ReadFieldDateTime("Orders", "End Time", result1);
                                    DateTime endtime2 = preactor.ReadFieldDateTime("Orders", "End Time", result2);

                                    // Comparando as datas para encontrar a mais recente
                                    ultimoEndTimeProgramado = endtime1 > endtime2 ? endtime1 : endtime2;
                                }
                                else if (resourcesFound.Count == 1)
                                {
                                    // Fazendo chamadas para obter os resultados com a data ajustada para 120 dias
                                    var result1 = preactor.PlanningBoard.GetResourcesPreviousOperation(resourcesFound[0].ResourceId, DateTime.Today.AddDays(120));

                                    // Obtendo os tempos de término
                                    DateTime endtime1 = preactor.ReadFieldDateTime("Orders", "End Time", result1);

                                    // Comparando as datas para encontrar a mais recente
                                    ultimoEndTimeProgramado = endtime1;
                                }

                                dimc = ultimoEndTimeProgramado;

                                int ultimoOrderIdProrgamado = preactor.PlanningBoard.GetResourcesPreviousOperation(item.recursoId, item.changeStart.AddDays(15));

                                // Para ordens subsequentes, usa o dimc do ciclo anterior
                                DateTime tempoInicio = dimc;

                                // Aloca a operação no recurso com base no dimc
                                preactor.PlanningBoard.PutOperationOnResource(ordem.Record, item.recursoId, dimc);

                                // Marca a ordem como programada
                                ordem.Programada = true;

                                // Atualiza o valor de ordenação na lista original para todos os registros com o mesmo OrderNo
                                var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada
                                    .Where(o => o.Record == ordem.Record)
                                    .ToList();

                                foreach (var ordemOriginal in ordensParaAtualizar)
                                {
                                    ordemOriginal.Programada = ordem.Programada;
                                }
                                valorOrdenacaoCounter++;

                                // Define o tempo de fim da operação
                                DateTime tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", ordem.Record);

                                // Adiciona ao sequenciamento
                                sequenciamentoOperacoes.Add((ordem.Record, ordem.OrderNo, item.recursoId, tempoInicio, tempoFim));

                                // Atualiza o dimc para o tempo de fim da operação
                                dimc = tempoFim;


                                // Adiciona a nova ordem ao recurso que acabou a ordem


                                if (quantidadeRestanteRobo1Mesa1 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 1 MESA 1");

                                    // Lista para armazenar os resultados dos testes
                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int, string, DateTime)>();

                                    int totalOrdens = queueOrdens.Count;

                                    for (int interacaoQueu = 0; interacaoQueu < totalOrdens; interacaoQueu++)
                                    {
                                        var ordemAtual = queueOrdens.Dequeue();

                                        // Filtra as ordens que não estão programadas
                                        if (!ordemAtual.Programada)
                                        {
                                            // Realiza o teste de operação para o recurso atual e a ordem
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordemAtual.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                // Armazena o recurso e o tempo de início do teste
                                                listaResultadosTestes.Add((ordemAtual.Record, ordemAtual.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                            else
                                            {
                                                // Se o teste falhar, adiciona a ordem de volta à fila
                                                queueOrdens.Enqueue(ordemAtual);
                                            }
                                        }
                                        
                                    }
                                        
                                    // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        // Ordena os resultados pelo tempo de início (ChangeStart)
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        // Define o tempo de fim da operação
                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        // Adiciona ao sequenciamento
                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        // Atualiza o dimc para o tempo de fim da operação
                                        dimc = tempoFim;

                                        // Adiciona o recurso à lista de recursos programados
                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
                                        quantidadeRestanteRobo1Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item); // Remove o item da tabela de ordens e recursos programados

                                        i = 0; // Reinicia o loop para verificar novamente as ordens 
                                         
                                    }
                                }
                                else if (quantidadeRestanteRobo1Mesa2 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 1 MESA 2");

                                    // Lista para armazenar os resultados dos testes
                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int, string, DateTime)>();

                                    int totalOrdens = queueOrdens.Count;

                                    for (int interacaoQueu = 0; interacaoQueu < totalOrdens; interacaoQueu++)
                                    {
                                        var ordemAtual = queueOrdens.Dequeue();

                                        // Filtra as ordens que não estão programadas
                                        if (!ordemAtual.Programada)
                                        {
                                            // Realiza o teste de operação para o recurso atual e a ordem
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordemAtual.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                // Armazena o recurso e o tempo de início do teste
                                                listaResultadosTestes.Add((ordemAtual.Record, ordemAtual.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                            else
                                            {
                                                // Se o teste falhar, adiciona a ordem de volta à fila
                                                queueOrdens.Enqueue(ordemAtual);
                                            }
                                        }

                                    }

                                    // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        // Ordena os resultados pelo tempo de início (ChangeStart)
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        // Define o tempo de fim da operação
                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        // Adiciona ao sequenciamento
                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        // Atualiza o dimc para o tempo de fim da operação
                                        dimc = tempoFim;

                                        // Adiciona o recurso à lista de recursos programados
                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
                                        quantidadeRestanteRobo1Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item); // Remove o item da tabela de ordens e recursos programados

                                        i = 0; // Reinicia o loop para verificar novamente as ordens 

                                    }
                                }

                                else if (quantidadeRestanteRobo2Mesa1 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 2 MESA 1");

                                    // Lista para armazenar os resultados dos testes
                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int, string, DateTime)>();

                                    int totalOrdens = queueOrdens.Count;

                                    for (int interacaoQueu = 0; interacaoQueu < totalOrdens; interacaoQueu++)
                                    {
                                        var ordemAtual = queueOrdens.Dequeue();

                                        // Filtra as ordens que não estão programadas
                                        if (!ordemAtual.Programada)
                                        {
                                            // Realiza o teste de operação para o recurso atual e a ordem
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordemAtual.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                // Armazena o recurso e o tempo de início do teste
                                                listaResultadosTestes.Add((ordemAtual.Record, ordemAtual.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                            else
                                            {
                                                // Se o teste falhar, adiciona a ordem de volta à fila
                                                queueOrdens.Enqueue(ordemAtual);
                                            }
                                        }

                                    }

                                    // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        // Ordena os resultados pelo tempo de início (ChangeStart)
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        // Define o tempo de fim da operação
                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        // Adiciona ao sequenciamento
                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        // Atualiza o dimc para o tempo de fim da operação
                                        dimc = tempoFim;

                                        // Adiciona o recurso à lista de recursos programados
                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
                                        quantidadeRestanteRobo2Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item); // Remove o item da tabela de ordens e recursos programados

                                        i = 0; // Reinicia o loop para verificar novamente as ordens 

                                    }
                                }
                                else if (quantidadeRestanteRobo2Mesa2 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 2 MESA 2");

                                    // Lista para armazenar os resultados dos testes
                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int, string, DateTime)>();

                                    int totalOrdens = queueOrdens.Count;

                                    for (int interacaoQueu = 0; interacaoQueu < totalOrdens; interacaoQueu++)
                                    {
                                        var ordemAtual = queueOrdens.Dequeue();

                                        // Filtra as ordens que não estão programadas
                                        if (!ordemAtual.Programada)
                                        {
                                            // Realiza o teste de operação para o recurso atual e a ordem
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordemAtual.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                // Armazena o recurso e o tempo de início do teste
                                                listaResultadosTestes.Add((ordemAtual.Record, ordemAtual.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                            else
                                            {
                                                // Se o teste falhar, adiciona a ordem de volta à fila
                                                queueOrdens.Enqueue(ordemAtual);
                                            }
                                        }

                                    }

                                    // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        // Ordena os resultados pelo tempo de início (ChangeStart)
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        // Define o tempo de fim da operação
                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        // Adiciona ao sequenciamento
                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        // Atualiza o dimc para o tempo de fim da operação
                                        dimc = tempoFim;

                                        // Adiciona o recurso à lista de recursos programados
                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
                                        quantidadeRestanteRobo2Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item); // Remove o item da tabela de ordens e recursos programados

                                        i = 0; // Reinicia o loop para verificar novamente as ordens 

                                    }
                                }

                                else if (quantidadeRestanteRobo3Mesa1 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 3 MESA 1");

                                    // Lista para armazenar os resultados dos testes
                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int, string, DateTime)>();

                                    int totalOrdens = queueOrdens.Count;

                                    for (int interacaoQueu = 0; interacaoQueu < totalOrdens; interacaoQueu++)
                                    {
                                        var ordemAtual = queueOrdens.Dequeue();

                                        // Filtra as ordens que não estão programadas
                                        if (!ordemAtual.Programada)
                                        {
                                            // Realiza o teste de operação para o recurso atual e a ordem
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordemAtual.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                // Armazena o recurso e o tempo de início do teste
                                                listaResultadosTestes.Add((ordemAtual.Record, ordemAtual.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                            else
                                            {
                                                // Se o teste falhar, adiciona a ordem de volta à fila
                                                queueOrdens.Enqueue(ordemAtual);
                                            }
                                        }

                                    }

                                    // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        // Ordena os resultados pelo tempo de início (ChangeStart)
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        // Define o tempo de fim da operação
                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        // Adiciona ao sequenciamento
                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        // Atualiza o dimc para o tempo de fim da operação
                                        dimc = tempoFim;

                                        // Adiciona o recurso à lista de recursos programados
                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
                                        quantidadeRestanteRobo3Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item); // Remove o item da tabela de ordens e recursos programados

                                        i = 0; // Reinicia o loop para verificar novamente as ordens 

                                    }
                                }
                                else if (quantidadeRestanteRobo3Mesa2 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 3 MESA 2");

                                    // Lista para armazenar os resultados dos testes
                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int, string, DateTime)>();

                                    int totalOrdens = queueOrdens.Count;

                                    for (int interacaoQueu = 0; interacaoQueu < totalOrdens; interacaoQueu++)
                                    {
                                        var ordemAtual = queueOrdens.Dequeue();

                                        // Filtra as ordens que não estão programadas
                                        if (!ordemAtual.Programada)
                                        {
                                            // Realiza o teste de operação para o recurso atual e a ordem
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordemAtual.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                // Armazena o recurso e o tempo de início do teste
                                                listaResultadosTestes.Add((ordemAtual.Record, ordemAtual.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                            else
                                            {
                                                // Se o teste falhar, adiciona a ordem de volta à fila
                                                queueOrdens.Enqueue(ordemAtual);
                                            }
                                        }

                                    }

                                    // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        // Ordena os resultados pelo tempo de início (ChangeStart)
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        // Define o tempo de fim da operação
                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        // Adiciona ao sequenciamento
                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        // Atualiza o dimc para o tempo de fim da operação
                                        dimc = tempoFim;

                                        // Adiciona o recurso à lista de recursos programados
                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
                                        quantidadeRestanteRobo3Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item); // Remove o item da tabela de ordens e recursos programados

                                        i = 0; // Reinicia o loop para verificar novamente as ordens 

                                    }
                                }

                                else if (quantidadeRestanteRobo4Mesa1 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 4 MESA 1");

                                    // Lista para armazenar os resultados dos testes
                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int, string, DateTime)>();

                                    int totalOrdens = queueOrdens.Count;

                                    for (int interacaoQueu = 0; interacaoQueu < totalOrdens; interacaoQueu++)
                                    {
                                        var ordemAtual = queueOrdens.Dequeue();

                                        // Filtra as ordens que não estão programadas
                                        if (!ordemAtual.Programada)
                                        {
                                            // Realiza o teste de operação para o recurso atual e a ordem
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordemAtual.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                // Armazena o recurso e o tempo de início do teste
                                                listaResultadosTestes.Add((ordemAtual.Record, ordemAtual.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                            else
                                            {
                                                // Se o teste falhar, adiciona a ordem de volta à fila
                                                queueOrdens.Enqueue(ordemAtual);
                                            }
                                        }

                                    }

                                    // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        // Ordena os resultados pelo tempo de início (ChangeStart)
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        // Define o tempo de fim da operação
                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        // Adiciona ao sequenciamento
                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        // Atualiza o dimc para o tempo de fim da operação
                                        dimc = tempoFim;

                                        // Adiciona o recurso à lista de recursos programados
                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
                                        quantidadeRestanteRobo4Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item); // Remove o item da tabela de ordens e recursos programados

                                        i = 0; // Reinicia o loop para verificar novamente as ordens 

                                    }
                                }
                                else if (quantidadeRestanteRobo4Mesa2 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 4 MESA 2");

                                    // Lista para armazenar os resultados dos testes
                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int, string, DateTime)>();

                                    int totalOrdens = queueOrdens.Count;

                                    for (int interacaoQueu = 0; interacaoQueu < totalOrdens; interacaoQueu++)
                                    {
                                        var ordemAtual = queueOrdens.Dequeue();

                                        // Filtra as ordens que não estão programadas
                                        if (!ordemAtual.Programada)
                                        {
                                            // Realiza o teste de operação para o recurso atual e a ordem
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordemAtual.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                // Armazena o recurso e o tempo de início do teste
                                                listaResultadosTestes.Add((ordemAtual.Record, ordemAtual.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                            else
                                            {
                                                // Se o teste falhar, adiciona a ordem de volta à fila
                                                queueOrdens.Enqueue(ordemAtual);
                                            }
                                        }

                                    }

                                    // Verifica se obteve algum resultado e faz o "PutOp" com o menor tempo de início
                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        // Ordena os resultados pelo tempo de início (ChangeStart)
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        // Realiza o "PutOperation" com o menor tempo de início (ChangeStart)
                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        // Define o tempo de fim da operação
                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        // Adiciona ao sequenciamento
                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        // Atualiza o dimc para o tempo de fim da operação
                                        dimc = tempoFim;

                                        // Adiciona o recurso à lista de recursos programados
                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim)); // Substitua 'primeiraOrdem.Id' pelo campo correto que identifica a ordem
                                        quantidadeRestanteRobo4Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item); // Remove o item da tabela de ordens e recursos programados

                                        i = 0; // Reinicia o loop para verificar novamente as ordens 

                                    }
                                }

                            }
                        }
                    }
                }
            }
            return 0;
        }

        // Busca os dados dos recursos (ListaRecursos.Add(Rec);)
        private static void GetResources(IPreactor preactor, IList<Resources> ListaRecursos)
        {
            // Coleta de dados das operações para cada uma dos OrdersId
            for (int resourceRecord = 1; resourceRecord <= preactor.RecordCount("Resources"); resourceRecord++)
            {
                Resources Rec = new Resources();

                Rec.ResourceName = preactor.ReadFieldString("Resources", "Name", resourceRecord);
                Rec.Attribute1 = preactor.ReadFieldString("Resources", "Attribute 1", resourceRecord);
                Rec.Attribute4 = preactor.ReadFieldString("Resources", "Attribute 4", resourceRecord);
                Rec.ResourceId = resourceRecord;

                ListaRecursos.Add(Rec);
            }
        }

        // Busca os dados das ordens (ListaOrders.Add(Ord))
        private static void GetOrders(IPreactor preactor, IList<Orders> ListaOrders)
        {
            // Coleta de dados das operações para cada uma dos OrdersId
            for (int OrdersRecord = 1; OrdersRecord <= preactor.RecordCount("Orders"); OrdersRecord++)
            {

                Orders Ord = new Orders();

                Ord.Record = OrdersRecord;
                Ord.OrderNo = preactor.ReadFieldString("Orders", "Order No.", OrdersRecord);
                Ord.PartNo = preactor.ReadFieldString("Orders", "Part No.", OrdersRecord);
                Ord.OpNo = preactor.ReadFieldString("Orders", "Op. No.", OrdersRecord);
                Ord.OperationName = preactor.ReadFieldString("Orders", "Operation Name", OrdersRecord);
                Ord.SetupStart = preactor.ReadFieldDateTime("Orders", "Setup Start", OrdersRecord);
                Ord.StartTime = preactor.ReadFieldDateTime("Orders", "Start Time", OrdersRecord);
                Ord.EndTime = preactor.ReadFieldDateTime("Orders", "End Time", OrdersRecord);
                Ord.DueDate = preactor.ReadFieldDateTime("Orders", "Due Date", OrdersRecord);
                Ord.Programada = false; // Inicializa a variável de controle
                Ord.RecursoRequerido = -1; // Inicializa a variável de controle
                Ord.OrdenacaoPeca = preactor.ReadFieldInt("Orders", "Numerical Attribute 1", OrdersRecord); // Inicializa a variável de controle para ordenação das peças
                Ord.ValorOrdenacao = 10000; // Inicializa a variável de controle para ordenação das ordens de produção

                ListaOrders.Add(Ord);

            }
        }

    }
}
