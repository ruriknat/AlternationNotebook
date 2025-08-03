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

            IList<Orders> listaOrders = new List<Orders>();
            IList<Resources> listaRecursos = new List<Resources>();

            GetOrders(preactor, listaOrders);
            GetResources(preactor, listaRecursos);

            preactor.PlanningBoard.SequenceAll(SequenceAllDirection.Forwards, SequencePriority.DueDate);

            var listaOrdemSoldarRobo = listaOrders
                .Where(s => s.OperationName == "SOLDAR ROBO")
                .ToList();

            foreach (var registoOrdem in listaOrdemSoldarRobo)
            {
                int previousRecord = preactor.PlanningBoard.GetPreviousOperation(registoOrdem.Record, 1);

                if (previousRecord > 0)
                {
                    preactor.PlanningBoard.UnallocateOperation(previousRecord, OperationSelection.SubsequentOperations);
                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.ThisOperation);
                }
                else
                {
                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.BiDirectionalOperations);
                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.ThisOperation);
                }
            }

            // --------------------------------------------- Excluir daqui

            return 0;

        }

        public int SelectedResources(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

            IList<Orders> listaOrders = new List<Orders>();
            IList<Resources> listaRecursos = new List<Resources>();

            GetOrders(preactor, listaOrders);
            GetResources(preactor, listaRecursos);

            var listaOrdemSoldarRobo = listaOrders
                .Where(s => s.OperationName == "SOLDAR ROBO")
                .OrderBy(x => x.DueDate)
                .ToList();

            // --------------------------------------------- Até aqui excluir daqui

            var listaRecursoRoboSolda = listaRecursos
                .Where(r => r.ResourceName != null &&
                            r.ResourceName.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0 &&
                            r.Attribute4.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(r => r.ResourceName)
                .ToList();

            var listaGrupoRecursoRoboSolda = listaRecursoRoboSolda
                .GroupBy(r => r.Attribute4)
                .Select(s => new { Attribute4 = s.Key })
                .ToList();

            var listaOrdensAgrupadasPorOrderNo = listaOrdemSoldarRobo
                .GroupBy(o => o.OrderNo)
                .ToList();

            int valorOrdenacaoCounter = 0;
            foreach (var grupo in listaOrdensAgrupadasPorOrderNo)
            {
                var listaOrdensDoGrupoOrdenadas = grupo.OrderBy(o => o.DueDate).ToList();

                foreach (var ordem in listaOrdensDoGrupoOrdenadas)
                {
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

            var ListaOrdemSoldarRoboOrdenada = listaOrdemSoldarRobo
                .OrderBy(x => x.OrdenacaoPeca)
                .ThenBy(x => x.ValorOrdenacao)
                .ThenBy(x => x.DueDate)
                .ToList();



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

            List<int> recursosProgramados = new List<int>();
            List<(int recursoId, string OrderNo, int ordemId, DateTime changeStart)> tabelaOrdensRecurso = new List<(int, string, int, DateTime)>();

            while (recursosProgramados.Count < listaRecursoRoboSolda.Count)
            {
                var primeiraOrdem = queueOrdens.Dequeue();

                List<(int recursoId, string Attribute4, DateTime changeStart)> listaResultadosTestes = new List<(int recursoId, string Attribute4, DateTime changeStart)>();

                foreach (var recurso in listaRecursoRoboSolda.Where(r => !recursosProgramados.Contains(r.ResourceId)))
                {
                    var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(primeiraOrdem.Record, recurso.ResourceId, dimc);

                    if (resultadoTeste.HasValue)
                    {
                        listaResultadosTestes.Add((recurso.ResourceId, recurso.Attribute4, resultadoTeste.Value.ChangeStart));
                    }
                }

                if (listaResultadosTestes.Count > 0)
                {
                    var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

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
                        preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);

                        primeiraOrdem.Programada = true;

                        var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada
                            .Where(o => o.Record == primeiraOrdem.Record)
                            .ToList();

                        foreach (var ordemOriginal in ordensParaAtualizar)
                        {
                            ordemOriginal.Programada = primeiraOrdem.Programada;
                        }
                        valorOrdenacaoCounter++;

                        recursosProgramados.Add(resultadoMinimo.recursoId);

                        tabelaOrdensRecurso.Add((resultadoMinimo.recursoId, primeiraOrdem.OrderNo, primeiraOrdem.Record, resultadoMinimo.changeStart));

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



            List<(int ordemId, string OrderNo, int recursoId, DateTime startTime, DateTime endTime)> sequenciamentoOperacoes = new List<(int, string, int, DateTime, DateTime)>();

            for (int i = 0; i < ListaOrdemSoldarRoboOrdenada.Count; i++)
            {
                foreach (var ordem in ListaOrdemSoldarRoboOrdenada)
                {
                    if (ordem.Programada == false)  // Verifica se a ordem não está programada
                    {
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

                            if (ordem.OrdenacaoPeca == 1)
                            {
                                DateTime tempoInicio = item.changeStart;
                                DateTime tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", ordem.Record);
                                sequenciamentoOperacoes.Add((ordem.Record, ordem.OrderNo, item.recursoId, tempoInicio, tempoFim));
                            }
                            else
                            {
                                string grupoRecurso = preactor.ReadFieldString("Resources", "Attribute 4", item.recursoId);

                                var resourcesFound = listaRecursos
                                    .Where(r => r.Attribute4.Equals(grupoRecurso, StringComparison.OrdinalIgnoreCase))
                                    .ToList();

                                DateTime ultimoEndTimeProgramado = ordem.EndTime;

                                if (resourcesFound.Count >= 2)
                                {
                                    var result1 = preactor.PlanningBoard.GetResourcesPreviousOperation(resourcesFound[0].ResourceId, DateTime.MaxValue);
                                    var result2 = preactor.PlanningBoard.GetResourcesPreviousOperation(resourcesFound[1].ResourceId, DateTime.MaxValue);

                                    DateTime endtime1 = preactor.ReadFieldDateTime("Orders", "End Time", result1);
                                    DateTime endtime2 = preactor.ReadFieldDateTime("Orders", "End Time", result2);

                                    ultimoEndTimeProgramado = endtime1 > endtime2 ? endtime1 : endtime2;
                                }
                                else if (resourcesFound.Count == 1)
                                {
                                    var result1 = preactor.PlanningBoard.GetResourcesPreviousOperation(resourcesFound[0].ResourceId, DateTime.Today.AddDays(120));

                                    DateTime endtime1 = preactor.ReadFieldDateTime("Orders", "End Time", result1);

                                    ultimoEndTimeProgramado = endtime1;
                                }

                                dimc = ultimoEndTimeProgramado;

                                int ultimoOrderIdProrgamado = preactor.PlanningBoard.GetResourcesPreviousOperation(item.recursoId, item.changeStart.AddDays(15));

                                DateTime tempoInicio = dimc;

                                preactor.PlanningBoard.PutOperationOnResource(ordem.Record, item.recursoId, dimc);

                                ordem.Programada = true;

                                var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada
                                    .Where(o => o.Record == ordem.Record)
                                    .ToList();

                                foreach (var ordemOriginal in ordensParaAtualizar)
                                {
                                    ordemOriginal.Programada = ordem.Programada;
                                }
                                valorOrdenacaoCounter++;

                                DateTime tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", ordem.Record);

                                sequenciamentoOperacoes.Add((ordem.Record, ordem.OrderNo, item.recursoId, tempoInicio, tempoFim));

                                dimc = tempoFim;


                                // ============================================================================================================
                                // Parte três, continuar o sequenciamento das operações SOLDAR ROBO com as demais ordne ao acabar uma ordem
                                // ============================================================================================================


                                if (quantidadeRestanteRobo1Mesa1 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 1 MESA 1");

                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int OrdersId, string OrderNo, DateTime changeStart)>();

                                    foreach (var ordem2 in ListaOrdemSoldarRoboOrdenada)
                                    {
                                        if (ordem2.Programada == false)
                                        {
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem2.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                listaResultadosTestes.Add((ordem2.Record, ordem2.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                        }
                                    }

                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        dimc = tempoFim;

                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim));
                                        quantidadeRestanteRobo1Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item);

                                        i = 0;

                                    }
                                }


                                else if (quantidadeRestanteRobo1Mesa2 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 1 MESA 2");

                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int OrdersId, string OrderNo, DateTime changeStart)>();

                                    foreach (var ordem2 in ListaOrdemSoldarRoboOrdenada)
                                    {
                                        if (ordem2.Programada == false)  // Verifica se a ordem não está programada
                                        {
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem2.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                listaResultadosTestes.Add((ordem2.Record, ordem2.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                        }
                                    }

                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        dimc = tempoFim;

                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim));
                                        quantidadeRestanteRobo1Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item);

                                        i = 0;



                                    }
                                }


                                else if (quantidadeRestanteRobo2Mesa1 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 2 MESA 1");

                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int OrdersId, string OrderNo, DateTime changeStart)>();

                                    foreach (var ordem2 in ListaOrdemSoldarRoboOrdenada)
                                    {
                                        if (ordem2.Programada == false)
                                        {
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem2.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                listaResultadosTestes.Add((ordem2.Record, ordem2.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                        }
                                    }

                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        dimc = tempoFim;

                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim));
                                        quantidadeRestanteRobo2Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item);

                                        i = 0;

                                    }
                                }


                                else if (quantidadeRestanteRobo2Mesa2 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 2 MESA 2");

                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int OrdersId, string OrderNo, DateTime changeStart)>();

                                    foreach (var ordem2 in ListaOrdemSoldarRoboOrdenada)
                                    {
                                        if (ordem2.Programada == false)
                                        {
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem2.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                listaResultadosTestes.Add((ordem2.Record, ordem2.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                        }
                                    }

                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        dimc = tempoFim;

                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim));
                                        quantidadeRestanteRobo2Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item);

                                        i = 0;

                                    }
                                }


                                else if (quantidadeRestanteRobo3Mesa1 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 3 MESA 1");

                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int OrdersId, string OrderNo, DateTime changeStart)>();

                                    foreach (var ordem2 in ListaOrdemSoldarRoboOrdenada)
                                    {
                                        if (ordem2.Programada == false)
                                        {
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem2.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                listaResultadosTestes.Add((ordem2.Record, ordem2.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                        }
                                    }

                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        dimc = tempoFim;

                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim));
                                        quantidadeRestanteRobo3Mesa1 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item);

                                        i = 0;

                                    }
                                }


                                else if (quantidadeRestanteRobo3Mesa2 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 3 MESA 2");

                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int OrdersId, string OrderNo, DateTime changeStart)>();

                                    foreach (var ordem2 in ListaOrdemSoldarRoboOrdenada)
                                    {
                                        if (ordem2.Programada == false)
                                        {
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem2.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                listaResultadosTestes.Add((ordem2.Record, ordem2.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                        }
                                    }

                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        dimc = tempoFim;

                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim));
                                        quantidadeRestanteRobo3Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item);

                                        i = 0;

                                    }
                                }


                                else if (quantidadeRestanteRobo4Mesa1 <= 1)
                                {
                                    int recursoIDProgramado = preactor.PlanningBoard.GetResourceNumber("ROBO 4 MESA 1");

                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int OrdersId, string OrderNo, DateTime changeStart)>();

                                    foreach (var ordem2 in ListaOrdemSoldarRoboOrdenada)
                                    {
                                        if (ordem2.Programada == false)
                                        {
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem2.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                listaResultadosTestes.Add((ordem2.Record, ordem2.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                        }
                                    }

                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        dimc = tempoFim;

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

                                    List<(int OrdersId, string OrderNo, DateTime changeStart)> listaResultadosTestes = new List<(int OrdersId, string OrderNo, DateTime changeStart)>();

                                    foreach (var ordem2 in ListaOrdemSoldarRoboOrdenada)
                                    {
                                        if (ordem2.Programada == false)
                                        {
                                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem2.Record, recursoIDProgramado, dimc);

                                            if (resultadoTeste.HasValue)
                                            {
                                                listaResultadosTestes.Add((ordem2.Record, ordem2.OrderNo, resultadoTeste.Value.ChangeStart));
                                            }
                                        }
                                    }

                                    if (listaResultadosTestes.Count > 0)
                                    {
                                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

                                        preactor.PlanningBoard.PutOperationOnResource(resultadoMinimo.OrdersId, recursoIDProgramado, resultadoMinimo.changeStart);

                                        tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", resultadoMinimo.OrdersId);

                                        sequenciamentoOperacoes.Add((resultadoMinimo.OrdersId, resultadoMinimo.OrderNo, recursoIDProgramado, dimc, tempoFim));

                                        dimc = tempoFim;

                                        recursosProgramados.Add(recursoIDProgramado);

                                        tabelaOrdensRecurso.Add((recursoIDProgramado, resultadoMinimo.OrderNo, resultadoMinimo.OrdersId, tempoFim));
                                        quantidadeRestanteRobo4Mesa2 = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == resultadoMinimo.OrderNo && r.Programada == false);
                                        tabelaOrdensRecurso.Remove(item);

                                        i = 0;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return 0;
        }

        private static void GetResources(IPreactor preactor, IList<Resources> ListaRecursos)
        {
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

        private static void GetOrders(IPreactor preactor, IList<Orders> ListaOrders)
        {
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
                // Inicializa a variável de controle
                Ord.Programada = false;
                Ord.RecursoRequerido = -1;
                Ord.OrdenacaoPeca = preactor.ReadFieldInt("Orders", "Numerical Attribute 1", OrdersRecord);
                Ord.ValorOrdenacao = 10000;

                ListaOrders.Add(Ord);
            }
        }

    }
}