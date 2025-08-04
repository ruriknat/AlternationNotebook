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
//            // -- // ok // 2.4 - Corrigir tempo de operação para que seja possivel realizar 1 operacao de forma alternada no recurso Robo N
//            // -- // nk // 2.5 - Consolidar/Agrupar operações SOLDAR ROBO de uma mesma ordem
//            // nk // 3 - Sequenciar as operações posteriores a solda robo

//            // ============================================================================================================
//            // Parte 0: criar lista de ordens, recursos e agrupar por recurso
//            // ============================================================================================================

//            IList<Orders> listaOrders = new List<Orders>();
//            IList<Resources> listaRecursos = new List<Resources>();

//            GetOrders(preactor, listaOrders);
//            GetResources(preactor, listaRecursos);

//            preactor.PlanningBoard.SequenceAll(SequenceAllDirection.Forwards, SequencePriority.DueDate);

//            var listaOrdemSoldarRobo = listaOrders
//                .Where(s => s.OperationName == "SOLDAR ROBO")
//                .ToList();

//            foreach (var registoOrdem in listaOrdemSoldarRobo)
//            {
//                int previousRecord = preactor.PlanningBoard.GetPreviousOperation(registoOrdem.Record, 1);

//                if (previousRecord > 0)
//                {
//                    preactor.PlanningBoard.UnallocateOperation(previousRecord, OperationSelection.SubsequentOperations);
//                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.ThisOperation);
//                }
//                else
//                {
//                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.BiDirectionalOperations);
//                    preactor.PlanningBoard.UnallocateOperation(registoOrdem.Record, OperationSelection.ThisOperation);
//                }
//            }

//            // --------------------------------------------- Excluir daqui

//            return 0;

//        }

//        public int SelectedResources(ref PreactorObj preactorComObject, ref object pespComObject)
//        {
//            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

//            IList<Orders> listaOrders = new List<Orders>();
//            IList<Resources> listaRecursos = new List<Resources>();

//            GetOrders(preactor, listaOrders);
//            GetResources(preactor, listaRecursos);

//            var listaOrdemSoldarRobo = listaOrders
//                .Where(s => s.OperationName == "SOLDAR ROBO")
//                .OrderBy(x => x.DueDate)
//                .ToList();

//            // ============================================================================================================
//            // Parte 0: criar lista de ordens, recursos e agrupar por recurso
//            // ============================================================================================================


//            // --------------------------------------------- Até aqui excluir daqui

//            var listaRecursoRoboSolda = listaRecursos
//                .Where(r => r.ResourceName != null &&
//                            r.ResourceName.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0 &&
//                            r.Attribute4.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0)
//                .OrderBy(r => r.ResourceName)
//                .ToList();

//            var listaGrupoRecursoRoboSolda = listaRecursoRoboSolda
//                .GroupBy(r => r.Attribute4)
//                .Select(s => new { Attribute4 = s.Key })
//                .ToList();

//            var listaOrdensAgrupadasPorOrderNo = listaOrdemSoldarRobo
//                .GroupBy(o => o.OrderNo)
//                .ToList();

//            int valorOrdenacaoCounter = 0;

//            foreach (var grupo in listaOrdensAgrupadasPorOrderNo)
//            {
//                var listaOrdensDoGrupoOrdenadas = grupo.OrderBy(o => o.DueDate).ToList();

//                foreach (var ordem in listaOrdensDoGrupoOrdenadas)
//                {
//                    var ordensParaAtualizar = listaOrdemSoldarRobo
//                        .Where(o => o.OrderNo == ordem.OrderNo)
//                        .ToList();

//                    foreach (var ordemOriginal in ordensParaAtualizar)
//                    {
//                        ordemOriginal.ValorOrdenacao = valorOrdenacaoCounter;
//                    }
//                    valorOrdenacaoCounter++;
//                }
//            }

//            var ListaOrdemSoldarRoboOrdenada = listaOrdemSoldarRobo
//                .OrderBy(x => x.OrdenacaoPeca)
//                .ThenBy(x => x.ValorOrdenacao)
//                .ThenBy(x => x.DueDate)
//                .ToList();

//            // ============================================================================================================
//            // Parte 1: Inicia o sequenciamento das operações SOLDAR ROBO
//            // ============================================================================================================

//            var queueOrdens = new Queue<Orders>(ListaOrdemSoldarRoboOrdenada);
//            DateTime dimc = preactor.PlanningBoard.TerminatorTime;

//            // Dicionário para armazenar o estado de cada robô
//            Dictionary<string, (bool mesa1, bool mesa2, string ordmeMesa1, string ordmeMesa2, int? quantidadeOrdemMesa1, int? quantidadeOrdemMesa2, DateTime tempo)> roboEstados = new Dictionary<string, (bool, bool, string, string, int?, int?, DateTime)>
//            {
//                { "ROBO 1", (false, false, "", "", null, null, preactor.PlanningBoard.TerminatorTime) },
//                { "ROBO 2", (false, false, "", "", null, null, preactor.PlanningBoard.TerminatorTime) },
//                { "ROBO 3", (false, false, "", "", null, null, preactor.PlanningBoard.TerminatorTime) },
//                { "ROBO 4", (false, false, "", "", null, null, preactor.PlanningBoard.TerminatorTime) }
//            };

//            List<int> recursosProgramados = new List<int>();
//            List<(int recursoId, string OrderNo, int ordemId, DateTime changeStart)> tabelaOrdensRecurso = new List<(int, string, int, DateTime)>();

//            for (int i = 0; i < 4; i++)
//            {


//                while (recursosProgramados.Count < listaRecursoRoboSolda.Count)
//                {
//                    var primeiraOrdem = queueOrdens.Dequeue();
//                    List<(int recursoId, string Attribute4, DateTime changeStart)> listaResultadosTestes = new List<(int recursoId, string Attribute4, DateTime changeStart)>();

//                    foreach (var recurso in listaRecursoRoboSolda.Where(r => !recursosProgramados.Contains(r.ResourceId)))
//                    {
//                        var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(primeiraOrdem.Record, recurso.ResourceId, dimc);
//                        if (resultadoTeste.HasValue)
//                        {
//                            listaResultadosTestes.Add((recurso.ResourceId, recurso.Attribute4, resultadoTeste.Value.ChangeStart));
//                        }
//                    }

//                    if (listaResultadosTestes.Count > 0)
//                    {
//                        var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();

//                        if (primeiraOrdem.RecursoRequerido == -1 && roboEstados.ContainsKey(resultadoMinimo.Attribute4))
//                        {
//                            var robo = roboEstados[resultadoMinimo.Attribute4];

//                            if (preactor.PlanningBoard.GetResourceName(resultadoMinimo.recursoId).IndexOf("mesa 1", StringComparison.OrdinalIgnoreCase) >= 0)
//                            {
//                                if (robo.tempo != preactor.PlanningBoard.TerminatorTime)
//                                {
//                                    resultadoMinimo.changeStart = robo.tempo;
//                                }
//                                preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);
//                                int quantidadeRestante = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == primeiraOrdem.OrderNo && r.Programada == false) - 1;
//                                var tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
//                                roboEstados[resultadoMinimo.Attribute4] = (true, robo.mesa2, primeiraOrdem.OrderNo, robo.ordmeMesa2, quantidadeRestante, robo.quantidadeOrdemMesa2, tempoFim);
//                                primeiraOrdem.Programada = true;

//                                var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada.Where(o => o.OrderNo == primeiraOrdem.OrderNo).ToList();
//                                foreach (var ordemOriginal in ordensParaAtualizar)
//                                {
//                                    if (primeiraOrdem.Record == ordemOriginal.Record)
//                                    {
//                                        ordemOriginal.Programada = primeiraOrdem.Programada;
//                                    }
//                                    ordemOriginal.RecursoRequerido = resultadoMinimo.recursoId;
//                                }

//                                // Variável para looping
//                                recursosProgramados.Add(resultadoMinimo.recursoId);
//                            }
//                            else if (preactor.PlanningBoard.GetResourceName(resultadoMinimo.recursoId).IndexOf("mesa 2", StringComparison.OrdinalIgnoreCase) >= 0)
//                            {
//                                if (robo.tempo != preactor.PlanningBoard.TerminatorTime)
//                                {
//                                    resultadoMinimo.changeStart = robo.tempo;
//                                }
//                                preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);
//                                int quantidadeRestante = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == primeiraOrdem.OrderNo && r.Programada == false) - 1;
//                                var tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
//                                roboEstados[resultadoMinimo.Attribute4] = (robo.mesa1, true, robo.ordmeMesa1, primeiraOrdem.OrderNo, robo.quantidadeOrdemMesa1, quantidadeRestante, tempoFim);
//                                primeiraOrdem.Programada = true;

//                                var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada.Where(o => o.OrderNo == primeiraOrdem.OrderNo).ToList();
//                                foreach (var ordemOriginal in ordensParaAtualizar)
//                                {
//                                    if (primeiraOrdem.Record == ordemOriginal.Record)
//                                    {
//                                        ordemOriginal.Programada = primeiraOrdem.Programada;
//                                    }
//                                    ordemOriginal.RecursoRequerido = resultadoMinimo.recursoId;
//                                }

//                                // Variável para looping
//                                recursosProgramados.Add(resultadoMinimo.recursoId);
//                            }
//                        }
//                    }
//                    else
//                    {
//                        queueOrdens.Enqueue(primeiraOrdem);
//                    }

//                }

//                // ============================================================================================================
//                // Parte dois, continuar o sequenciamento das operações SOLDAR ROBO com a mesma ordem já sequenciada
//                // ============================================================================================================

//                var ordensComRecursoAlocado = roboEstados
//                    .SelectMany(kv => { var estado = kv.Value; return new[] { estado.ordmeMesa1, estado.ordmeMesa2 }; })
//                    .Where(orderNo => !string.IsNullOrWhiteSpace(orderNo))
//                    .Distinct()
//                    .ToList();

//                var operacoesComRecursoAlocado = ListaOrdemSoldarRoboOrdenada
//                    .Where(ordem => !ordem.Programada && ordensComRecursoAlocado.Contains(ordem.OrderNo))
//                    .ToList();

//                foreach (var ordem in operacoesComRecursoAlocado)
//                {
//                    foreach (var recursoSelecinado in listaRecursoRoboSolda)
//                    {
//                        if (ordem.RecursoRequerido == recursoSelecinado.ResourceId)
//                        {
//                            dimc = roboEstados[recursoSelecinado.Attribute4].tempo;
//                            int quantidadeRestante = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == ordem.OrderNo && r.Programada == false) - 1;

//                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(ordem.Record, recursoSelecinado.ResourceId, dimc);
//                            if (resultadoTeste.HasValue)
//                            {
//                                preactor.PlanningBoard.PutOperationOnResource(ordem.Record, recursoSelecinado.ResourceId, resultadoTeste.Value.ChangeStart);
//                                ordem.Programada = true;
//                                var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada.Where(o => o.Record == ordem.Record).ToList();
//                                foreach (var ordemOriginal in ordensParaAtualizar)
//                                {
//                                    ordemOriginal.Programada = ordem.Programada;
//                                    ordemOriginal.RecursoRequerido = recursoSelecinado.ResourceId;
//                                }

//                                var robo = roboEstados[recursoSelecinado.Attribute4];

//                                if (preactor.PlanningBoard.GetResourceName(recursoSelecinado.ResourceId).IndexOf("mesa 1", StringComparison.OrdinalIgnoreCase) >= 0)
//                                {
//                                    roboEstados[recursoSelecinado.Attribute4] = (true, robo.mesa2, ordem.OrderNo, robo.ordmeMesa2, quantidadeRestante, robo.quantidadeOrdemMesa2, preactor.ReadFieldDateTime("Orders", "End Time", ordem.Record));
//                                }
//                                else if (preactor.PlanningBoard.GetResourceName(recursoSelecinado.ResourceId).IndexOf("mesa 2", StringComparison.OrdinalIgnoreCase) >= 0)
//                                {
//                                    roboEstados[recursoSelecinado.Attribute4] = (robo.mesa1, true, robo.ordmeMesa1, ordem.OrderNo, robo.quantidadeOrdemMesa1, quantidadeRestante, preactor.ReadFieldDateTime("Orders", "End Time", ordem.Record));
//                                }

//                            }
//                        }
//                    }
//                }

//                // ============================================================================================================
//                // Parte três, continuar o sequenciamento das operações SOLDAR ROBO com as demais ordne ao acabar uma ordem
//                // ============================================================================================================

//                foreach (var estado in roboEstados)
//                {
//                    var robo = estado.Key;
//                    var dados = estado.Value;

//                    if (dados.quantidadeOrdemMesa1.HasValue && dados.quantidadeOrdemMesa1.Value == 0 && dados.ordmeMesa1 != "")
//                    {

//                        var primeiraOrdem = queueOrdens.Dequeue();
//                        List<(int recursoId, string Attribute4, DateTime changeStart)> listaResultadosTestes = new List<(int recursoId, string Attribute4, DateTime changeStart)>();
//                        foreach (var recurso in listaRecursoRoboSolda.Where(r => !recursosProgramados.Contains(r.ResourceId)))
//                        {
//                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(primeiraOrdem.Record, recurso.ResourceId, dimc);
//                            if (resultadoTeste.HasValue)
//                            {
//                                listaResultadosTestes.Add((recurso.ResourceId, recurso.Attribute4, resultadoTeste.Value.ChangeStart));
//                            }
//                        }
//                        if (listaResultadosTestes.Count > 0)
//                        {
//                            var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();
//                            if (primeiraOrdem.RecursoRequerido == -1 && roboEstados.ContainsKey(resultadoMinimo.Attribute4))
//                            {
//                                if (preactor.PlanningBoard.GetResourceName(resultadoMinimo.recursoId).IndexOf("mesa 1", StringComparison.OrdinalIgnoreCase) >= 0)
//                                {
//                                    preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);
//                                    int quantidadeRestante = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == primeiraOrdem.OrderNo && r.Programada == false) - 1;
//                                    var tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
//                                    roboEstados[resultadoMinimo.Attribute4] = (true, dados.mesa2, primeiraOrdem.OrderNo, dados.ordmeMesa2, quantidadeRestante, dados.quantidadeOrdemMesa2, tempoFim);
//                                    primeiraOrdem.Programada = true;
//                                    var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada.Where(o => o.OrderNo == primeiraOrdem.OrderNo).ToList();
//                                    foreach (var ordemOriginal in ordensParaAtualizar)
//                                    {
//                                        if (primeiraOrdem.Record == ordemOriginal.Record)
//                                        {
//                                            ordemOriginal.Programada = primeiraOrdem.Programada;
//                                        }
//                                        ordemOriginal.RecursoRequerido = resultadoMinimo.recursoId;
//                                    }

//                                    // Variável para looping
//                                    recursosProgramados.Add(resultadoMinimo.recursoId);
//                                }
//                                else if (preactor.PlanningBoard.GetResourceName(resultadoMinimo.recursoId).IndexOf("mesa 2", StringComparison.OrdinalIgnoreCase) >= 0)
//                                {

//                                    preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);
//                                    int quantidadeRestante = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == primeiraOrdem.OrderNo && r.Programada == false) - 1;
//                                    var tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
//                                    roboEstados[resultadoMinimo.Attribute4] = (dados.mesa1, true, dados.ordmeMesa1, primeiraOrdem.OrderNo, dados.quantidadeOrdemMesa1, quantidadeRestante, tempoFim);
//                                    primeiraOrdem.Programada = true;

//                                    var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada.Where(o => o.OrderNo == primeiraOrdem.OrderNo).ToList();
//                                    foreach (var ordemOriginal in ordensParaAtualizar)
//                                    {
//                                        if (primeiraOrdem.Record == ordemOriginal.Record)
//                                        {
//                                            ordemOriginal.Programada = primeiraOrdem.Programada;
//                                        }
//                                        ordemOriginal.RecursoRequerido = resultadoMinimo.recursoId;
//                                    }

//                                }
//                            }
//                            else
//                            {
//                                queueOrdens.Enqueue(primeiraOrdem);
//                            }
//                        }
//                    }


//                    else if (dados.quantidadeOrdemMesa2.HasValue && dados.quantidadeOrdemMesa2.Value == 0 && dados.ordmeMesa1 != "")
//                    {

//                        var primeiraOrdem = queueOrdens.Dequeue();
//                        List<(int recursoId, string Attribute4, DateTime changeStart)> listaResultadosTestes = new List<(int recursoId, string Attribute4, DateTime changeStart)>();
//                        foreach (var recurso in listaRecursoRoboSolda.Where(r => !recursosProgramados.Contains(r.ResourceId)))
//                        {
//                            var resultadoTeste = preactor.PlanningBoard.TestOperationOnResource(primeiraOrdem.Record, recurso.ResourceId, dimc);
//                            if (resultadoTeste.HasValue)
//                            {
//                                listaResultadosTestes.Add((recurso.ResourceId, recurso.Attribute4, resultadoTeste.Value.ChangeStart));
//                            }
//                        }
//                        if (listaResultadosTestes.Count > 0)
//                        {
//                            var resultadoMinimo = listaResultadosTestes.OrderBy(r => r.changeStart).First();
//                            if (primeiraOrdem.RecursoRequerido == -1 && roboEstados.ContainsKey(resultadoMinimo.Attribute4))
//                            {
//                                if (preactor.PlanningBoard.GetResourceName(resultadoMinimo.recursoId).IndexOf("mesa 1", StringComparison.OrdinalIgnoreCase) >= 0)
//                                {
//                                    preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);
//                                    int quantidadeRestante = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == primeiraOrdem.OrderNo && r.Programada == false) - 1;
//                                    var tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
//                                    roboEstados[resultadoMinimo.Attribute4] = (true, dados.mesa2, primeiraOrdem.OrderNo, dados.ordmeMesa2, quantidadeRestante, dados.quantidadeOrdemMesa2, tempoFim);
//                                    primeiraOrdem.Programada = true;
//                                    var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada.Where(o => o.OrderNo == primeiraOrdem.OrderNo).ToList();
//                                    foreach (var ordemOriginal in ordensParaAtualizar)
//                                    {
//                                        if (primeiraOrdem.Record == ordemOriginal.Record)
//                                        {
//                                            ordemOriginal.Programada = primeiraOrdem.Programada;
//                                        }
//                                        ordemOriginal.RecursoRequerido = resultadoMinimo.recursoId;
//                                    }

//                                    // Variável para looping
//                                    recursosProgramados.Add(resultadoMinimo.recursoId);
//                                }
//                                else if (preactor.PlanningBoard.GetResourceName(resultadoMinimo.recursoId).IndexOf("mesa 2", StringComparison.OrdinalIgnoreCase) >= 0)
//                                {

//                                    preactor.PlanningBoard.PutOperationOnResource(primeiraOrdem.Record, resultadoMinimo.recursoId, resultadoMinimo.changeStart);
//                                    int quantidadeRestante = ListaOrdemSoldarRoboOrdenada.Count(r => r.OrderNo == primeiraOrdem.OrderNo && r.Programada == false) - 1;
//                                    var tempoFim = preactor.ReadFieldDateTime("Orders", "End Time", primeiraOrdem.Record);
//                                    roboEstados[resultadoMinimo.Attribute4] = (dados.mesa1, true, dados.ordmeMesa1, primeiraOrdem.OrderNo, dados.quantidadeOrdemMesa1, quantidadeRestante, tempoFim);
//                                    primeiraOrdem.Programada = true;

//                                    var ordensParaAtualizar = ListaOrdemSoldarRoboOrdenada.Where(o => o.OrderNo == primeiraOrdem.OrderNo).ToList();
//                                    foreach (var ordemOriginal in ordensParaAtualizar)
//                                    {
//                                        if (primeiraOrdem.Record == ordemOriginal.Record)
//                                        {
//                                            ordemOriginal.Programada = primeiraOrdem.Programada;
//                                        }
//                                        ordemOriginal.RecursoRequerido = resultadoMinimo.recursoId;
//                                    }

//                                }
//                            }
//                            else
//                            {
//                                queueOrdens.Enqueue(primeiraOrdem);
//                            }
//                        }
//                    }

//                }
//            }
//            return 0;
//        }

//        private static void GetResources(IPreactor preactor, IList<Resources> ListaRecursos)
//        {
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
//                // Inicializa a variável de controle
//                Ord.Programada = false;
//                Ord.RecursoRequerido = -1;
//                Ord.OrdenacaoPeca = preactor.ReadFieldInt("Orders", "Numerical Attribute 1", OrdersRecord);
//                Ord.ValorOrdenacao = 10000;

//                ListaOrders.Add(Ord);
//            }
//        }

//    }
//}