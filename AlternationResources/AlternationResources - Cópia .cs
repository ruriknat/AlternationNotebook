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

//                Ord.OrderNo = preactor.ReadFieldString("Orders", "Order No.", OrdersRecord);
//                Ord.PartNo = preactor.ReadFieldString("Orders", "Part No.", OrdersRecord);
//                Ord.OpNo = preactor.ReadFieldString("Orders", "Op. No.", OrdersRecord);
//                Ord.OperationName = preactor.ReadFieldString("Orders", "Operation Name", OrdersRecord);
//                Ord.SetupStart = preactor.ReadFieldDateTime("Orders", "Setup Start", OrdersRecord);
//                Ord.StartTime = preactor.ReadFieldDateTime("Orders", "Start Time", OrdersRecord);
//                Ord.EndTime = preactor.ReadFieldDateTime("Orders", "End Time", OrdersRecord);
//                Ord.DueDate = preactor.ReadFieldDateTime("Orders", "Due Date", OrdersRecord);
//                Ord.Record = OrdersRecord;

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

//            // Excluir daqui


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

//                Ord.OrderNo = preactor.ReadFieldString("Orders", "Order No.", OrdersRecord);
//                Ord.PartNo = preactor.ReadFieldString("Orders", "Part No.", OrdersRecord);
//                Ord.OpNo = preactor.ReadFieldString("Orders", "Op. No.", OrdersRecord);
//                Ord.OperationName = preactor.ReadFieldString("Orders", "Operation Name", OrdersRecord);
//                Ord.SetupStart = preactor.ReadFieldDateTime("Orders", "Setup Start", OrdersRecord);
//                Ord.StartTime = preactor.ReadFieldDateTime("Orders", "Start Time", OrdersRecord);
//                Ord.EndTime = preactor.ReadFieldDateTime("Orders", "End Time", OrdersRecord);
//                Ord.DueDate = preactor.ReadFieldDateTime("Orders", "Due Date", OrdersRecord);
//                Ord.Record = OrdersRecord;

//                ListaOrders.Add(Ord);

//            }


//            // Até aqui excluir daqui


//            // == Dados Recursos ==

//            // Lista em que sera salvo os dados dos recursso
//            IList<Resources> ListaRecursos = new List<Resources>();

//            // Inicia o contador de recursos
//            int resourceRecord = 1;

//            // Coleta de dados das operações para cada uma dos OrdersId
//            for (resourceRecord = 1; resourceRecord <= preactor.RecordCount("Resources"); resourceRecord++)
//            {
//                Resources Rec = new Resources();

//                Rec.ResourceName = preactor.ReadFieldString("Resources", "Name", resourceRecord);
//                Rec.Attribute1 = preactor.ReadFieldString("Resources", "Attribute 1", resourceRecord);
//                Rec.Attribute4 = preactor.ReadFieldString("Resources", "Attribute 4", resourceRecord);
//                Rec.ResourceId = resourceRecord;

//                ListaRecursos.Add(Rec);
//            }

//            // Cria uma lista com os dados filtrando os recursso de solda robo
//            var ListaRoboResources = ListaRecursos.Where(r => r.ResourceName != null && r.ResourceName.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0 && r.Attribute4.IndexOf("ROBO", StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(r => r.ResourceName).ToList();

//            // Cria uma lista com os agrupamentos de recursos de solda robo
//            var ListaGrupoRoboResources = ListaRoboResources.GroupBy(r => r.Attribute4).Select(s => new { Attribute4 = s.Key }).ToList();

//            // == Dados Recursos ==

//            // == Inicio Dados Ordens Soldar Robo ==

//            // Ordena a lista de ordens  "SOLDAR ROBO" e ordenando por DueDate
//            var ListaOrdemSoldarRoboOrdenada = ListaOrders.Where(s => s.OperationName == "SOLDAR ROBO").OrderBy(x => x.DueDate).ToList();


//            // == Fim Dados Ordens Soldar Robo ==


//            // == Interações Solda Robo ==

//            int recurso1 = -1;
//            int recurso2 = -1;

//            foreach (var RegistroGrupoRecurso in ListaGrupoRoboResources)
//            {
//                foreach (var RegistroRecurso in ListaRoboResources)
//                {
//                    if (RegistroRecurso.Attribute4 == RegistroGrupoRecurso.Attribute4 && recurso1 == -1)
//                    {
//                        // Atribui o recurso1
//                        recurso1 = RegistroRecurso.ResourceId;
//                    }
//                    else if (RegistroRecurso.Attribute4 == RegistroGrupoRecurso.Attribute4 && recurso2 == -1 && RegistroRecurso.ResourceId != recurso1)
//                    {
//                        // Atribui o recurso2
//                        recurso2 = RegistroRecurso.ResourceId;
//                    }
//                }
//                // Iterar sobre as operacoes de "SOLDAR ROBO"
//                while (ListaOrdemSoldarRoboOrdenada.Count > 0)
//                {
//                    // Inicializa a variável de data para a programação das ordens
//                    DateTime DIMC = DateTime.MinValue;

//                    foreach (var RegistoOrdem in ListaOrdemSoldarRoboOrdenada)
//                    {
//                        // Atualiza o DIMC (Data Inicial de Mais Cedo da Ordem) 
//                        preactor.WriteField("Orders", "Earliest Start Date", RegistoOrdem.Record, DIMC);

//                        // Testa a operação nos recursos
//                        var testeop1 = preactor.PlanningBoard.TestOperationOnResource(RegistoOrdem.Record, recurso1, DIMC);
//                        var testeop2 = preactor.PlanningBoard.TestOperationOnResource(RegistoOrdem.Record, recurso2, DIMC);

//                        // Lógica para verificar qual recurso deve ser utilizado
//                        if (testeop1.HasValue && testeop2.HasValue)
//                        {
//                            if (testeop1.Value.ChangeStart <= testeop2.Value.ChangeStart)
//                            {
//                                // Aloca recurso1
//                                preactor.PlanningBoard.PutOperationOnResource(RegistoOrdem.Record, recurso1, testeop1.Value.ChangeStart);
//                            }
//                            else
//                            {
//                                // Aloca recurso2
//                                preactor.PlanningBoard.PutOperationOnResource(RegistoOrdem.Record, recurso2, testeop2.Value.ChangeStart);
//                            }
//                        }
//                        else if (!testeop1.HasValue && testeop2.HasValue)
//                        {
//                            // Aloca recurso2 se testeop1 for nulo
//                            preactor.PlanningBoard.PutOperationOnResource(RegistoOrdem.Record, recurso2, testeop2.Value.ChangeStart);
//                        }
//                        else if (testeop1.HasValue && !testeop2.HasValue)
//                        {
//                            // Aloca recurso1 se testeop2 for nulo
//                            preactor.PlanningBoard.PutOperationOnResource(RegistoOrdem.Record, recurso1, testeop1.Value.ChangeStart);
//                        }

//                        DIMC = preactor.ReadFieldDateTime("Orders", "End Time", RegistoOrdem.Record);
//                    }
//                    // Remove a ordem da lista de ordens de soldar robo desprogramada
//                    ListaOrdemSoldarRoboOrdenada.RemoveAt(0);

//                }

//            }

//            return 0;
//        }
//    }
//    }
//}
