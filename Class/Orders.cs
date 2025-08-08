using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NativeRules
{
    class Orders
    {
        public int Record { get; set; }                             // OrdersId
        public string OrderNo { get; set; }
        public string PartNo { get; set; }
        public string OpNo { get; set; }
        public string OperationName { get; set; }
        public int Resource { get; set; }
        public DateTime SetupStart { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public DateTime DueDate { get; set; }
        public Boolean Programada { get; set; }                     // Programda = true -> ordem programada pela regra | Programda= false ordem não programada pela regra (var de controle)
        public int RecursoRequerido { get; set; }                   // RecursoRequerido = recurso que a ordem precisa para ser executada (var de controle)
        public int OrdenacaoPeca { get; set; }                      // As ordens de Solda Robo foram desmenbradas em pecas, essa variável controla a ordenação das peças de uma mesma ordem de solda robo
        public int ValorOrdenacao { get; set; }                     // Valor para realizar a ordenação das ordens de produção
        public int tentativasSequenciamento { get; set; }           // As ordens de Solda Robo foram desmenbradas em pecas, essa variável controla a ordenação das peças de uma mesma ordem de solda robo
        public DateTime? MaxEndTime { get; set; }
    }
}
