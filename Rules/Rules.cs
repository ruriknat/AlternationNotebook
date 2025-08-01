using Preactor;
using Preactor.Interop.PreactorObject;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace NativeRules
{
    [Guid("827bac54-089c-4897-b966-f23a3e42b708")]
    [ComVisible(true)]
    public interface IRules
    {
        //int Run(ref PreactorObj preactorComObject, ref object pespComObject);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("f199c654-6156-4888-bb46-8688ec931584")]
    public class Rules : IRules
    {
        public int Run(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

            // TODO : Your code here

            return 0;
        }
        public int ForwardDueDate(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

            preactor.PlanningBoard.SequenceAll(SequenceAllDirection.Forwards, SequencePriority.DueDate);

            return 0;
        }

        public int ForwardPriority(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);

            preactor.PlanningBoard.SequenceAll(SequenceAllDirection.Forwards, SequencePriority.Priority);

            return 0;
        }
        public int ForwardByWeight(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);
            preactor.PlanningBoard.SequenceAll(SequenceAllDirection.Forwards, SequencePriority.Weight);
            return 0;
        }
        public int EventBasedRule(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            IPreactor preactor = PreactorFactory.CreatePreactorObject(preactorComObject);
            IPlanningBoard planningBoard = preactor.PlanningBoard;
            if (planningBoard == null)
            {
                MessageBox.Show("This Rule must be run from the Sequencer");
                return 0;
            } // if the planning board wasn't available

            int ResourceRecord;
            string QName;

            Preactor.EventDetails? EventParameters = planningBoard.NextEvent();
            while (EventParameters.HasValue)
            {
                switch (EventParameters.Value.EventType)
                {
                    case EventTypes.OperationFinished:
                        // Event Parameter 1 is the Operation record that finished
                        // Event Parameter 2 is the Resource record that became available
                        // check all resources for this event because secondary constraints may have changed

                        for (ResourceRecord = 1; ResourceRecord <= preactor.RecordCount("Resources"); ResourceRecord++)
                        {
                            QName = planningBoard.GetResourceQueueName(ResourceRecord);
                            ScheduleOperations(preactor, QName, ResourceRecord, EventParameters.Value.EventTime);

                        } // for each Resource

                        break;

                    case EventTypes.QueueChange:
                        // Event Parameter 1 is the number of the queue that changed
                        // check all resources which use this queue
                        int ResIndex = 1;
                        ResourceRecord = 0;
                        QName = planningBoard.GetQueueName(EventParameters.Value.Parameter1);
                        while (planningBoard.GetQueuesResource(QName, ResIndex, ref ResourceRecord))
                        {
                            ScheduleOperations(preactor, QName, ResourceRecord, EventParameters.Value.EventTime);

                            ResIndex++;
                        } // whilst there is another resource for this queue
                        break;

                    case EventTypes.ShiftChange:
                        // Event Parameter 2 is the Resource record that had a shift change
                        // check the resource that had the shift change
                        int QNumber = planningBoard.GetResourceQueue(EventParameters.Value.Parameter2);
                        QName = planningBoard.GetQueueName(QNumber);

                        ScheduleOperations(preactor, QName, EventParameters.Value.Parameter2,
                                           EventParameters.Value.EventTime);
                        break;
                    case EventTypes.UserEvent:
                        break;
                    default:
                        break;
                }

                EventParameters = planningBoard.NextEvent();
            } // whilst there is another event

            return 0;
        }

        private void ScheduleOperations(IPreactor preactor, string QName, int ResourceRecord, DateTime TestEventTime)
        {
            IPlanningBoard planningBoard = preactor.PlanningBoard;
            planningBoard.RankQueueBySetupTime(QName, ResourceRecord, TestEventTime, QueueRanking.Ascending);
            int CurrentRank = 1;
            int OpRecord = 0;
            bool ResourceFree = planningBoard.IsResourceFree(ResourceRecord, TestEventTime.AddDays(planningBoard.SchedulingAccuracy));
            if (ResourceFree)
            {
                planningBoard.RankQueueByPreferredSequence(QName, ResourceRecord, TestEventTime);
            }
            while (planningBoard.GetOperationInQueue(QName, CurrentRank, ref OpRecord) && ResourceFree)
            {
                var TestOpResults = planningBoard.TestOperationOnResource(OpRecord, ResourceRecord,
                    TestEventTime);
                if (!TestOpResults.HasValue)
                {
                    CurrentRank++;
                    continue;
                }
                if (TestOpResults.Value.ChangeStart <= TestEventTime.AddDays(planningBoard.SchedulingAccuracy))
                {
                    planningBoard.PutOperationOnResource(OpRecord, ResourceRecord, TestOpResults.Value.ChangeStart);
                } // if the operation could start now
                else
                    CurrentRank++; // increment the rank so that we test the next job in the queue

                // is the resource still free at this time?
                ResourceFree = planningBoard.IsResourceFree(ResourceRecord, TestEventTime.AddDays(planningBoard.SchedulingAccuracy));
            } // whilst there is another operation in the queue        }
        } // End of ScheduleOperations
    }
}
