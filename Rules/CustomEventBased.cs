using Preactor;
using Preactor.Interop.PreactorObject;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace NativeRules
{
    [Guid("b94bd19e-a4cb-498c-8ed0-aa033a1aae93")]
    [ComVisible(true)]
    public interface ICustomEventBased
    {
        int EventBasedRule(ref PreactorObj preactorComObject, ref object pespComObject);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("b0ab54ad-a40a-435b-8b53-a4aff0c7cc6f")]
    public class CustomEventBased : ICustomEventBased
    {
        IPlanningBoard planningBoard;
        IPreactor preactor;
        List<Resources> resources = new List<Resources>();
        public int EventBasedRule(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            preactor = PreactorFactory.CreatePreactorObject(preactorComObject);
            planningBoard = preactor.PlanningBoard;
            LoadResources();
            if (planningBoard == null)
            {
                MessageBox.Show("This Rule must be run from the Sequencer");
                return 0;
            } // if the planning board wasn't available

            int ResourceRecord;
            string QName;
            preactor.DisplayStatus("Sequenciamento", "Sequenciando");
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
            } // while there is another event
            preactor.DestroyStatus();
            return 0;
        }

        private void ScheduleOperations(IPreactor preactor, string QName, int ResourceRecord, DateTime TestEventTime)
        {
            IPlanningBoard planningBoard = preactor.PlanningBoard;

            int CurrentRank = 1;
            int OpRecord = 0;
            bool ResourceFree = planningBoard.IsResourceFree(ResourceRecord, TestEventTime.AddDays(planningBoard.SchedulingAccuracy));
            if (ResourceFree)
            {
                planningBoard.RankQueueByPreferredSequence(QName, ResourceRecord, TestEventTime);
            }
            while (planningBoard.GetOperationInQueue(QName, CurrentRank, ref OpRecord) && ResourceFree)
            {
                if (planningBoard.GetOperationLocateState(OpRecord) == false)
                {
                    CurrentRank++;
                    continue;
                }
                var TestOpResults = planningBoard.TestOperationOnResource(OpRecord, ResourceRecord, TestEventTime);
                if (!TestOpResults.HasValue)
                {
                    CurrentRank++;
                    continue;
                }
                if (TestOpResults.Value.ChangeStart <= TestEventTime.AddDays(planningBoard.SchedulingAccuracy))
                {
                    int bestResource = FindBestResource(preactor, OpRecord, TestEventTime, TestOpResults.Value, ResourceRecord);
                    if (bestResource == ResourceRecord)
                    {
                        planningBoard.PutOperationOnResource(OpRecord, ResourceRecord, TestOpResults.Value.ChangeStart);
                    }
                    else
                    {
                        CurrentRank++;
                        continue;
                    }
                } // if the operation could start now
                else
                    CurrentRank++; // increment the rank so that we test the next job in the queue

                // is the resource still free at this time?
                ResourceFree = planningBoard.IsResourceFree(ResourceRecord, TestEventTime.AddDays(planningBoard.SchedulingAccuracy));
            } // while there is another operation in the queue
        } // End of ScheduleOperations
        private int FindBestResource(IPreactor preactor, int OpRecord, DateTime TestEventTime, OperationTimes currentOpTimes, int CurrentResource)
        {
            IPlanningBoard planningBoard = preactor.PlanningBoard;
            try
            {
                MatrixDimensions dimensions = preactor.MatrixFieldSize("Orders", "Resource Data", OpRecord);
                TimeSpan bestSetup = currentOpTimes.ProcessStart - currentOpTimes.ChangeStart;
                TimeSpan currentSetup = bestSetup;
                int bestResourceRecord = 0;
                if (dimensions.X > 0)
                {
                    for (int i = 1; i <= dimensions.X; i++)
                    {
                        int resourceNumber = preactor.ReadFieldInt("Orders", "Resource Data", OpRecord, i);
                        Resources resource = resources.Where(x => x.Number == resourceNumber).FirstOrDefault();
                        int resourceRecord = resource.Record;
                        var TestOpResults = planningBoard.TestOperationOnResource(OpRecord, resourceRecord, TestEventTime.AddDays(planningBoard.SchedulingAccuracy));
                        if (TestOpResults.HasValue && TestOpResults.Value.ProcessEnd <= currentOpTimes.ProcessEnd)
                        {
                            TimeSpan newSetup = TestOpResults.Value.ProcessStart - TestOpResults.Value.ChangeStart;
                            if (newSetup.TotalMinutes < bestSetup.TotalMinutes)
                            {
                                bestSetup = newSetup;
                                bestResourceRecord = resourceRecord;
                            }
                        }
                    }
                }
                if (bestResourceRecord > 0)
                    return bestResourceRecord;
                else
                    return CurrentResource;
            }
            catch (Exception)
            {
                int resourceGroup = preactor.ReadFieldInt("Orders", "Resource Group", OpRecord);
                TimeSpan bestSetup = currentOpTimes.ProcessStart - currentOpTimes.ChangeStart;
                TimeSpan currentSetup = bestSetup;
                int bestResourceRecord = 0;
                if (resourceGroup > 0)
                {
                    int resourceGroupRecord = preactor.FindMatchingRecord("Resource Group", "Number", 0, resourceGroup);
                    MatrixDimensions dimensions = preactor.MatrixFieldSize("Resource Group", "Resources", resourceGroupRecord);
                    if (dimensions.X > 0)
                    {
                        for (int i = 1; i <= dimensions.X; i++)
                        {
                            int resourceNumber = preactor.ReadFieldInt("Resource Group", "Resources", resourceGroupRecord, i);
                            Resources resource = resources.Where(x => x.Number == resourceNumber).FirstOrDefault();
                            int resourceRecord = resource.Record;
                            var TestOpResults = planningBoard.TestOperationOnResource(OpRecord, resourceRecord, TestEventTime.AddDays(planningBoard.SchedulingAccuracy));
                            if (TestOpResults.HasValue && TestOpResults.Value.ProcessEnd <= currentOpTimes.ProcessEnd)
                            {
                                TimeSpan newSetup = TestOpResults.Value.ProcessStart - TestOpResults.Value.ChangeStart;
                                if (newSetup.TotalMinutes < bestSetup.TotalMinutes)
                                {
                                    bestSetup = newSetup;
                                    bestResourceRecord = resourceRecord;
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (var resource in resources)
                    {
                        int resourceRecord = resource.Record;
                        var TestOpResults = planningBoard.TestOperationOnResource(OpRecord, resourceRecord, TestEventTime.AddDays(planningBoard.SchedulingAccuracy));
                        if (TestOpResults.HasValue && TestOpResults.Value.ProcessEnd <= currentOpTimes.ProcessEnd)
                        {
                            TimeSpan newSetup = TestOpResults.Value.ProcessStart - TestOpResults.Value.ChangeStart;
                            if (newSetup.TotalMinutes < bestSetup.TotalMinutes)
                            {
                                bestSetup = newSetup;
                                bestResourceRecord = resourceRecord;
                            }
                        }
                    }
                }

                if (bestResourceRecord > 0)
                    return bestResourceRecord;
                else
                    return CurrentResource;
            }

        }
        private void LoadResources()
        {
            int records = preactor.RecordCount("Resources");
            for (int record = 1; record <= records; record++)
            {
                int number = preactor.ReadFieldInt("Resources", "Number", record);
                string name = preactor.ReadFieldString("Resources", "Name", record);
                Resources resource = new Resources() { Record = record, Number = number, Name = name };
                resources.Add(resource);
            }
        }
        private class Resources
        {
            public int Record;
            public int Number;
            public string Name;
        }
    }
}
