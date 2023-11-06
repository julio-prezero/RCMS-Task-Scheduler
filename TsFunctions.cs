using System;
using Microsoft.Win32.TaskScheduler;

namespace RCMSTaskScheduler.GlobalCode
{
    public class TsFunctions
    {
        public static TaskService RCMSTaskService = new TaskService();
        public static string TaskNamePrefix()
        {
            return "RCMS®";
        }
        public static string Action()
        {
            return @"C:\Docs\RCMS\RCMSTaskScheduler.exe";
        }
        public static string Description()
        {
            return "RCMS® Report-Scheduler Job. ";
        }
         public static bool isLocal()
        {
            if (Environment.MachineName == "RCMS-DEV" | Environment.MachineName == "RCMS01")
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public static void CreateTaskRunOnce(string profileID, DateTime StartDateTime, string desc)
        {
            StartTaskService();
            TaskDefinition td = RCMSTaskService.NewTask();
            td.RegistrationInfo.Author = "";//GetUserName();
            td.RegistrationInfo.Description = desc;
            if (!isLocal())
            {
                //S4U means it will run task whether itadmin is logged in or not...
                //This is mandatory if the task is run on the server...
                //This isn't needed if running locally...
                td.Principal.LogonType = TaskLogonType.S4U;
            }

            td.Triggers.Add(new TimeTrigger() { StartBoundary = StartDateTime });
            td.Actions.Add(new ExecAction(Action(), profileID, Program.RCMSFolder));
            RCMSTaskService.RootFolder.RegisterTaskDefinition(TaskNamePrefix() + profileID, td);
            DisposeTaskService();
        }


        public static bool TestTaskExists(string TestTask)
        {
            StartTaskService();
            if (RCMSTaskService.GetTask(TestTask) != null)
            {
                DisposeTaskService();
                return true;
            }
            else
            {
                DisposeTaskService();
                return false;
            }
        }

        public static bool TaskExists(string taskname)
        {
            StartTaskService();
            if (RCMSTaskService.GetTask(taskname) != null)
            {
                DisposeTaskService();
                return true;
            }
            else
            {
                DisposeTaskService();
                return false;
            }
        }
        public static void DeleteTask(string taskname)
        {
            StartTaskService();
            if (RCMSTaskService.GetTask(taskname) != null)
            {
                RCMSTaskService.RootFolder.DeleteTask(taskname);
            }
            DisposeTaskService();
        }
        public static void StartTaskService()
        {
            if (isLocal())//Local environment....
            {
                RCMSTaskService = new TaskService();

            }
            else//server...
            {
                RCMSTaskService = new TaskService(Environment.MachineName, "itadmin", "DMRMG", MiscFunctions.GetTSPassWord());
            }
        }
        public static void DisposeTaskService()
        {
            if (RCMSTaskService != null)
            {
                RCMSTaskService.Dispose();
            }
        }
 
    }
}