/// <summary>
/// ©2017 Eugene Odeluga
/// This is a Windows Update API wrapper for Windows PowerShell
///  
/// It wraps some of the functions of the WU API contained in WUApiLib.dll into a PS Cmdlet module called WUWrapper
/// 
/// The Cmdlets allow the management of Windows Update through PowerShell for local and remote machines (needs admin)
///  
/// My first time making a PowerShell module in C# (my first C# project too :D ) and still WIP
///  
/// Current Cmdlet list:
///     Get-WUUpdateHistory
///     Receive-WUUpdate
/// 
/// </summary>
namespace EIO.WUWrapper {
    using System;
    using System.Management.Automation;
    using WUApiLib;
    using Containers;
    
    // Cmdlets
    #region Get-WUUpdateHistory

    /// <summary>
    /// Retrieves the history of Windows Updates installed on a local or the specified remote computer
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "WUUpdateHistory")]
    // The cmdlet outputs UpdateHistory object to PowerShell
    [OutputType(typeof(UpdateHistory))]

    public class GetUpdateHistoryCmdlet : Cmdlet {

        private static IUpdateHistoryEntryCollection history;
        private IUpdateHistoryEntryCollection _updateHistoryResults;
        private string _computerName;   // Store the names of the computers to work on from command line
        private string _findUpdate; // Store the name of a specific update to search for

        #region Parameters
        /// <summary>
        /// ComputerName PS parameter to specify the name of the computer or computers to retrieve update history from
        /// This a pipeline(able) parameter and also can set parameter by taking from a pipelined cmdlet
        /// that outputs a property with the same name
        /// </summary>
        [Parameter(
            Position = 0,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true)
            ValidateNotNull]
        public string ComputerName {
            get { return _computerName; }
            // Set parameter to localhost if left blank
            set { if (_computerName == "") {
                    _computerName = Environment.MachineName;
                } else { _computerName = value; }
            }
        }

        /// <summary>
        /// FindUpdate PS parameter is used to specify a particular update to search the history for
        /// TODO: Currently has no effect
        /// </summary>
        [Parameter(Position = 1)]
        public string FindUpdate {
            get { return _findUpdate; }
            set { _findUpdate = value; }
        }
        #endregion Parameters

        #region Class Methods

        protected override void BeginProcessing() {
            base.BeginProcessing();
        }

        protected override void ProcessRecord() {
            base.ProcessRecord();
            // Get the specified computer's Windows Updates history
            _updateHistoryResults = GetUpdateHistory(_computerName);

            foreach (IUpdateHistoryEntry result in _updateHistoryResults) {
                UpdateHistory output = new UpdateHistory(result);
                // Output the UpdateHistory object to PowerShell console
                WriteObject(output);
            }
        }

        protected override void EndProcessing() {
            base.EndProcessing();
            // Clean up
            GC.Collect();
        }

        private IUpdateHistoryEntryCollection GetUpdateHistory(string _computerName) {
            Type type = Type.GetTypeFromProgID("Microsoft.Update.Session", _computerName, true);
            UpdateSession session = (UpdateSession)Activator.CreateInstance(type);
            IUpdateSearcher searcher = session.CreateUpdateSearcher();
            int count = searcher.GetTotalHistoryCount();
            history = searcher.QueryHistory(0, count);
            
            return history;
        }

        #endregion Class Methods
    }

    #endregion Get-WUUpdateHistory

    #region Receive-WUUpdate

    /// <summary>
    /// Downloads a specific specified update or all available if not speciifed 
    /// </summary>
    [Cmdlet(VerbsCommunications.Receive, "WUUpdate")]
    // Returns installation result
    [OutputType(typeof(IInstallationResult))]
    public class ReceiveUpdate : Cmdlet {

        #region Parameters
        private string _computerName;   // Store the names of the computers to work on from command line
        private string _updateName;   // Store the name of the specified update from command line

        // ComputerName
        [Parameter(
            Position = 0,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true)
            ValidateNotNull]
        public string ComputerName {
            get { return _computerName; }
            // Set parameter to localhost if left blank
            set {
                if (_computerName == "") {
                    _computerName = Environment.MachineName;
                }
                else { _computerName = value; }
            }
        }
        // UpdateName
        [Parameter(Position = 1)
            ValidateNotNull]
        public string UpdateName {
            get { return _updateName; }
        }
    }

        #endregion Parameters

        #region Class Methods
        protected override void BeginProcessing() { }

        #endregion Class Methods

    #endregion Receive-WUUpdate
}

// Cmdlet objects
#region Containers
/// <summary>
/// Objects used by the cmdlets
/// </summary>
namespace EIO.WUWrapper.Containers {
    using Enums;
    using System.Text.RegularExpressions;

    public class UpdateHistory {
        // This regex pattern is used to match MS KB numbers from the name of an update
        private static string pattern = @"KB?\d+";
        private static Regex regex = new Regex (pattern);

        // Class getter variables
        private static string _name;
        private static int _operation;
        private static int _status;
        private static string _description;
        private static string _kb;
        private static string _date;

        // Constructor
        public UpdateHistory(WUApiLib.IUpdateHistoryEntry result) {
            _name = result.Title;
            _operation = (int)result.Operation;
            _status = (int)result.ResultCode;
            _description = result.Description;
            // Use regex to capture KB article number from title
            _kb = regex.Match(_name).Value;
            _date = (result.Date).ToString();
        }

        // Convert operation codes to text equivalent
        private static string GetUpdateOperation (int operation) {
            /// <summary>
            /// Uses Windows Update operation code as input to return 
            /// text string name of operation
            /// </summary>
            switch (operation) {
                default:
                    return "N/A";
                case (int)enums.OperationCode.Installed:
                    return enums.OperationCode.Installed.ToString();
                case (int)enums.OperationCode.Uninstalled:
                    return enums.OperationCode.Uninstalled.ToString();
            }
        }

        // Convert operation result codes to text equivalent
        private static string GetUpdateOperationResult(int status) {
            /// <summary>
            /// Uses Windows Update result code as input to return 
            /// text string status result equivalent
            /// </summary>
           switch (status) {
                default:
                    return "No Status";
                case (int)enums.OperationResultCode.NotStarted:
                    return enums.OperationResultCode.NotStarted.ToString();
                case (int)enums.OperationResultCode.InProgress:
                    return enums.OperationResultCode.InProgress.ToString();
                case (int)enums.OperationResultCode.Succeeded:
                    return enums.OperationResultCode.Succeeded.ToString();
                case (int)enums.OperationResultCode.SucceededWithErrors:
                    return enums.OperationResultCode.SucceededWithErrors.ToString();
                case (int)enums.OperationResultCode.Failed:
                    return enums.OperationResultCode.Failed.ToString();
                case (int)enums.OperationResultCode.Aborted:
                    return enums.OperationResultCode.Aborted.ToString();
            }
            
        }

        // Output values returned by class
        public string Name { get { return _name; } }
        public string Operation { get { return GetUpdateOperation(_operation); } }
        public string Status { get { return GetUpdateOperationResult(_status); } }
        public string Description { get { return _description; } }
        public string KB { get { return _kb; } }
        public string Date { get { return _date; } }
    }

}
#endregion Containers

#region Enums
namespace EIO.WUWrapper.Enums {
    public class enums {
        // Windows Updates operation codes
        public enum OperationCode {
            // This enumeration doesn't start at 0 so must define start value
            Installed = 1,
            Uninstalled
        }

        // Windows Updates operation result codes
        public enum OperationResultCode {
            NotStarted,
            InProgress,
            Succeeded,
            SucceededWithErrors,
            Failed,
            Aborted
        }
    }
}
#endregion Enums