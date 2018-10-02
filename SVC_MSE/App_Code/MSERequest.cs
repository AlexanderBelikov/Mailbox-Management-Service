using System;
using System.Runtime.Serialization;
using System.Web.Services;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.Web.Script.Services;
using System.Web.Script.Serialization;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text.RegularExpressions;

/// <summary>
/// Summary description for MSERequest
/// </summary>
[WebService(Namespace = "http://zzz.su.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class MSERequest : System.Web.Services.WebService
{
    private static String DomainControllerName = "";
    private EventLog eventLog;
    public MSERequest()
    {
        eventLog = new EventLog();
        if (!EventLog.SourceExists("svc_idm_mse"))
        {
            EventLog.CreateEventSource("svc_idm_mse", "svc_idm_mse");
        }
        eventLog.Source = "svc_idm_mse";
        eventLog.Log = "svc_idm_mse";        
        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }

    //--------------------------------------------------

    //  ENABLE MAILBOX ROLES
    //--------------------------------------------------
    //  EnableMailboxRolesMapiOwa
    [WebMethod]
    [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
    public string EnableMailboxRolesMapiOwa(string RequestId, string SamAccountName, string RequestedAddress)
    {
        string[] Roles = new string[] { "MAPIEnabled", "OWAEnabled", "OWAforDevicesEnabled" };
        Response EnableMailboxRolesResult = EnableMailboxRoles(RequestId, SamAccountName, RequestedAddress, Roles);
        return new JavaScriptSerializer().Serialize(EnableMailboxRolesResult);
    }
    //  EnableMailboxRolesMapiOwa
    //--------------------------------------------------
    //  EnableMailboxRolesActiveSync
    [WebMethod]
    [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
    public string EnableMailboxRolesActiveSync(string RequestId, string SamAccountName, string RequestedAddress)
    {
        string[] Roles = new string[] { "ActiveSyncEnabled" };
        Response EnableMailboxRolesResult = EnableMailboxRoles(RequestId, SamAccountName, RequestedAddress, Roles);
        return new JavaScriptSerializer().Serialize(EnableMailboxRolesResult);
    }
    //  EnableMailboxRolesActiveSync
    //--------------------------------------------------
    //  EnableMailboxRoles
    public Response EnableMailboxRoles(string RequestId, string SamAccountName, string RequestedAddress, string[] Roles)
    {
        List<string> EventLogMessage = new List<string>();
        EventLogMessage.Add("RequestId: " + RequestId);
        EventLogMessage.Add("Requester: " + System.Web.HttpContext.Current.User.Identity.Name);
        EventLogMessage.Add("EnableMailboxRoles: " + String.Join(", ", Roles));
        EventLogMessage.Add("SamAccountName: " + SamAccountName);
        EventLogMessage.Add("RequestedAddress: " + RequestedAddress);


        // Get Domain Controller Name
        Response TestDomainControllerResult = TestDomainController();
        if (TestDomainControllerResult.Error > 0)
        {
            EventLogMessage.Add("Error: " + TestDomainControllerResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return TestDomainControllerResult;
        }
        EventLogMessage.Add("DomainControllerName: " + DomainControllerName);

        // Get User
        Response GetUserResult = GetUser(SamAccountName);
        if (GetUserResult.Error > 0)
        {
            EventLogMessage.Add("Error: " + GetUserResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return GetUserResult;
        }

        PSObject User = GetUserResult.Object;
        string RecipientType = User.Properties["RecipientType"].Value.ToString();
        string RecipientTypeDetails = User.Properties["RecipientTypeDetails"].Value.ToString();
        string DistinguishedName = User.Properties["DistinguishedName"].Value.ToString();
        string EmailAddress = null;
        EventLogMessage.Add("RecipientType: " + RecipientType);

        if (string.Equals(RecipientType, "User", StringComparison.InvariantCultureIgnoreCase))
        {
            EventLogMessage.Add("RecipientTypeDetails: " + RecipientTypeDetails);

            // Get new email address
            EventLogMessage.Add("NewEmailAddress");
            if (RequestedAddress.IndexOf('@') != -1)
            {
                Response EmailExistResult = EmailExist(RequestedAddress);
                if (!(EmailExistResult.Error == 0 && EmailExistResult.Message == null))
                {
                    EventLogMessage.Add("Error: " + EmailExistResult.Message);
                    eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
                    return EmailExistResult;
                }
                EmailAddress = RequestedAddress;
                EventLogMessage.Add("EmailAddress: " + EmailAddress);
            }
            else
            {
                Response NewEmailAddressResult = NewEmailAddress(User, RequestedAddress);
                if (NewEmailAddressResult.Error != 0)
                {
                    EventLogMessage.Add("Error: " + NewEmailAddressResult.Message);
                    eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
                    return NewEmailAddressResult;
                }
                EmailAddress = NewEmailAddressResult.Message;
                EventLogMessage.Add("EmailAddress: " + EmailAddress);
            }

            // Enable ADAccount
            if (string.Equals(RecipientTypeDetails, "DisabledUser", StringComparison.InvariantCultureIgnoreCase))
            {
                EventLogMessage.Add("EnableADAccount");
                Response EnableADAccountResult = EnableADAccount(DistinguishedName);
                if (EnableADAccountResult.Error > 0)
                {
                    EventLogMessage.Add("Error: " + EnableADAccountResult.Message);
                    eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
                    return EnableADAccountResult;
                }
            }

            // Enable Mailbox
            EventLogMessage.Add("EnableMailbox");
            Response EnableMailboxResult = EnableMailbox(DistinguishedName, EmailAddress);
            if (EnableMailboxResult.Error != 0)
            {
                EventLogMessage.Add("Error: " + EnableMailboxResult.Message);
                eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
                return EnableMailboxResult;
            }

            // Disable ADAccount
            if (string.Equals(RecipientTypeDetails, "DisabledUser", StringComparison.InvariantCultureIgnoreCase))
            {
                EventLogMessage.Add("DisableADAccount");
                Response DisableADAccountResult = DisableADAccount(DistinguishedName);
                if (DisableADAccountResult.Error > 0)
                {
                    EventLogMessage.Add("Error: " + DisableADAccountResult.Message);
                    eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
                    return DisableADAccountResult;
                }
            }
        }
        else if (string.Equals(RecipientType, "UserMailbox", StringComparison.InvariantCultureIgnoreCase))
        {
            EmailAddress = User.Properties["WindowsEmailAddress"].Value.ToString().ToLower();
        }
        else
        {
            EventLogMessage.Add("EnableMailboxRoles: Error(s) occurred: Unknown RecipientType " + RecipientType);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return new Response() { Error = 1, Message = "EnableMailboxRoles: Error(s) occurred: Unknown RecipientType " + RecipientType };
        }

        // Set CASMailbox
        EventLogMessage.Add("SetCASMailbox");
        Response SetCASMailboxResult = SetCASMailbox(DistinguishedName, Roles, true);
        if (SetCASMailboxResult.Error != 0)
        {
            EventLogMessage.Add("Error: " + SetCASMailboxResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return SetCASMailboxResult;
        }

        EventLogMessage.Add("Success");
        eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Information);
        return new Response() { Error = 0, Message = EmailAddress };
    }
    //  EnableMailboxRoles
    //--------------------------------------------------
    //  EnableMailbox
    private Response EnableMailbox(string DistinguishedName, string EmailAddress)
    {
        using (Runspace EnableMailboxRunspace = RunspaceFactory.CreateRunspace(MSEConnectionInfo()))
        {
            EnableMailboxRunspace.Open();

            using (Pipeline EnableMailboxPipeline = EnableMailboxRunspace.CreatePipeline())
            {
                // Enable mailbox
                Command cmdEnableMailbox = new Command("Enable-Mailbox");
                cmdEnableMailbox.Parameters.Add("Identity", @DistinguishedName);
                cmdEnableMailbox.Parameters.Add("Alias", @EmailAddress.Split('@')[0]);
                cmdEnableMailbox.Parameters.Add("PrimarySmtpAddress", @EmailAddress);
                cmdEnableMailbox.Parameters.Add("DomainController", @DomainControllerName);
                EnableMailboxPipeline.Commands.Add(cmdEnableMailbox);
                try
                {
                    EnableMailboxPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    return new Response() { Error = 1, Message = "EnableMailbox: Enable-Mailbox Exception: " + ex.Message };
                }

                if (EnableMailboxPipeline.Error.Count > 0)
                {
                    string sError = "EnableMailbox: Enable-Mailbox: Error(s) occurred: ";
                    if (EnableMailboxPipeline.Error.Count == 1)
                    {
                        var Error = EnableMailboxPipeline.Error.Read() as ErrorRecord;
                        sError += Error.Exception.Message;
                    }
                    else
                    {
                        var Errors = EnableMailboxPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }
                return new Response() { Error = 0, Message = EmailAddress };
            }
        }
    }
    //  EnableMailbox
    //--------------------------------------------------

    //  DISABLE MAILBOX ROLES
    //--------------------------------------------------
    //  DisableMailboxRole
    [WebMethod]
    [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
    public string DisableMailboxRole(string RequestId, string SamAccountName)
    {
        List<string> EventLogMessage = new List<string>();
        EventLogMessage.Add("RequestId: " + RequestId);
        EventLogMessage.Add("Requester: " + System.Web.HttpContext.Current.User.Identity.Name);
        EventLogMessage.Add("DisableMailboxRole");
        EventLogMessage.Add("SamAccountName: " + SamAccountName);

        // Get Domain Controller Name
        Response TestDomainControllerResult = TestDomainController();
        if (TestDomainControllerResult.Error > 0)
        {
            EventLogMessage.Add("Error: " + TestDomainControllerResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return new JavaScriptSerializer().Serialize(TestDomainControllerResult);
        }
        EventLogMessage.Add("DomainControllerName: " + DomainControllerName);

        // Get User
        Response GetUserResult = GetUser(SamAccountName);
        if (GetUserResult.Error > 0)
        {
            EventLogMessage.Add("Error: " + GetUserResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return new JavaScriptSerializer().Serialize(GetUserResult);
        }

        PSObject User = GetUserResult.Object;
        string RecipientType = User.Properties["RecipientType"].Value.ToString();
        string DistinguishedName = User.Properties["DistinguishedName"].Value.ToString();
        EventLogMessage.Add("RecipientType: " + RecipientType);
        if (!string.Equals(RecipientType, "UserMailbox", StringComparison.InvariantCultureIgnoreCase))
        {
            EventLogMessage.Add("Warning: RecipientType " + RecipientType + " is not UserMailbox");
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Warning);
            return new JavaScriptSerializer().Serialize(new Response() { Error = 0, Message = "Warning: RecipientType is not UserMailbox" });
        }

        // Disable Mailbox
        EventLogMessage.Add("DisableMailbox");
        Response DisableMailboxResult = DisableMailbox(DistinguishedName);
        if (DisableMailboxResult.Error != 0)
        {
            EventLogMessage.Add("Error: " + DisableMailboxResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return new JavaScriptSerializer().Serialize(DisableMailboxResult);
        }

        EventLogMessage.Add("Success");
        eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Information);
        return new JavaScriptSerializer().Serialize(new Response() { Error = 0, Message = "Success" });
    }
    //  DisableMailboxRole    //  DisableMailboxRolesMapiOwa
    [WebMethod]
    [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
    public string DisableMailboxRolesMapiOwa(string RequestId, string SamAccountName)
    {
        string[] Roles = new string[] { "MAPIEnabled", "OWAEnabled", "OWAforDevicesEnabled" };
        Response DisableMailboxRolesResult = DisableMailboxRoles(RequestId, SamAccountName, Roles);
        return new JavaScriptSerializer().Serialize(DisableMailboxRolesResult);
    }
    //  DisableMailboxRolesMapiOwa
    //--------------------------------------------------
    //  DisableMailboxRolesActiveSync
    [WebMethod]
    [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
    public string DisableMailboxRolesActiveSync(string RequestId, string SamAccountName)
    {
        string[] Roles = new string[] { "ActiveSyncEnabled" };
        Response DisableMailboxRolesResult = DisableMailboxRoles(RequestId, SamAccountName, Roles);
        return new JavaScriptSerializer().Serialize(DisableMailboxRolesResult);
    }
    //  DisableMailboxRolesActiveSync
    //--------------------------------------------------
    //  DisableMailboxRoles
    public Response DisableMailboxRoles(string RequestId, string SamAccountName, string[] Roles)
    {
        List<string> EventLogMessage = new List<string>();
        EventLogMessage.Add("RequestId: " + RequestId);
        EventLogMessage.Add("Requester: " + System.Web.HttpContext.Current.User.Identity.Name);
        EventLogMessage.Add("DisableMailboxRoles: " + String.Join(", ", Roles));
        EventLogMessage.Add("SamAccountName: " + SamAccountName);

        // Get Domain Controller Name
        Response TestDomainControllerResult = TestDomainController();
        if (TestDomainControllerResult.Error > 0)
        {
            EventLogMessage.Add("Error: " + TestDomainControllerResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return TestDomainControllerResult;
        }
        EventLogMessage.Add("DomainControllerName: " + DomainControllerName);

        // Get User
        Response GetUserResult = GetUser(SamAccountName);
        if (GetUserResult.Error > 0)
        {
            EventLogMessage.Add("Error: " + GetUserResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return GetUserResult;
        }

        PSObject User = GetUserResult.Object;
        string RecipientType = User.Properties["RecipientType"].Value.ToString();
        string DistinguishedName = User.Properties["DistinguishedName"].Value.ToString();
        EventLogMessage.Add("RecipientType: " + RecipientType);
        if (!string.Equals(RecipientType, "UserMailbox", StringComparison.InvariantCultureIgnoreCase))
        {
            EventLogMessage.Add("Warning: RecipientType " + RecipientType + " is not UserMailbox");
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Warning);
            return new Response() { Error = 0, Message = "Warning: RecipientType " + RecipientType + " is not UserMailbox" };
        }

        // Set CASMailbox
        EventLogMessage.Add("SetCASMailbox");
        Response SetCASMailboxResult = SetCASMailbox(DistinguishedName, Roles, false);
        if (SetCASMailboxResult.Error != 0)
        {
            EventLogMessage.Add("Error: " + SetCASMailboxResult.Message);
            eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Error);
            return SetCASMailboxResult;
        }

        EventLogMessage.Add("Success");
        eventLog.WriteEntry(String.Join(System.Environment.NewLine, EventLogMessage), EventLogEntryType.Information);
        return new Response() { Error = 0, Message = "Success" };
    }
    //  DisableMailboxRoles
    //--------------------------------------------------
    //  DisableMailbox
    private Response DisableMailbox(string DistinguishedName)
    {

        using (Runspace DisableMailboxRunspace = RunspaceFactory.CreateRunspace(MSEConnectionInfo()))
        {
            DisableMailboxRunspace.Open();

            using (Pipeline DisableMailboxPipeline = DisableMailboxRunspace.CreatePipeline())
            {
                // Enable mailbox
                Command cmdDisableMailbox = new Command("Disable-Mailbox");
                cmdDisableMailbox.Parameters.Add("Identity", @DistinguishedName);
                cmdDisableMailbox.Parameters.Add("Confirm", false);
                cmdDisableMailbox.Parameters.Add("DomainController", DomainControllerName);
                DisableMailboxPipeline.Commands.Add(cmdDisableMailbox);
                try
                {
                    DisableMailboxPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    return new Response() { Error = 1, Message = "DisableMailbox: Disable-Mailbox Exception: " + ex.Message };
                }

                if (DisableMailboxPipeline.Error.Count > 0)
                {
                    string sError = "DisableMailbox: Disable-Mailbox: Error(s) occurred: ";
                    if (DisableMailboxPipeline.Error.Count == 1)
                    {
                        var Error = DisableMailboxPipeline.Error.Read() as ErrorRecord;
                        sError += Error.Exception.Message;
                    }
                    else
                    {
                        var Errors = DisableMailboxPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }
                return new Response() { Error = 0 };
            }
        }
    }
    //  DisableMailbox
    //--------------------------------------------------
    //  TestDomainController
    private Response TestDomainController()
    {
        Response DomainControllerAvailabilityResult;
        // Test DCName if not empty and return if DC avaiable   
        if (DomainControllerName != "")
        {
            DomainControllerAvailabilityResult = DomainControllerAvailability(DomainControllerName);
            if (DomainControllerAvailabilityResult.Error == 0)
            {
                return DomainControllerAvailabilityResult;
            }
        }

        // Get new DC, return if error
        eventLog.WriteEntry("New DomainControllerName", EventLogEntryType.Information);
        Response GetDomainControllerResult = GetDomainController();
        if (GetDomainControllerResult.Error > 0)
        {
            return GetDomainControllerResult;
        }
        string newDomainControllerName = GetDomainControllerResult.Object.Properties["DefaultGlobalCatalog"].Value.ToString();
        eventLog.WriteEntry(newDomainControllerName+" instead "+DomainControllerName, EventLogEntryType.Information);
        DomainControllerName = newDomainControllerName;

        // Test DCName if not empty and return if DC avaiable   
        DomainControllerAvailabilityResult = DomainControllerAvailability(DomainControllerName);
        return DomainControllerAvailabilityResult;
    }
    //  TestDomainController
    //--------------------------------------------------
    //  DomainControllerAvailability
    private Response DomainControllerAvailability(string DCName)
    {
        using (Runspace DomainControllerAvailabilityRunspace = RunspaceFactory.CreateRunspace(MSEConnectionInfo()))
        {
            DomainControllerAvailabilityRunspace.Open();
            using (Pipeline DomainControllerAvailabilityPipeline = DomainControllerAvailabilityRunspace.CreatePipeline())
            {
                Command cmdDomainControllerAvailability = new Command("Get-User");
                cmdDomainControllerAvailability.Parameters.Add("Filter", "SamAccountName -eq \"" + System.Web.HttpContext.Current.User.Identity.Name + "\"");
                cmdDomainControllerAvailability.Parameters.Add("DomainController", DCName);
                DomainControllerAvailabilityPipeline.Commands.Add(cmdDomainControllerAvailability);
                Collection<PSObject> user = null;
                try
                {
                    user = DomainControllerAvailabilityPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    eventLog.WriteEntry("DomainControllerAvailability "+DomainControllerName+": Error: "+ex.Message, EventLogEntryType.Error);
                    return new Response() { Error = 1, Message = "DomainControllerAvailability: " + ex.Message };
                }

                if (DomainControllerAvailabilityPipeline.Error.Count > 0)
                {
                    string sError = "DomainControllerAvailability: Error(s) occurred: ";
                    if (DomainControllerAvailabilityPipeline.Error.Count == 1)
                    {
                        var Error = DomainControllerAvailabilityPipeline.Error.Read() as ErrorRecord;
                        sError += Error.Exception.Message;
                    }
                    else
                    {
                        var Errors = DomainControllerAvailabilityPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }
                return new Response() { Error = 0 };
            }
        }
    }
    //  DomainControllerAvailability
    //--------------------------------------------------
    //  GetDomainController
    private Response GetDomainController()
    {
        using (Runspace GetDomainControllerRunspace = RunspaceFactory.CreateRunspace(MSEConnectionInfo()))
        {
            GetDomainControllerRunspace.Open();

            using (Pipeline GetDomainControllerPipeline = GetDomainControllerRunspace.CreatePipeline())
            {
                Command cmdGetDomainController = new Command("Get-AdServerSettings");
                GetDomainControllerPipeline.Commands.Add(cmdGetDomainController);
                Collection<PSObject> AdServerSettings = null;
                try
                {
                    AdServerSettings = GetDomainControllerPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    return new Response() { Error = 1, Message = ex.Message };
                }

                if (GetDomainControllerPipeline.Error.Count > 0)
                {
                    string sError = "GetDomainController: Error(s) occurred: ";
                    if (GetDomainControllerPipeline.Error.Count == 1)
                    {
                        var Error = GetDomainControllerPipeline.Error.Read() as ErrorRecord;
                        sError += Error.Exception.Message;
                    }
                    else
                    {
                        var Errors = GetDomainControllerPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }
                if (AdServerSettings.Count != 1)
                {
                    return new Response() { Error = 1, Message = "GetDomainController: No unique object found" };
                }
                else
                {
                    return new Response() { Error = 0, Object = AdServerSettings[0] };
                }
            }
        }
    }
    //  GetDomainController
    //--------------------------------------------------
    //  GetUser
    private Response GetUser(string SamAccountName)
    {
        using (Runspace GetUserRunspace = RunspaceFactory.CreateRunspace(MSEConnectionInfo()))
        {
            GetUserRunspace.Open();

            using (Pipeline GetUserPipeline = GetUserRunspace.CreatePipeline())
            {
                Command cmdGetUser = new Command("Get-User");
                cmdGetUser.Parameters.Add("Filter", "SamAccountName -eq \"" + @SamAccountName + "\"");
                cmdGetUser.Parameters.Add("DomainController", @DomainControllerName);
                GetUserPipeline.Commands.Add(cmdGetUser);
                Collection<PSObject> user = null;
                try
                {
                    user = GetUserPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    return new Response() { Error = 1, Message = ex.Message };
                }

                if (GetUserPipeline.Error.Count > 0)
                {
                    string sError = "GetUser: Error(s) occurred: ";
                    if (GetUserPipeline.Error.Count == 1)
                    {
                        var Error = GetUserPipeline.Error.Read() as ErrorRecord;
                        sError += Error.Exception.Message;
                    }
                    else
                    {
                        var Errors = GetUserPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }
                if (user.Count != 1)
                {
                    return new Response() { Error = 1, Message = "GetUser: No unique user found" };
                }
                else
                {
                    return new Response() { Error = 0, Object = user[0] };
                }
            }
        }
    }
    //  GetUser
    //--------------------------------------------------
    //  SetCASMailbox
    private Response SetCASMailbox(string DistinguishedName, string[] Roles, bool Enable)
    {
        if (Roles.Length == 0)
        {
            return new Response() { Error = 1, Message = "SetCASMailbox: Empty roles" };
        }
        using (Runspace SetCASMailboxRunspace = RunspaceFactory.CreateRunspace(MSEConnectionInfo()))
        {
            SetCASMailboxRunspace.Open();

            using (Pipeline SetCASMailboxPipeline = SetCASMailboxRunspace.CreatePipeline())
            {
                Command cmdSetCASMailbox = new Command("Set-CASMailbox");
                cmdSetCASMailbox.Parameters.Add("Identity", @DistinguishedName);
                cmdSetCASMailbox.Parameters.Add("DomainController", @DomainControllerName);
                foreach (string Role in Roles)
                {
                    cmdSetCASMailbox.Parameters.Add(Role, Enable);
                }
                SetCASMailboxPipeline.Commands.Add(cmdSetCASMailbox);
                try
                {
                    SetCASMailboxPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    return new Response() { Error = 1, Message = ex.Message };
                }

                if (SetCASMailboxPipeline.Error.Count > 0)
                {
                    string sError = "SetCASMailbox: Error(s) occurred: ";
                    if (SetCASMailboxPipeline.Error.Count == 1)
                    {
                        var Error = SetCASMailboxPipeline.Error.Read() as ErrorRecord;
                        sError += Error.Exception.Message;
                    }
                    else
                    {
                        var Errors = SetCASMailboxPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }
                return new Response() { Error = 0 };
            }
        }
    }
    //  SetCASMailbox
    //--------------------------------------------------
    //  EnableADAccount
    private Response EnableADAccount(string DistinguishedName)
    {
        InitialSessionState iss = InitialSessionState.CreateDefault();
        iss.ImportPSModule(new string[] { "activedirectory" });

        using (Runspace EnableADAccountRunspace = RunspaceFactory.CreateRunspace(iss))
        {
            EnableADAccountRunspace.Open();
            using (Pipeline EnableADAccountPipeline = EnableADAccountRunspace.CreatePipeline())
            {
                Command cmdEnableADAccount = new Command("Enable-ADAccount");
                cmdEnableADAccount.Parameters.Add("Identity", @DistinguishedName);
                cmdEnableADAccount.Parameters.Add("Server", @DomainControllerName);
                EnableADAccountPipeline.Commands.Add(cmdEnableADAccount);
                try
                {
                    EnableADAccountPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    return new Response() { Error = 1, Message = "EnableADAccount: Exception: " + ex.Message };
                }
                if (EnableADAccountPipeline.Error.Count > 0)
                {
                    string sError = "EnableADAccount: Error(s) occurred: ";
                    if (EnableADAccountPipeline.Error.Count == 1)
                    {
                        var Error = EnableADAccountPipeline.Error.Read() as ErrorRecord;
                        sError += Error.Exception.Message;
                    }
                    else
                    {
                        var Errors = EnableADAccountPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }
                return new Response() { Error = 0 };
            }
        }
    }
    //  EnableADAccount
    //--------------------------------------------------
    //  DisableADAccount
    private Response DisableADAccount(string DistinguishedName)
    {
        InitialSessionState iss = InitialSessionState.CreateDefault();
        iss.ImportPSModule(new string[] { "activedirectory" });

        using (Runspace DisableADAccountRunspace = RunspaceFactory.CreateRunspace(iss))
        {
            DisableADAccountRunspace.Open();
            using (Pipeline DisableADAccountPipeline = DisableADAccountRunspace.CreatePipeline())
            {
                Command cmdDisableADAccount = new Command("Disable-ADAccount");
                cmdDisableADAccount.Parameters.Add("Identity", @DistinguishedName);
                cmdDisableADAccount.Parameters.Add("Server", @DomainControllerName);
                DisableADAccountPipeline.Commands.Add(cmdDisableADAccount);
                try
                {
                    DisableADAccountPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    return new Response() { Error = 1, Message = "DisableADAccount: Exception: " + ex.Message };
                }
                if (DisableADAccountPipeline.Error.Count > 0)
                {
                    string sError = "DisableADAccount: Error(s) occurred: ";
                    if (DisableADAccountPipeline.Error.Count == 1)
                    {
                        var Error = DisableADAccountPipeline.Error.Read() as ErrorRecord;
                        sError += Error.Exception.Message;
                    }
                    else
                    {
                        var Errors = DisableADAccountPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }
                return new Response() { Error = 0 };
            }
        }
    }
    //  DisableADAccount
    //--------------------------------------------------
    //  NewEmailAddress
    private Response NewEmailAddress(PSObject User, string EmailAddress)
    {
        if (EmailAddress == "")
        {
            return new Response() { Error = 1, Message = "NewEmailAddress: RequestedAddress is empty" };
        }
        using (Runspace NewEmailAddressRunspace = RunspaceFactory.CreateRunspace(MSEConnectionInfo()))
        {
            NewEmailAddressRunspace.Open();
            using (Pipeline NewEmailAddressPipeline = NewEmailAddressRunspace.CreatePipeline())
            {
                string FirstName = User.Properties["FirstName"].Value.ToString();
                string Initials = User.Properties["Initials"].Value.ToString();
                string LastName = User.Properties["LastName"].Value.ToString();
                if (FirstName == "" || Initials == "" || LastName == "")
                {
                    return new Response() { Error = 1, Message = "NewEmailAddress: FirstName or Initials or LastName is empty" };
                }
                string AliasTemplate = LastName + ".#" + Initials[0];
                AliasTemplate = AliasTemplate.ToLower();
                Regex regex = new Regex(@"^[a-z]{2,64}\.[a-z]{2,64}$");
                for (int i = 1; i <= FirstName.Length; i++)
                {
                    string EmailAlias = AliasTemplate.Replace("#", FirstName.Substring(0, i).ToLower());
                    Match match = regex.Match(EmailAlias);
                    if (match.Success)
                    {

                        string NewEmail = EmailAlias + "@" + EmailAddress;
                        Response EmailExistResult = EmailExist(NewEmail);
                        if (EmailExistResult.Error != 0)
                        {
                            return EmailExistResult;
                        }
                        if (EmailExistResult.Message == null)
                        {
                            return new Response() { Error = 0, Message = NewEmail };
                        }
                    }
                    else
                    {
                        return new Response() { Error = 1, Message = "NewEmailAddress: EmailAlias not match pattern | " + EmailAlias };
                    }
                }
                return new Response() { Error = 1, Message = "NewEmailAddress: no valid email found" };
            }
        }
    }
    //  NewEmailAddress
    //--------------------------------------------------
    //  EmailExist
    //  Return Error 0 and Message null if email does not exist or Message = EmailAddress if exist
    private Response EmailExist(string EmailAddress)
    {
        if (EmailAddress == "")
        {
            return new Response() { Error = 1, Message = "EmailExist: RequestedAddress is empty" };
        }
        using (Runspace EmailExistRunspace = RunspaceFactory.CreateRunspace(MSEConnectionInfo()))
        {
            EmailExistRunspace.Open();

            using (Pipeline EmailExistPipeline = EmailExistRunspace.CreatePipeline())
            {
                Command cmdEmailExist = new Command("Get-Recipient");
                cmdEmailExist.Parameters.Add("Identity", @EmailAddress);
                cmdEmailExist.Parameters.Add("DomainController", @DomainControllerName);
                EmailExistPipeline.Commands.Add(cmdEmailExist);
                Collection<PSObject> user = null;
                try
                {
                    user = EmailExistPipeline.Invoke();
                }
                catch (System.Exception ex)
                {
                    return new Response() { Error = 1, Message = ex.Message };
                }

                if (EmailExistPipeline.Error.Count > 0)
                {
                    string sError = "EmailExist: Error(s) occurred: ";
                    if (EmailExistPipeline.Error.Count == 1)
                    {
                        var Error = EmailExistPipeline.Error.Read() as ErrorRecord;
                        sError += Error.CategoryInfo.Reason;
                        if (string.Equals(Error.CategoryInfo.Reason, "ManagementObjectNotFoundException", StringComparison.OrdinalIgnoreCase))
                        {
                            return new Response() { Error = 0, Message = null };
                        }
                    }
                    else
                    {
                        var Errors = EmailExistPipeline.Error.Read() as Collection<ErrorRecord>;
                        foreach (ErrorRecord Error in Errors)
                        {
                            sError += Error.Exception.Message + " ";
                            if (string.Equals(Error.CategoryInfo.Reason, "ManagementObjectNotFoundException", StringComparison.OrdinalIgnoreCase))
                            {
                                return new Response() { Error = 0, Message = null };
                            }
                        }
                    }
                    return new Response() { Error = 1, Message = sError };
                }

                if (user.Count > 0)
                {
                    return new Response() { Error = 0, Message = "EmailExist: "+ EmailAddress + " already exist" };
                }
                return new Response() { Error = 1, Message = "EmailExist: Unknown error" };
            }
        }
    }
    //  EmailExist
    //--------------------------------------------------
    //  MSEConnectionInfo
    private WSManConnectionInfo MSEConnectionInfo()
    {
        string ConnectionUri = Resources.MSERequest.MSEConnectionUri;
        WSManConnectionInfo ConnectionInfo = new WSManConnectionInfo((new Uri(ConnectionUri)), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", (PSCredential)null);
        ConnectionInfo.AuthenticationMechanism = AuthenticationMechanism.Default;
        return ConnectionInfo;
    }
    //  MSEConnectionInfo
}

[DataContract]
public class Response
{
    [DataMember]
    public int Error { get; set; }
    [DataMember]
    public string Message { get; set; }
    [DataMember]
    public PSObject Object { get; set; }
}