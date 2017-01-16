// =====================================================================
//  This file is part of the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Crm.UnifiedServiceDesk.CommonUtility;
using Microsoft.Crm.UnifiedServiceDesk.Dynamics;
using Microsoft.Crm.UnifiedServiceDesk.Dynamics.Utilities;
using Microsoft.Uii.Desktop.SessionManager;
using System.IO;
using Microsoft.Xrm.Sdk;
using Microsoft.Uii.Csr;
using System.IO.Compression;
using System.Linq;

namespace HiddenPasteHandler
{
    /// <summary>
    /// Interaction logic for USDControl.xaml
    /// This is a base control for building Unified Service Desk Aware add-ins
    /// See USD API documentation for full API Information available via this control.
    /// </summary>
    public partial class USDControl : DynamicsBaseHostedControl
    {
        #region Vars
        /// <summary>
        /// Log writer for USD 
        /// </summary>
        private TraceLogger LogWriter = null;

        #endregion

        /// <summary>
        /// UII Constructor 
        /// </summary>
        /// <param name="appID">ID of the application</param>
        /// <param name="appName">Name of the application</param>
        /// <param name="initString">Initializing XML for the application</param>
        public USDControl(Guid appID, string appName, string initString)
            : base(appID, appName, initString)
        {
            InitializeComponent();

            // This will create a log writer with the default provider for Unified Service desk
            LogWriter = new TraceLogger();

            #region Enhanced LogProvider Info
            // This will create a log writer with the same name as your hosted control. 
            // LogWriter = new TraceLogger(traceSourceName:"MyTraceSource");

            // If you utilize this feature,  you would need to add a section to the system.diagnostics settings area of the UnifiedServiceDesk.exe.config
            //<source name="MyTraceSource" switchName="MyTraceSwitchName" switchType="System.Diagnostics.SourceSwitch">
            //    <listeners>
            //        <add name="console" type="System.Diagnostics.DefaultTraceListener"/>
            //        <add name="fileListener"/>
            //        <add name="USDDebugListener" />
            //        <remove name="Default"/>
            //    </listeners>
            //</source>

            // and then in the switches area : 
            //<add name="MyTraceSwitchName" value="Verbose"/>

            #endregion

        }

        /// <summary>
        /// Raised when the Desktop Ready event is fired. 
        /// </summary>
        protected override void DesktopReady()
        {
            // this will populate any toolbars assigned to this control in config. 
            PopulateToolbars(ProgrammableToolbarTray);
            base.DesktopReady();
        }

        /// <summary>
        /// Raised when an action is sent to this control
        /// </summary>
        /// <param name="args">args for the action</param>
        protected override void DoAction(Microsoft.Uii.Csr.RequestActionEventArgs args)
        {
            // Log process.
            LogWriter.Log(string.Format(CultureInfo.CurrentCulture, "{0} -- DoAction called for action: {1}", this.ApplicationName, args.Action), System.Diagnostics.TraceEventType.Information);

            #region Example process action
            //// Process Actions. 
            //if (args.Action.Equals("your action name", StringComparison.OrdinalIgnoreCase))
            //{
            //    // Do some work

            //    // Access CRM and fetch a Record
            //    Microsoft.Xrm.Sdk.Messages.RetrieveRequest req = new Microsoft.Xrm.Sdk.Messages.RetrieveRequest(); 
            //    req.Target = new Microsoft.Xrm.Sdk.EntityReference( "account" , Guid.Parse("0EF05F4F-0D39-4219-A3F5-07A0A5E46FD5")); 
            //    req.ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("accountid" , "name" );
            //    Microsoft.Xrm.Sdk.Messages.RetrieveResponse response = (Microsoft.Xrm.Sdk.Messages.RetrieveResponse)this._client.CrmInterface.ExecuteCrmOrganizationRequest(req, "Requesting Account"); 


            //    // Example of pulling some data out of the passed in data array
            //    List<KeyValuePair<string, string>> actionDataList = Utility.SplitLines(args.Data, CurrentContext, localSession);
            //    string valueIwant = Utility.GetAndRemoveParameter(actionDataList, "mykey"); // asume there is a myKey=<value> in the data. 



            //    // Example of pushing data to USD
            //    string global = Utility.GetAndRemoveParameter(actionDataList, "global"); // Assume there is a global=true/false in the data
            //    bool saveInGlobalSession = false;
            //    if (!String.IsNullOrEmpty(global))
            //        saveInGlobalSession = bool.Parse(global);

            //    Dictionary<string, CRMApplicationData> myDataToSet = new Dictionary<string, CRMApplicationData>();
            //    // add a string: 
            //    myDataToSet.Add("myNewKey", new CRMApplicationData() { name = "myNewKey", type = "string", value = "TEST" });

            //    // add a entity lookup:
            //    myDataToSet.Add("myNewKey", new CRMApplicationData() { name = "myAccount", type = "lookup", value = "account,0EF05F4F-0D39-4219-A3F5-07A0A5E46FD5,MyAccount" }); 

            //    if (saveInGlobalSession) 
            //    {
            //        // add context item to the global session
            //        ((DynamicsCustomerRecord)((AgentDesktopSession)localSessionManager.GlobalSession).Customer.DesktopCustomer).MergeReplacementParameter(this.ApplicationName, myDataToSet, true);
            //    }
            //    else
            //    {
            //        // Add context item to the current session. 
            //        ((DynamicsCustomerRecord)((AgentDesktopSession)localSessionManager.ActiveSession).Customer.DesktopCustomer).MergeReplacementParameter(this.ApplicationName, myDataToSet, true);
            //    }
            //}
            #endregion

            if (args.Action == "Paste")
            {
                
                PasteContent(args);
            }

            base.DoAction(args);
        }

        /// <summary>
        /// Raised when a context change occurs in USD
        /// </summary>
        /// <param name="context"></param>
        public override void NotifyContextChange(Microsoft.Uii.Csr.Context context)
        {
            base.NotifyContextChange(context);
        }



        #region User Code Area
        //Xrm.Page.ui.setFormNotification("check this", "INFO", "PASTEINFOMESSAGE");
        private void PasteContent(Microsoft.Uii.Csr.RequestActionEventArgs args)
        {
            try
            {
                if (!ValidClipboard())
                {
                    FireRequestAction(new RequestActionEventArgs("CRM Global Manager", "DisplayMessage", $"Unable to capture copied content !!"));
                    return;
                }

                var activeApplication = localSession.FocusedApplication.ApplicationName;  // TODO: What it is if it is custom entity?
                var activeAppId = GetIdOfApplication(activeApplication);
                var activeAppLogicalName = GetLogicalName(activeApplication);
                ClearNotification(activeApplication);

                byte[] bytes = null;
                string fileName = string.Empty;
                GetContentFromClipboard(ref bytes, ref fileName);

                CreateAnnotation($"Pasted on {DateTime.Now}", fileName, bytes, activeAppLogicalName, activeAppId);
                SetNotification(activeApplication);
                return;
            }
            catch (Exception ex)
            {
                FireRequestAction(new RequestActionEventArgs("CRM Global Manager", "DisplayMessage", $"Failed to attach notes, Error :- {ex.Message}"));
                throw ex;
            }

        }

        private void GetContentFromClipboard(ref byte[] bytes, ref string fileName)
        {
            if (Clipboard.ContainsImage())
            {
                bytes = ConvertImageToBytes(out fileName);
            }
            if (Clipboard.ContainsFileDropList())
            {
                bytes = ConvertFilesToBytes(out fileName);
            }
            if (Clipboard.ContainsText())
            {
                bytes = ConvertTextToBytes(out fileName);
            }
        }

        private byte[] ConvertImageToBytes(out string fileName)
        {
            var content = Clipboard.GetImage();
            byte[] bytes = ConvertImageToBytes(content);
            fileName = "CapturedImage.png";
            return bytes;
        }

        private byte[] ConvertFilesToBytes(out string fileName)
        {
            var files = Clipboard.GetFileDropList();
            fileName = files[0];
            if (files.Count > 1)
            {
                string zipFileName = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + $"\\temp\\CompressedContent.zip";
                if (File.Exists(zipFileName))
                    File.Delete(zipFileName);
                ZipUtils.ZipFiles(files.Cast<string>().ToList(), zipFileName, true);
                fileName = zipFileName;
            }
            byte[] bytes = File.ReadAllBytes(fileName);
            fileName = fileName.Substring(fileName.LastIndexOf("\\") + 1);
            return bytes;
        }

        private byte[] ConvertTextToBytes(out string fileName)
        {
            var text = Clipboard.GetText();
            byte[] bytes = new UnicodeEncoding().GetBytes(text);
            fileName = "CapturedText.txt";
            return bytes;
        }

        private static bool ValidClipboard()
        {
            return Clipboard.ContainsImage() || Clipboard.ContainsText() || Clipboard.ContainsFileDropList();
        }

        private void SetNotification(string activeApplication)
        {
            var message = "Xrm.Page.ui.setFormNotification('Attachment created. Please check !!', 'INFO', 'PASTEINFOMESSAGE');setTimeout(function(){ Xrm.Page.ui.clearFormNotification('PASTEINFOMESSAGE')},5000);";
            FireRequestAction(new RequestActionEventArgs(activeApplication, "RunXrmCommand", message));
        }

        private void ClearNotification(string activeApplication)
        {
            var message = "Xrm.Page.ui.clearFormNotification('PASTEINFOMESSAGE');";
            FireRequestAction(new RequestActionEventArgs(activeApplication, "RunXrmCommand", message));
        }

        private void CreateAnnotation(string subject, string fileName, byte[] data, string parentEnt, Guid parentEntId)
        {
            try
            {
                Entity note = new Entity("annotation");
                note["subject"] = subject;
                note["filename"] = fileName;
                note["documentbody"] = Convert.ToBase64String(data);
                note["objectid"] = new EntityReference(parentEnt, parentEntId);

                IOrganizationService _orgService = (IOrganizationService)_client.CrmInterface.OrganizationWebProxyClient != null ? (IOrganizationService)_client.CrmInterface.OrganizationWebProxyClient : (IOrganizationService)_client.CrmInterface.OrganizationServiceProxy;
                _orgService.Create(note);
            }
            catch
            {
                throw;
            }

        }

        private string GetLogicalName(string applicationName)
        {
            var logicalName = ((DynamicsCustomerRecord)((AgentDesktopSession)localSessionManager.ActiveSession).Customer.DesktopCustomer).GetReplaceableParameter(applicationName, "LogicalName");

            if (string.IsNullOrEmpty(logicalName))
            {
                throw new Exception("Invalid LogicalName of the Active Application.", new Exception($"Activate application name: {applicationName}, Guid found for application : {logicalName}"));
            }

            return logicalName;
        }

        private Guid GetIdOfApplication(string applicationName)
        {
            
            var id=((DynamicsCustomerRecord)((AgentDesktopSession)localSessionManager.ActiveSession).Customer.DesktopCustomer).GetReplaceableParameter(applicationName, "Id");

            if (string.IsNullOrEmpty(id) || id==Guid.Empty.ToString())
            {
                throw new Exception("Invalid Guid of the Active Application.", new Exception($"Activate application name: {applicationName}, Guid found for application : {id}"));
            }

            return Guid.Parse(id);
        }

        private byte[] ConvertImageToBytes(BitmapSource content)
        {
            try
            {
                byte[] bytes;
                PngBitmapEncoder encoder = new PngBitmapEncoder();
                using (MemoryStream stream = new MemoryStream())
                {
                    encoder.Frames.Add(BitmapFrame.Create(content));
                    encoder.Save(stream);
                    bytes = stream.ToArray();
                }

                return bytes;
            }
            catch {
                throw;
            }
        
        }

        #endregion
    }
}
