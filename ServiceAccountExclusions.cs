using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using System.Globalization;
using Microsoft.Win32;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using ActiveDs;

namespace ServiceAccountExclusions
{
    class SAEMain
    {
        struct ServiceAccountParams
        {
            public string strExclusionGroupLocation;
        }
        
        struct CMDArguments
        {
            public bool bParseCmdArguments;
        }

        static ManagementObjectCollection funcSysQueryData(string sysQueryString, string sysServerName)
        {

            // [Comment] Connect to the server via WMI
            System.Management.ConnectionOptions objConnOptions = new System.Management.ConnectionOptions();
            string strServerNameforWMI = "\\\\" + sysServerName + "\\root\\cimv2";

            // [DebugLine] Console.WriteLine("Construct WMI scope...");
            System.Management.ManagementScope objManagementScope = new System.Management.ManagementScope(strServerNameforWMI, objConnOptions);

            // [DebugLine] Console.WriteLine("Construct WMI query...");
            System.Management.ObjectQuery objQuery = new System.Management.ObjectQuery(sysQueryString);
            //if (objQuery != null)
            //    Console.WriteLine("objQuery was created successfully");

            // [DebugLine] Console.WriteLine("Construct WMI object searcher...");
            System.Management.ManagementObjectSearcher objSearcher = new System.Management.ManagementObjectSearcher(objManagementScope, objQuery);
            //if (objSearcher != null)
            //    Console.WriteLine("objSearcher was created successfully");

            // [DebugLine] Console.WriteLine("Get WMI data...");

            System.Management.ManagementObjectCollection objReturnCollection = null;

            try
            {
                objReturnCollection = objSearcher.Get();
                return objReturnCollection;
            }
            catch (SystemException ex)
            {
                // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
                string strRPCUnavailable = "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)";
                // [DebugLine] System.Console.WriteLine(ex.Message);
                if (ex.Message == strRPCUnavailable)
                {
                    Console.WriteLine("WMI: Server unavailable");
                }

                // Next line will return an object that is equal to null
                return objReturnCollection;
            }
        }

        static bool funcPingServer(string strServerName)
        {
            bool bPingSuccess = false;
            //bool bWMISuccess = false;

            // [DebugLine] Console.WriteLine("Contact start for {0}: {1}", strServerName, DateTime.Now.ToLocalTime().ToString("MMddyyy HH:mm:ss"));

            //string strServerNameforWMI = "";

            // [Comment] Ping the server
            // [DebugLine] Console.WriteLine(); // Helper line just to make output clearer
            // [DebugLine] Console.WriteLine("Ping attempt for: " + strServerName);

            try
            {
                System.Net.NetworkInformation.Ping objPing1 = new System.Net.NetworkInformation.Ping();
                System.Net.NetworkInformation.PingReply objPingReply1 = objPing1.Send(strServerName);
                if (objPingReply1.Status.ToString() != "TimedOut")
                {
                    // [DebugLine] Console.WriteLine("Ping Reply: " + objPingReply1.Address + "     RTT: " + objPingReply1.RoundtripTime);
                    bPingSuccess = true;
                }
                else
                {
                    // [DebugLine] Console.WriteLine("Ping Reply: " + objPingReply1.Status);
                    bPingSuccess = false;
                }
            }
            catch (SystemException ex)
            {
                //// [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
                //string strPingError = "An exception occurred during a Ping request.";
                //// [DebugLine] System.Console.WriteLine(ex.Message);
                //if (ex.Message == strPingError)
                //{
                //    Console.WriteLine("Ping Error for {0}. No ip address was found during name resolution.", strServerName);
                //}

                bPingSuccess = false;
            }

            //// [Comment] Connect to the server via WMI
            //// [DebugLine] Console.WriteLine(); // Helper line just to make output clearer
            //Console.WriteLine("WMI connection attempt for: " + strServerName);

            //System.Management.ConnectionOptions objConnOptions = new System.Management.ConnectionOptions();
            //strServerNameforWMI = "\\\\" + strServerName + "\\root\\cimv2";

            //// [DebugLine] Console.WriteLine("Construct WMI scope...");
            //System.Management.ManagementScope objManagementScope = new System.Management.ManagementScope(strServerNameforWMI, objConnOptions);
            //// [DebugLine] Console.WriteLine("Construct WMI query...");
            //System.Management.ObjectQuery objQuery = new System.Management.ObjectQuery("select * from Win32_ComputerSystem");
            //// [DebugLine] Console.WriteLine("Construct WMI object searcher...");
            //System.Management.ManagementObjectSearcher objSearcher = new System.Management.ManagementObjectSearcher(objManagementScope, objQuery);
            //Console.WriteLine("Get WMI data...");

            //try
            //{
            //    System.Management.ManagementObjectCollection objObjCollection = objSearcher.Get();

            //    foreach (System.Management.ManagementObject objMgmtObject in objObjCollection)
            //    {
            //        Console.WriteLine("Hostname: " + objMgmtObject["Caption"].ToString());
            //        bWMISuccess = true;
            //    }
            //}
            //catch (SystemException ex)
            //{
            //    // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
            //    string strRPCUnavailable = "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)";
            //    // [DebugLine] System.Console.WriteLine(ex.Message);
            //    if (ex.Message == strRPCUnavailable)
            //    {
            //        Console.WriteLine("WMI: Server unavailable");
            //    }
            //    bWMISuccess = false;
            //}

            // [DebugLine] Console.WriteLine("Contact stop for {0}: {1}", strServerName, DateTime.Now.ToLocalTime().ToString("MMddyyy HH:mm:ss"));

            if (bPingSuccess)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static void funcPrintParameterWarning()
        {
            Console.WriteLine("A parameter is missing or is incorrect.");
            Console.WriteLine("Run ServiceAccountExclusions -? to get the parameter syntax.");
        }

        static void funcPrintParameterSyntax()
        {
            Console.WriteLine("ServiceAccountExclusions (c) 2011 SystemsAdminPro.com");
            Console.WriteLine();
            Console.WriteLine("Description: Find service accounts and create exclusions");
            Console.WriteLine();
            Console.WriteLine("Parameter syntax:");
            Console.WriteLine();
            Console.WriteLine("Use the following for the first parameter:");
            Console.WriteLine("-run                required parameter");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("ServiceAccountExclusions -run");
        } // funcPrintParameterSyntax

        static void funcLogToEventLog(string strAppName, string strEventMsg, int intEventType)
        {
            string sLog;

            sLog = "Application";

            if (!EventLog.SourceExists(strAppName))
                EventLog.CreateEventSource(strAppName, sLog);

            //EventLog.WriteEntry(strAppName, strEventMsg);
            EventLog.WriteEntry(strAppName, strEventMsg, EventLogEntryType.Information, intEventType);

        } // LogToEventLog

        static DirectorySearcher funcCreateDSSearcher()
        {
            try
            {
                System.DirectoryServices.DirectorySearcher objDSSearcher = new DirectorySearcher();
                // [Comment] Get local domain context

                string rootDSE;

                System.DirectoryServices.DirectorySearcher objrootDSESearcher = new System.DirectoryServices.DirectorySearcher();
                rootDSE = objrootDSESearcher.SearchRoot.Path;
                //Console.WriteLine(rootDSE);

                // [Comment] Construct DirectorySearcher object using rootDSE string
                System.DirectoryServices.DirectoryEntry objrootDSEentry = new System.DirectoryServices.DirectoryEntry(rootDSE);
                objDSSearcher = new System.DirectoryServices.DirectorySearcher(objrootDSEentry);
                //Console.WriteLine(objDSSearcher.SearchRoot.Path);

                return objDSSearcher;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }
        }

        static PrincipalContext funcCreatePrincipalContext()
        {
            PrincipalContext newctx = new PrincipalContext(ContextType.Machine);

            try
            {
                //Console.WriteLine("Entering funcCreatePrincipalContext");
                Domain objDomain = Domain.GetComputerDomain();
                string strDomain = objDomain.Name;
                DirectorySearcher tempDS = funcCreateDSSearcher();
                string strDomainRoot = tempDS.SearchRoot.Path.Substring(7);
                // [DebugLine] Console.WriteLine(strDomainRoot);
                // [DebugLine] Console.WriteLine(strDomainRoot);

                newctx = new PrincipalContext(ContextType.Domain,
                                    strDomain,
                                    strDomainRoot);

                // [DebugLine] Console.WriteLine(newctx.ConnectedServer);
                // [DebugLine] Console.WriteLine(newctx.Container);



                //if (strContextType == "Domain")
                //{

                //    PrincipalContext newctx = new PrincipalContext(ContextType.Domain,
                //                                    strDomain,
                //                                    strDomainRoot);
                //    return newctx;
                //}
                //else
                //{
                //    PrincipalContext newctx = new PrincipalContext(ContextType.Machine);
                //    return newctx;
                //}
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

            if (newctx.ContextType == ContextType.Machine)
            {
                Exception newex = new Exception("The Active Directory context did not initialize properly.");
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, newex);
            }

            return newctx;
        }

        static void funcProgramRegistryTag(string strProgramName)
        {
            try
            {
                string strRegistryProfilesPath = "SOFTWARE";
                RegistryKey objRootKey = Microsoft.Win32.Registry.LocalMachine;
                RegistryKey objSoftwareKey = objRootKey.OpenSubKey(strRegistryProfilesPath, true);
                RegistryKey objSystemsAdminProKey = objSoftwareKey.OpenSubKey("SystemsAdminPro", true);
                if (objSystemsAdminProKey == null)
                {
                    objSystemsAdminProKey = objSoftwareKey.CreateSubKey("SystemsAdminPro");
                }
                if (objSystemsAdminProKey != null)
                {
                    if (objSystemsAdminProKey.GetValue(strProgramName) == null)
                        objSystemsAdminProKey.SetValue(strProgramName, "1", RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static CMDArguments funcParseCmdArguments(string[] cmdargs)
        {
            CMDArguments objCMDArguments = new CMDArguments();

            try
            {
                objCMDArguments.bParseCmdArguments = false;

                if (cmdargs[0] == "-run" & cmdargs.Length == 1)
                {
                    objCMDArguments.bParseCmdArguments = true;
                }
                else
                {
                    objCMDArguments.bParseCmdArguments = false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                objCMDArguments.bParseCmdArguments = false;
            }

            return objCMDArguments;
        }

        static ServiceAccountParams funcParseConfigFile(CMDArguments objCMDArguments)
        {
            ServiceAccountParams newParams = new ServiceAccountParams();

            TextReader trConfigFile = new StreamReader("configServiceAccountExclusions.txt");

            using (trConfigFile)
            {
                string strNewLine = "";

                while ((strNewLine = trConfigFile.ReadLine()) != null)
                {
                    if (strNewLine.StartsWith("ExclusionGroupLocation=") & strNewLine != "ExclusionGroupLocation=")
                    {
                        newParams.strExclusionGroupLocation = strNewLine.Substring(23);
                        //[DebugLine] Console.WriteLine(newParams.strExclusionGroupLocation);
                    }
                }
            }

            //[DebugLine] Console.WriteLine("# of Exclude= : {0}", newParams.lstExclude.Count.ToString());
            //[DebugLine] Console.WriteLine("# of ExcludePrefix= : {0}", newParams.lstExcludePrefix.Count.ToString());

            trConfigFile.Close();

            return newParams;
        }

        static void funcProgramExecution(CMDArguments objCMDArguments2)
        {
            try
            {
                funcProgramRegistryTag("ServiceAccountExclusions");

                funcLogToEventLog("ServiceAccountExclusions", "ServiceAccountExclusions started", 100);

                TextWriter twOutputLog = funcOpenOutputLog();
                string strOutputMsg = String.Empty;

                strOutputMsg = "--------ServiceAccountExclusions started";
                funcWriteToOutputLog(twOutputLog, strOutputMsg);

                ServiceAccountParams newParams = funcParseConfigFile(objCMDArguments2);

                PrincipalContext ctxDomain = funcCreatePrincipalContext();
                
                Domain dmCurrent = Domain.GetCurrentDomain();
                PrincipalContext ctxExclusionGroupLocation = new PrincipalContext(ContextType.Domain, dmCurrent.Name, newParams.strExclusionGroupLocation);

                GroupPrincipal grpServiceAccountExclusions = GroupPrincipal.FindByIdentity(ctxExclusionGroupLocation, IdentityType.SamAccountName, "ServiceAccountExclusions");

                bool bRecreateGroup = false;

                if (grpServiceAccountExclusions != null)
                {
                    DirectoryEntry grpDE = (DirectoryEntry)grpServiceAccountExclusions.GetUnderlyingObject();
                    string strGrpCreated = funcGetAccountCreationDate(grpDE);
                    DateTime dtGrpCreated = Convert.ToDateTime(strGrpCreated);
                    if (dtGrpCreated < DateTime.Today.AddDays(-14))
                    {
                        grpDE.Close();
                        grpServiceAccountExclusions.Delete();
                        bRecreateGroup = true;
                    }
                }

                if (grpServiceAccountExclusions == null | bRecreateGroup)
                {
                    grpServiceAccountExclusions = new GroupPrincipal(ctxExclusionGroupLocation);
                    grpServiceAccountExclusions.Name = "ServiceAccountExclusions";
                    grpServiceAccountExclusions.SamAccountName = "ServiceAccountExclusions";
                    grpServiceAccountExclusions.Description = "SystemsAdminPro Exclusions";
                    grpServiceAccountExclusions.Save();
                }

                string strQueryFilter = "(&(&(&(sAMAccountType=805306369)(objectCategory=computer)(|(operatingSystem=Windows Server 2008*)(operatingSystem=Windows Server 2003*)(operatingSystem=Windows 2000 Server*)(operatingSystem=Windows NT*)(operatingSystem=*2008*)))))";

                List<string> lstExcludeServiceName = new List<string>();

                lstExcludeServiceName.Add("LocalSystem");
                lstExcludeServiceName.Add("NT AUTHORITY\\LocalService");
                lstExcludeServiceName.Add("NT Authority\\LocalService");
                lstExcludeServiceName.Add("NT AUTHORITY\\LOCALSERVICE");
                lstExcludeServiceName.Add("NT AUTHORITY\\LOCAL SERVICE");
                lstExcludeServiceName.Add("NT AUTHORITY\\NetworkService");
                lstExcludeServiceName.Add("NT Authority\\NetworkService");
                lstExcludeServiceName.Add("NT AUTHORITY\\NETWORKSERVICE");
                lstExcludeServiceName.Add("NT AUTHORITY\\NETWORK SERVICE");
                lstExcludeServiceName.Add("localSystem");
                //lstExcludeServiceName.Add("");

                System.DirectoryServices.DirectorySearcher objComputerObjectSearcher = funcCreateDSSearcher();
                // [DebugLine]Console.WriteLine(objComputerObjectSearcher.SearchRoot.Path);

                // [Comment] Add filter to DirectorySearcher object
                objComputerObjectSearcher.Filter = (strQueryFilter);

                // [Comment] Execute query, return results, display name and path values
                System.DirectoryServices.SearchResultCollection objComputerResults = objComputerObjectSearcher.FindAll();
                // [DebugLine] Console.WriteLine(objComputerResults.Count.ToString());



                foreach (SearchResult sr in objComputerResults)
                {
                    DirectoryEntry newDE = sr.GetDirectoryEntry();
                    // [DebugLine] Console.WriteLine(newDE.Name.Substring(3));

                    ManagementObjectCollection oQueryCollection = null;
                    string strHostName = "";

                    strHostName = newDE.Name.Substring(3);

                    if (funcPingServer(strHostName))
                    {
                        //**********************************
                        //Begin-Win32_Service
                        //**********************************
                        //Get the query results for Win32_Service
                        oQueryCollection = null;
                        oQueryCollection = funcSysQueryData("select * from Win32_Service", strHostName);

                        // [DebugLine] Console.WriteLine("{0} \t Service Count: {1}", strHostName, oQueryCollection.Count.ToString());
                        strOutputMsg = String.Format("{0} \t Service Count: {1}", strHostName, oQueryCollection.Count.ToString());
                        funcWriteToOutputLog(twOutputLog, strOutputMsg);

                        foreach (ManagementObject oReturn in oQueryCollection)
                        {
                            // "Caption","Description","PathName","Status","State","StartMode","StartName"

                            //string[] strElementBag = new string[] { "Caption", "Description", "PathName", "Status", "State", "StartMode", "StartName" };
                            //foreach (string strElement in strElementBag)
                            //{
                            //    string strElementTemp = strElement.ToLower(new CultureInfo("en-US", false));
                            //    try
                            //    {
                            //        Console.WriteLine("\"" + strElementTemp + "\"" + " : " + "\"" + oReturn[strElement].ToString().Trim() + "\"");
                            //    }
                            //    catch
                            //    {
                            //        Console.WriteLine("\"" + strElementTemp + "\"" + " : " + "\"" + "<na>" + "\"");
                            //    }
                            //}
                            if (!lstExcludeServiceName.Contains(oReturn.Properties["StartName"].Value.ToString()))
                            {
                                string strServiceStartName = oReturn.Properties["StartName"].Value.ToString();
                                if (strServiceStartName.IndexOf("\\") != -1)
                                {
                                    strServiceStartName = strServiceStartName.Substring(strServiceStartName.IndexOf("\\") + 1);
                                }
                                if (strServiceStartName.IndexOf("@") != -1)
                                {
                                    strServiceStartName = strServiceStartName.Substring(0, strServiceStartName.IndexOf("@"));
                                }
                                //[DebugLine] Console.WriteLine(strServiceStartName);
                                string strServiceAccountOutput = String.Format("{0} \t {1} \t {2}", strHostName, oReturn.Properties["Caption"].Value.ToString(), oReturn.Properties["StartName"].Value.ToString());
                                //[DebugLine] Console.WriteLine(strServiceAccountOutput);
                                funcWriteToOutputLog(twOutputLog, strServiceAccountOutput);

                                UserPrincipal upService = UserPrincipal.FindByIdentity(ctxDomain, IdentityType.SamAccountName, strServiceStartName);
                                if (upService != null)
                                {
                                    if (!upService.IsMemberOf(grpServiceAccountExclusions))
                                    {
                                        grpServiceAccountExclusions.Members.Add(upService);
                                        grpServiceAccountExclusions.Save();
                                    }
                                }
                                else
                                {
                                    funcWriteToOutputLog(twOutputLog, strServiceStartName + " could not be found and added to the ServiceAccountExclusions group");
                                }
                            }
                        }
                        //**********************************
                        //End-Win32_Service
                        //**********************************
                    }
                    else
                    {
                        strOutputMsg = String.Format("{0} \t {1}", strHostName, ".");
                        funcWriteToOutputLog(twOutputLog, strOutputMsg);
                    }
                }

                objComputerResults.Dispose();

                strOutputMsg = "--------ServiceAccountExclusions stopped";
                funcWriteToOutputLog(twOutputLog, strOutputMsg);

                funcCloseOutputLog(twOutputLog);

                funcLogToEventLog("ServiceAccountExclusions", "ServiceAccountExclusions stopped", 101);

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcGetFuncCatchCode(string strFunctionName, Exception currentex)
        {
            string strCatchCode = "";

            Dictionary<string, string> dCatchTable = new Dictionary<string, string>();
            dCatchTable.Add("funcGetFuncCatchCode", "f0");
            dCatchTable.Add("funcLicenseCheck", "f1");
            dCatchTable.Add("funcPrintParameterWarning", "f2");
            dCatchTable.Add("funcPrintParameterSyntax", "f3");
            dCatchTable.Add("funcParseCmdArguments", "f4");
            dCatchTable.Add("funcProgramExecution", "f5");
            dCatchTable.Add("funcProgramRegistryTag", "f6");
            dCatchTable.Add("funcCreateDSSearcher", "f7");
            dCatchTable.Add("funcCreatePrincipalContext", "f8");
            dCatchTable.Add("funcCheckNameExclusion", "f9");
            dCatchTable.Add("funcMoveDisabledAccounts", "f10");
            dCatchTable.Add("funcFindAccountsToDisable", "f11");
            dCatchTable.Add("funcCheckLastLogin", "f12");
            dCatchTable.Add("funcRemoveUserFromGroup", "f13");
            dCatchTable.Add("funcToEventLog", "f14");
            dCatchTable.Add("funcCheckForFile", "f15");
            dCatchTable.Add("funcCheckForOU", "f16");
            dCatchTable.Add("funcWriteToErrorLog", "f17");

            if (dCatchTable.ContainsKey(strFunctionName))
            {
                strCatchCode = "err" + dCatchTable[strFunctionName] + ": ";
            }

            //[DebugLine] Console.WriteLine(strCatchCode + currentex.GetType().ToString());
            //[DebugLine] Console.WriteLine(strCatchCode + currentex.Message);

            funcWriteToErrorLog(strCatchCode + currentex.GetType().ToString());
            funcWriteToErrorLog(strCatchCode + currentex.Message);

        }

        static void funcWriteToErrorLog(string strErrorMessage)
        {
            try
            {
                FileStream newFileStream = new FileStream("Err-ServiceAccountExclusions.log", FileMode.Append, FileAccess.Write);
                TextWriter twErrorLog = new StreamWriter(newFileStream);

                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twErrorLog.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strErrorMessage);

                twErrorLog.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static bool funcCheckForOU(string strOUPath)
        {
            try
            {
                string strDEPath = "";

                if (!strOUPath.Contains("LDAP://"))
                {
                    strDEPath = "LDAP://" + strOUPath;
                }
                else
                {
                    strDEPath = strOUPath;
                }

                if (DirectoryEntry.Exists(strDEPath))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static bool funcCheckForFile(string strInputFileName)
        {
            try
            {
                if (System.IO.File.Exists(strInputFileName))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static TextWriter funcOpenOutputLog()
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat2 = "MMddyyyy"; // for log file directory creation

                string strPath = Directory.GetCurrentDirectory();

                if (!Directory.Exists(strPath + "\\Log"))
                {
                    Directory.CreateDirectory(strPath + "\\Log");
                    if (Directory.Exists(strPath + "\\Log"))
                    {
                        strPath = strPath + "\\Log";
                    }
                }
                else
                {
                    strPath = strPath + "\\Log";
                }

                string strLogFileName = strPath + "\\ServiceAccountExclusions" + dtNow.ToString(dtFormat2) + ".log";

                FileStream newFileStream = new FileStream(strLogFileName, FileMode.Append, FileAccess.Write);
                TextWriter twOuputLog = new StreamWriter(newFileStream);

                return twOuputLog;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }

        }

        static void funcWriteToOutputLog(TextWriter twCurrent, string strOutputMessage)
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                // string dtFormat = "MM/dd/yyyy";
                string dtFormat2 = "MM/dd/yyyy HH:mm";
                // string dtFormat3 = "MMddyyyy HH:mm:ss";

                twCurrent.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat2), strOutputMessage);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcCloseOutputLog(TextWriter twCurrent)
        {
            try
            {
                twCurrent.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static string funcGetLastLogonTimestamp(DirectoryEntry tmpDE)
        {
            try
            {
                string strTimestamp = String.Empty;

                if (tmpDE.Properties.Contains("lastLogonTimestamp"))
                {
                    //[DebugLine] Console.WriteLine(u.Name + " has lastLogonTimestamp attribute");
                    IADsLargeInteger lintLogonTimestamp = (IADsLargeInteger)tmpDE.Properties["lastLogonTimestamp"].Value;
                    if (lintLogonTimestamp != null)
                    {
                        DateTime dtLastLogonTimestamp = funcGetDateTimeFromLargeInteger(lintLogonTimestamp);
                        if (dtLastLogonTimestamp != null)
                        {
                            strTimestamp = dtLastLogonTimestamp.ToLocalTime().ToString();
                        }
                        else
                        {
                            strTimestamp = "(null)";
                        }
                    }
                }

                return strTimestamp;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return String.Empty;
            }
        }

        static string funcGetAccountCreationDate(DirectoryEntry tmpDE)
        {
            try
            {
                string strCreationDate = String.Empty;

                if (tmpDE.Properties.Contains("whenCreated"))
                {
                    strCreationDate = (string)tmpDE.Properties["whenCreated"].Value.ToString();
                }

                return strCreationDate;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return String.Empty;
            }
        }

        static DateTime funcGetDateTimeFromLargeInteger(IADsLargeInteger largeIntValue)
        {
            //
            // Convert large integer to int64 value
            //
            long int64Value = (long)((uint)largeIntValue.LowPart +
                     (((long)largeIntValue.HighPart) << 32));

            //
            // Return the DateTime in utc
            //
            // return DateTime.FromFileTimeUtc(int64Value);


            // return in Localtime
            return DateTime.FromFileTime(int64Value);
        }

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    funcPrintParameterWarning();
                }
                else
                {
                    if (args[0] == "-?")
                    {
                        funcPrintParameterSyntax();
                    }
                    else
                    {
                        string[] arrArgs = args;
                        CMDArguments objArgumentsProcessed = funcParseCmdArguments(arrArgs);

                        if (objArgumentsProcessed.bParseCmdArguments)
                        {
                            funcProgramExecution(objArgumentsProcessed);
                        }
                        else
                        {
                            funcPrintParameterWarning();
                        } // check objArgumentsProcessed.bParseCmdArguments
                    } // check args[0] = "-?"
                } // check args.Length == 0
            }
            catch (Exception ex)
            {
                Console.WriteLine("errm0: {0}", ex.Message);
            }
        }
    }
}
