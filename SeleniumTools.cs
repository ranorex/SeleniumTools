/*
 * Created by Ranorex
 * User: copresnik
 * Date: 07.03.2017
 * Time: 16:13
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.Diagnostics;
using System.Xml;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Reporting;
using Ranorex.Core.Testing;

namespace RanorexDoesSelenium
{
    /// <summary>
    /// Description of SeleniumTools.
    /// </summary>
    [TestModule("B92BB854-2418-41AD-9FBB-F630562D0A98", ModuleType.UserCode, 1)]
    public class SeleniumTools : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public SeleniumTools()
        {
            // Do not delete - a parameterless constructor is required!
        }

        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
        }
        
        /// <summary>
        /// Starts an external process and waits for it to finish
        /// </summary>
        /// <param name="command">Process to be executed</param>
        /// <param name="args">Arguments for the process</param>
        /// <param name="workdir">Working Directory</param>
        public void RunTestSynchronized(string command, string args, string workdir)
		{
			Report.Info("Selenium", "Starting " + command + " with Arguments " + args + " in " + workdir);
			
			ProcessStartInfo psi = new ProcessStartInfo();
			psi.FileName = command;
			psi.Arguments = args;
			psi.WorkingDirectory = workdir;
			
			try
			{
				var proc = Process.Start(psi);
				proc.WaitForExit();
				Report.Info("Selenium", "Done!");
			}
			catch(Exception e)
			{
				Report.Error("Selenium", "Could not start Selenium Test\n" + e.Message);
			}
		}
		
        /// <summary>
        /// Opens a JUNIT report and adds its results to the Ranorex Report
        /// </summary>
        /// <param name="workdir">The directory where the report file is located</param>
        /// <param name="fileName">Name of the report file</param>
		public void ParseExtTestResults(string workdir, string fileName)
		{			
			XmlDocument xdoc = new XmlDocument();
			
			try
			{
				xdoc.Load(workdir + fileName);
				
				XmlNode xtestsuite = xdoc.SelectSingleNode("/testsuite");
				XmlNodeList testcases = xdoc.SelectNodes("/testsuite/testcase");
				
				string sHead = string.Format("External Selenium Test completed. {0} out of {1} test cases successful",
				                             Int32.Parse(xtestsuite.Attributes["tests"].Value) - Int32.Parse(xtestsuite.Attributes["failures"].Value),
				                             Int32.Parse(xtestsuite.Attributes["tests"].Value));
				
				TestReport.BeginTestEntryContainer(1, "Selenium Test Suite (external)",  ActivityExecType.Execute, TestEntryActivityType.TestCase);
				TestReport.BeginSmartFolderContainer("Selenium", sHead);
				
				Report.Info("Selenium", "Parsing Results...");
				
				
				foreach(XmlNode n in testcases)
				{
					Ranorex.Core.Reporting.TestReport.BeginTestCaseContainer(n.Attributes["classname"].Value + "." + n.Attributes["name"].Value);
					Ranorex.Core.Reporting.TestReport.BeginTestModule("Module " + n.Attributes["name"].Value);
					
					if(n.SelectSingleNode("failure") != null)
					{
						Report.Failure("Selenium", "Test case " + n.Attributes["name"].Value + " (" + n.Attributes["classname"].Value + ") \n" +
						               n.SelectSingleNode("failure").Attributes["message"].Value);
						
						Ranorex.Core.Reporting.TestReport.EndTestModule();
						Ranorex.Core.Reporting.TestReport.EndTestCaseContainer(TestResult.Failed);
						continue;
					}
					
					
					Report.Success("Selenium", "Test case " + n.Attributes["name"].Value + " (" + n.Attributes["classname"].Value + ")");
					
					Ranorex.Core.Reporting.TestReport.EndTestModule();
					Ranorex.Core.Reporting.TestReport.EndTestCaseContainer(TestResult.Passed);
				}
				
				TestReport.EndTestCaseContainer();
				TestReport.EndTestEntryContainer();
			}
			catch(Exception e)
			{
				Report.Error("Selenium", "Error reading Results\n"+e.Message);
			}
		}
    }
}
