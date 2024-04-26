/*
 * Created by Ranorex
 * User: Nicolas Risso
 * Date: 1/4/2023
 * Time: 5:50 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace Qualis_Bot_SalesNavigator
{
	/// <summary>
	/// Description of CerrarExel.
	/// </summary>
	[TestModule("FA9286B7-A7E0-45E9-A6BD-3AC2F4472153", ModuleType.UserCode, 1)]
	public class CerrarInstanciaExcel : ITestModule
	{
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public CerrarInstanciaExcel()
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
			string processname = "excel";
			Process[] processes = Process.GetProcessesByName(processname);
			if(processes.Length == 0)
			{
				Report.Error(string.Format("Process '{0}' not found.", processname));
			}
			foreach(Process p in processes)
			{
				try
				{
					p.Kill();
					Report.Info("Process killed: "+ p.ProcessName);
					Delay.Duration(1000,false);
				}
				catch(Exception ex)
				{
				}
				
				
			}
		}
	}
}