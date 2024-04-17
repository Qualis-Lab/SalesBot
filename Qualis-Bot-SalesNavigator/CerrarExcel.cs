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

namespace Qualis_Bot_SalesNavigator
{
    /// <summary>
    /// Description of CerrarExel.
    /// </summary>
    [TestModule("FA9286B7-A7E0-45E9-A6BD-3AC2F4472153", ModuleType.UserCode, 1)]
    public class CerrarExcel : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public CerrarExcel()
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
            
            XlsWorker.CerrarExcel();
        }
    }
}
