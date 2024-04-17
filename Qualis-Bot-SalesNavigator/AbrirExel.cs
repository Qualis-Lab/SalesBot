/*
 * Created by Ranorex
 * User: Nicolas Risso
 * Date: 1/4/2023
 * Time: 5:44 PM
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
    /// Description of AbrirExel.
    /// </summary>
    [TestModule("05CDCE57-A03D-4016-9BB2-37E1A6A30879", ModuleType.UserCode, 1)]
    public class AbrirExel : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public AbrirExel()
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
            
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            RepoVar.setExelInstance(excelApp);
        }
    }
}
