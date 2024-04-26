/*
 * Created by Ranorex
 * User: Juan Cruz Diana
 * Date: 18/4/2024
 * Time: 16:14
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

namespace Qualis_Bot_SalesNavigator
{
    /// <summary>
    /// Description of ResetEjecutar.
    /// </summary>
    [TestModule("FACC1CDB-7DDC-4CF8-A5B5-B4CE299B65A1", ModuleType.UserCode, 1)]
    public class ResetEjecutar : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public ResetEjecutar()
        {
            // Do not delete - a parameterless constructor is required!
        }

string _Ejecutar = "";
[TestVariable("842316f2-1687-44f6-a15b-76e804abb3d0")]
public string Ejecutar
{
	get { return _Ejecutar; }
	set { _Ejecutar = value; }
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
            
            Ejecutar = "True";
        }
    }
}
