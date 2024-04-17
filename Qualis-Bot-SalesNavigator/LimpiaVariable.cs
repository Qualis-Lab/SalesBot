/*
 * Created by Ranorex
 * User: Nicolas Risso
 * Date: 1/16/2023
 * Time: 1:05 PM
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
    /// Description of LimpiaVariable.
    /// </summary>
    [TestModule("9D1B3FEE-FFE6-4777-911B-49B6A1E00D1D", ModuleType.UserCode, 1)]
    public class LimpiaVariable : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public LimpiaVariable()
        {
            // Do not delete - a parameterless constructor is required!
        }

        string _ValidarPaisNombreEmpresa = "";
        [TestVariable("75fb9077-cb8e-4f9f-94ec-699d6ef80eb3")]
        public string ValidarPaisNombreEmpresa
        {
        	get { return _ValidarPaisNombreEmpresa; }
        	set { _ValidarPaisNombreEmpresa = value; }
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
            
            ValidarPaisNombreEmpresa = string.Empty;
        }
    }
}
