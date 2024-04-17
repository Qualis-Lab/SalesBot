/*
 * Created by Ranorex
 * User: Juan Cruz Diana
 * Date: 16/4/2024
 * Time: 17:26
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
//    /// <summary>
//    /// Description of PagSiguiente1.
//    /// </summary>
//    [TestModule("55C47D8E-5A41-4A21-B6C1-C6664A3FDABE", ModuleType.UserCode, 1)]
//    public class PagSiguiente1 : ITestModule
//    {
//    	
//        public static Qualis_Bot_SalesNavigatorRepository repo = Qualis_Bot_SalesNavigatorRepository.Instance;
//      
////        /// <summary>
//        /// Constructs a new instance.
//        /// </summary>
//        public PagSiguiente1()
//        {
//       }
//
//        /// <summary>
//        /// Performs the playback of actions in this module.
//        /// </summary>
//        /// <remarks>You should not call this method directly, instead pass the module
//        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
//        /// that will in turn invoke this method.</remarks>
//        void ITestModule.Run()
//        {
//            Mouse.DefaultMoveTime = 300;
//            Keyboard.DefaultKeyPressTime = 100;
//            Delay.SpeedFactor = 1.0;
//            
//             {
//			int aux = int.Parse(CantPage); 
//			
//			if (aux > 1){
//				
//				repo.ContinueAndFail.ButtonSiguiente.Click();
//				repo.ContinueAndFail.Li_ExistUnResultadoInfo.WaitForExists(10000);
//				aux--;
//				CantPage = aux.ToString();
//			}
//			else{
//				Report.Info("Sin pagina siguiente.");
//			}
//        }
//
//    }
//}
}
      