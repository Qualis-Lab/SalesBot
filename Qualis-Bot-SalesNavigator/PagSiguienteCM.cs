/*
 * Created by Ranorex
 * User: Juan Cruz Diana
 * Date: 18/4/2024
 * Time: 13:10
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
	/// Description of UserCodeModule1.
	/// </summary>
	[TestModule("E7A411F1-6262-43C2-BC6E-B1ED2DD6EAAC", ModuleType.UserCode, 1)]
	public class PagSiguienteCM : ITestModule
	{
		
		public static Qualis_Bot_SalesNavigatorRepository repo = Qualis_Bot_SalesNavigatorRepository.Instance;
		
		
		string _CantPage = "";
		[TestVariable("a85fbbc5-fad4-42c6-83ae-918358daf8ab")]
		public string CantPage
		{
			get { return _CantPage; }
			set { _CantPage = value; }
		}
		
		
		string _IteracionActual = "";
		[TestVariable("c8a78dd2-9218-4d73-98ff-1a5b2d39f2a9")]
		public string IteracionActual
		{
			get { return _IteracionActual; }
			set { _IteracionActual = value; }
		}
		
		
		string _Ejecutar = "";
		[TestVariable("f30d2eee-34c3-4f47-b357-c76d0618dbb7")]
		public string Ejecutar
		{
			get { return _Ejecutar; }
			set { _Ejecutar = value; }
		}
		
		
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public PagSiguienteCM()
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
			{
				int aux = int.Parse(CantPage);
				
				if (aux > 1){
					
					repo.ContinueAndFail.ButtonSiguiente.Click();
					repo.ContinueAndFail.Li_ExistUnResultadoInfo.WaitForExists(10000);
					aux--;
					CantPage = aux.ToString();
				}
				else{
					
					Ejecutar = "False";
					Report.Info("Sin pagina siguiente.");
					
				}
			}

		}
	}
}