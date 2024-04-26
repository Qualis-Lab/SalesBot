/*
 * Created by Ranorex
 * User: Juan Cruz Diana
 * Date: 18/4/2024
 * Time: 10:18
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
	/// Description of CalcularCantidadIteraciones.
	/// </summary>
	[TestModule("8DA1031E-3033-45C5-BB7C-3FBCD069830C", ModuleType.UserCode, 1)]
	public class CalcularCantidadIteraciones : ITestModule
	{
		
		public static Qualis_Bot_SalesNavigatorRepository repo = Qualis_Bot_SalesNavigatorRepository.Instance;
		static CrearCantPageCSV instance = new CrearCantPageCSV();

		
		public static CrearCantPageCSV Instance
		{
			get { return instance; }
		}
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		
		
		public CalcularCantidadIteraciones(){}
		string _cantidadDeResultados = "";
		[TestVariable("663c090a-934b-4978-bc57-2046cae1fa34")]
		
		public string cantidadDeResultados
		{
			get { return _cantidadDeResultados; }
			set { _cantidadDeResultados = value; }
		}

		
		string _todosSeleccionados = "";
		[TestVariable("3e237dfd-38d7-42ff-a3a9-4594f2b34897")]
		public string todosSeleccionados
		{
			get { return _todosSeleccionados; }
			set { _todosSeleccionados = value; }
		}
		
		string _cantPage = "0";
		[TestVariable("02309581-2ed1-48f5-8f20-595395edf649")]
		public string cantPage
		{
			get { return _cantPage; }
			set { _cantPage = value; }
		}
//
		
		
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
			
			Report.Log(ReportLevel.Info, "Wait", "Waiting 30s to exist. Associated repository item: 'InicioSalesNavigator.SpanTag3Resultados'", repo.InicioSalesNavigator.SpanTag3ResultadosInfo, new ActionTimeout(30000), new RecordItemIndex(0));
			repo.InicioSalesNavigator.SpanTag3ResultadosInfo.WaitForExists(30000);
			
			Report.Log(ReportLevel.Info, "Get Value", "Getting attribute 'InnerText' from item 'InicioSalesNavigator.SpanTag3Resultados' and assigning the part of its value captured by '[0-9]{1,}' to variable 'cantidadDeResultados'.", repo.InicioSalesNavigator.SpanTag3ResultadosInfo, new RecordItemIndex(1));
			cantidadDeResultados = repo.InicioSalesNavigator.SpanTag3Resultados.Element.GetAttributeValueText("InnerText", new Regex("[0-9]{1,}"));
			Report.Info(cantidadDeResultados);
			Delay.Milliseconds(0);
			
			Report.Log(ReportLevel.Info, "Get Value", "Getting attribute 'Checked' from item 'InicioSalesNavigator.ChekYaEstanSeleccionados' and assigning its value to variable 'TodosSeleccionados'.", repo.InicioSalesNavigator.ChekYaEstanSeleccionadosInfo, new RecordItemIndex(2));
			todosSeleccionados = repo.InicioSalesNavigator.ChekYaEstanSeleccionados.Element.GetAttributeValueText("Checked");
			Delay.Milliseconds(0);
			
			Report.Log(ReportLevel.Info, "User", todosSeleccionados, new RecordItemIndex(3));
			
			Report.Log(ReportLevel.Info, "User", cantidadDeResultados, new RecordItemIndex(4));
			
			float cantResultados = float.Parse(cantidadDeResultados);
			float div = cantResultados/25;
			// Redondeamos el resultado hacia arriba
			double cantidad = 0;
			
			if (div>1){
				cantidad = Math.Ceiling(div);
				cantPage = cantidad.ToString();
				Report.Info("La cantPage es= " + cantPage);
				
			}
			else{
				cantPage = "1";
				cantidad = 1;
				Report.Info("La cantPage es= " + cantPage);        	        }
		}
	}
}