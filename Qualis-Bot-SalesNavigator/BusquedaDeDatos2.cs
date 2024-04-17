/*
 * Created by Ranorex
 * User: Ezequiel Rodas
 * Date: 14/12/2022
 * Time: 12:03
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
	/// Description of BusquedaDeDatos2.
	/// </summary>
	[TestModule("40575EB3-5723-4E48-83E1-1AF720E5886A", ModuleType.UserCode, 1)]
	public class BusquedaDeDatos2 : ITestModule
	{
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public BusquedaDeDatos2()
		{
			// Do not delete - a parameterless constructor is required!
		}

		string _cantidadDeResultados = "";
		[TestVariable("3662b0f8-6b68-45f3-b76b-0533b4e2f0c6")]
		public string cantidadDeResultados
		{
			get { return _cantidadDeResultados; }
			set { _cantidadDeResultados = value; }
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
			
			Mouse.DefaultMoveTime = 300;
			Keyboard.DefaultKeyPressTime = 100;
			Delay.SpeedFactor = 1.0;
			
			int cantidadIteraciones = Convert.ToInt32(cantidadDeResultados);
			
			string nombrePersona, linkPersona, puestoPersona; //Declaracion de variables de los datos que necesitamos
			
			string header = "NombrePersona,PuestoPersona,LinkPersona";
			string filePath = @"D:\TEMP\aux.csv";
			
			var excelApp = new Microsoft.Office.Interop.Excel.Application();
			
			//Abro el archivo excel
			excelApp.Workbooks.Open(@"D:\TEMP\OutputEmpresas.xlsx").Activate();
			//En caso que funcione sera esta la direccion donde se ecuentre el Excel
			//excelApp.Workbooks.Open(@"C:%userprofile%\TEMP\OutputEmpresas.xlsx").Activate();
			excelApp.Visible=true;
			
			/*
			if (!System.IO.File.Exists(filePath)){
				System.IO.File.WriteAllText(filePath, header);
			}*/
			
			
			if(cantidadIteraciones > 25){
				cantidadIteraciones = 25;
			}
			
			if (cantidadIteraciones == 0){
				Report.Info("Sin resultados");
			}else{
				
				for(int indice=1; indice<=cantidadIteraciones; indice++){
					
					//Host.Local.FindSingle("/dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-anonymize='person-name']").Focus();

					nombrePersona = Host.Local.FindSingle("./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']/span").GetAttributeValueText("InnerText");
					
					linkPersona = Host.Local.FindSingle("./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']").GetAttributeValueText("href");
					
					puestoPersona= Host.Local.FindSingle("/dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//span[@data-anonymize='title']").GetAttributeValueText("InnerText");
					
					Report.Info(nombrePersona+", "+linkPersona+","+puestoPersona);
					
					Mouse.MoveTo(Host.Local.FindSingle("./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']/span"));
					
					Mouse.ScrollWheel(-210);
					
				}
				
			}
		}
	}
}
