/*
 * Created by Ranorex
 * User: Ezequiel Rodas
 * Date: 13/12/2022
 * Time: 11:53
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
using System.Runtime.InteropServices;

namespace Qualis_Bot_SalesNavigator
{
	/// <summary>
	/// Description of BusquedaDeDatos.
	/// </summary>
	[TestModule("64F9CDA0-2E81-475E-8E11-B5B9CA41869A", ModuleType.UserCode, 1)]
	public class BusquedaDeDatos : ITestModule
	{
		/// <summary>
		/// Constructs a new instance.
		/// </summary>
		public BusquedaDeDatos()
		{
			// Do not delete - a parameterless constructor is required!
		}

		string _pais = "";
		[TestVariable("62f84f30-00b5-4465-b51e-f7702ed031c9")]
		public string pais
		{
			get { return _pais; }
			set { _pais = value; }
		}

		string _nombreEmpresa = "";
		[TestVariable("bb30c447-9719-41bd-9c4b-b991008f434c")]
		public string nombreEmpresa
		{
			get { return _nombreEmpresa; }
			set { _nombreEmpresa = value; }
		}

		string _industria = "";
		[TestVariable("627d7d90-9373-42aa-adad-bc0a1c21155a")]
		public string industria
		{
			get { return _industria; }
			set { _industria = value; }
		}

		string _keyword = "";
		[TestVariable("2a3ee7f5-a412-42a4-8543-f1ea657adeeb")]
		public string keyword
		{
			get { return _keyword; }
			set { _keyword = value; }
		}

		string _cantidadDeResultados = "";
		[TestVariable("fa433e03-21e2-455a-bb97-366a08c97969")]
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
			
			int cantidadIteraciones= Convert.ToInt32(cantidadDeResultados);
			int restoResultado= Convert.ToInt32(cantidadDeResultados);
			
			string nombrePersona, linkPersona, puestoPersona;
			Element elem;
			
			if(cantidadIteraciones > 25){
				cantidadIteraciones = 25;
			}
			
			//setea el valor de los resultados totales y le resta 25 por iteracion de la folder
			if (cantidadIteraciones<restoResultado){
				restoResultado = restoResultado - 25;
				cantidadDeResultados = restoResultado.ToString();
			}
			
			if (cantidadIteraciones == 0){
				Report.Info("Sin resultados para el cargo "+ keyword);
			}else{
				
				for(int indice=1; indice<=cantidadIteraciones; indice++){
					
					string pathNombre = "./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']/span";
					int x = 0;
					//Host.Local.FindSingle("/dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-anonymize='person-name']").Focus();

					//nombrePersona = Host.Local.FindSingle("./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']/span").GetAttributeValueText("InnerText");
					
					//linkPersona = Host.Local.FindSingle("./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']").GetAttributeValueText("href");
					
					//puestoPersona= Host.Local.FindSingle("/dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//span[@data-anonymize='title']").GetAttributeValueText("InnerText");
					
					//Report.Info(nombrePersona+", "+linkPersona+","+puestoPersona);
					
					//Mouse.MoveTo(Host.Local.FindSingle("./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']/span"));
					
					//GuardarExcel(pais, industria, nombreEmpresa, nombrePersona, puestoPersona, linkPersona, "DatosObtenidos");
					
					//if (puestoPersona.Contains(keyword)){
					//	GuardarExcel(pais, industria, nombreEmpresa, nombrePersona, puestoPersona, linkPersona, "DatosFiltrados");
					
					//}
					
					
					while(x < 2){
						if(Host.Local.TryFindSingle(pathNombre, out elem)){
							
							nombrePersona = elem.GetAttributeValueText("InnerText");
							
							linkPersona = Host.Local.FindSingle("./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']").GetAttributeValueText("href");
							
							try{
								puestoPersona= Host.Local.FindSingle("/dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//span[@data-anonymize='title']").GetAttributeValueText("InnerText");
							}
							catch(Exception)
							{ puestoPersona = "No posee";}
							
							Report.Info(nombrePersona+", "+linkPersona+","+puestoPersona);
							
//							Mouse.MoveTo(Host.Local.FindSingle("./dom[@domain='www.linkedin.com']//div[#'search-results-container']/?/?/ol/li["+Convert.ToString(indice)+"]//a[@data-control-name='view_lead_panel_via_search_lead_name']/span"));
							
							GuardarExcel(pais, industria, nombreEmpresa, nombrePersona, puestoPersona, linkPersona, "DatosObtenidos");
							
							if (puestoPersona.Contains(keyword)){
//								"DatosObtenidos" EliminarDato(nombrePersona) == puestoPersona.Contains(keyword);
								XlsWorker.EliminarDato(nombrePersona, "DatosObtenidos");
															
								GuardarExcel(pais, industria, nombreEmpresa, nombrePersona, puestoPersona, linkPersona, "DatosFiltrados");
								
							}
							
							break;
						}
						else{
//							Mouse.ScrollWheel(-200);
							x++;
						}
					}
					
//					Mouse.ScrollWheel(-200);
					
				}
				//setea el valor de cantidadIteraciones por el sobrante del restoResultado
				if (restoResultado<cantidadIteraciones){
					cantidadIteraciones = restoResultado;
				}
				
			}
			
		}
		
		
		private void GuardarExcel(string pais, string industria, string nombreEmpresa, string nombreCompleto, string puestoPersona, string linkedin, string hoja)
		{
			
			//var excelApp = new Microsoft.Office.Interop.Excel.Application();
			Microsoft.Office.Interop.Excel.Application excelApp = XlsWorker.getExelInstance();
			
			//Abro el archivo excel
			excelApp.Workbooks.Open(@"C:\TEMP\OutputEmpresas.xlsx").Activate();
			excelApp.Visible=false;
			
			//Selecciono el WorkSheet
			((Microsoft.Office.Interop.Excel.Worksheet)excelApp.ActiveWorkbook.Sheets[hoja]).Select();
			
			int fila=1;
			//string ultimaFilaEscrita=excelApp.Range["A"+fila].Value;
			
			//Tomo la ultima fila escrita en el documento Excel
			while(excelApp.Range["A"+fila].Value != null){
				fila=fila+1;
				
			}
			
			int filaAEscribir=fila;
			
			//Escribo el pais
			excelApp.Range["A"+filaAEscribir].Value = pais;
			
			//Escribo la industria
			excelApp.Range["B"+filaAEscribir].Value = industria;
			
			//Escribo el Nombre Empresa
			excelApp.Range["C"+filaAEscribir].Value = nombreEmpresa;
			
			//Escribo el Nombre Completo
			excelApp.Range["D"+filaAEscribir].Value = nombreCompleto;
			
			//Escribo el Rol
			excelApp.Range["E"+filaAEscribir].Value = puestoPersona;
			
			//Escribo el linkedin
			excelApp.Range["F"+filaAEscribir].Value = linkedin;
			
			//Guardo el excel y lo cierro
			excelApp.ActiveWorkbook.Save();
			//Cierro el Excel
			excelApp.Workbooks.Close();
			

		}
	}
	
}


