﻿///////////////////////////////////////////////////////////////////////////////
//
// This file was automatically generated by RANOREX.
// Your custom recording code should go in this file.
// The designer will only add methods to this file, so your custom code won't be overwritten.
// http://www.ranorex.com
//
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Repository;
using Ranorex.Core.Testing;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Qualis_Bot_SalesNavigator
{
	public partial class ExtraerDatos
	{
		/// <summary>
		/// This method gets called right after the recording has been started.
		/// It can be used to execute recording specific initialization code.
		/// </summary>
		private void Init()
		{
			// Your recording specific initialization code goes here.
		}
public void ObtenerValoresEmail()
{
    int index1 = 1;
    string Email = "";
    string ObtenerEmail = "";

    index = index1.ToString();
    
    while(repo.ContAndFail_Chrome.ObtenerEmailNewInfo.Exists())
    {
        ObtenerEmail = repo.ContAndFail_Chrome.ObtenerEmailNew.Element.GetAttributeValue("caption").ToString();
        if(!Email.Contains(ObtenerEmail))
        {
            Report.Info("Estoy en el if validando ObtenerEmail: " + ObtenerEmail + ". No está en la lista de Email: " + Email);
            Email = Email + ObtenerEmail + " , ";
        }
        
        index1++;
        index = index1.ToString();
        Report.Info("index= " + index);
        
        // Verifica si index es igual a 4 y sale del bucle si es así
        if(index == "4")
        {
            break;
        }
    }

    if(Email != "")
    {
        Report.Info("Emails encontrados: " + Email);
        OpenExcel(Email, "G");
    }
    else
    {
        Report.Info("No se encontraron correos electrónicos.");
    }

			
			index = "1";
		}

		public void ObtenerValoresPhone()
		{
			int index1 = 1;
			string Phone = "";
			string ObtenerPhone = "";
			
			index = index1.ToString();
			
			while(repo.ContAndFail_Chrome.ObtenerPhoneNewInfo.Exists()){
				
				ObtenerPhone = repo.ContAndFail_Chrome.ObtenerPhoneNew.Element.GetAttributeValue("caption").ToString();				
				if(!Phone.Contains(ObtenerPhone)){
					 Report.Info("Estoy en el if validando ObtenerPhone: " + ObtenerPhone + ". No está en la lista de Phone: " + Phone);
					Phone = Phone + ObtenerPhone + " , ";
					
				}
				
				index1++;
				index = index1.ToString();
			}
			
			if(Phone != ""){
				
				Report.Info("Phone: " + Phone);
				OpenExcel(Phone, "H");
				
			}else{
				Report.Info("No posee ningun Phone");
			}
			
			index = "1";
		}
		
		private void OpenExcel(string dato, string columna){
			
			Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
			
			
			//Abro el archivo excel
			excelApp.Workbooks.Open(@"C:\TEMP\OutputEmpresas.xlsx").Activate();
			excelApp.Visible=false;
			
			
			((Microsoft.Office.Interop.Excel.Worksheet)excelApp.ActiveWorkbook.Sheets["DatosFiltrados"]).Select();
			
			int fila=1;
			//string ultimaFilaEscrita=excelApp.Range["A"+fila].Value;
			
			//Tomo la ultima fila escrita en el documento Excel
			while(excelApp.Range["F"+fila].Value.ToString() != Linkedin){
				
				fila=fila+1;
			}
			
			//Escribo el dato
			excelApp.Range[columna+fila].Value = dato;
			
			excelApp.ActiveWorkbook.Save();
			
			excelApp.ActiveWorkbook.Close();
			
			Report.Info("Excel guardado y cerrado correctamente");
			
		}
	}
}
