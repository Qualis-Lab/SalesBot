using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
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
	/// Creates a Ranorex user code collection. A collection is used to publish user code methods to the user code library.
	/// </summary>
	[UserCodeCollection]
	public class XlsWorker
	{
		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static Application getExelInstance()
		{
			// Verificar si Excel está en ejecución
			var excelApp = GetExcelApplicationObject();
			if (excelApp == null)
			{
				Console.WriteLine("Excel no está en ejecución");
				var excelAppli = new Microsoft.Office.Interop.Excel.Application();
				RepoVar.setExelInstance(excelAppli);
				return excelAppli;
			}
			else
			{
				Console.WriteLine("Excel está en ejecución");
				return excelApp;
			}
		}

		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void EliminarDuplicados(string rango, string nombreHoja)
		{
			// Abrir una instancia de Excel y obtener una referencia a la primera hoja de cálculo
			var excelApp = getExelInstance();
			Workbook wb = excelApp.Workbooks.Open(@"C:\TEMP\OutputEmpresas.xlsx");
			Worksheet ws = (Worksheet)wb.Sheets[nombreHoja];
			// Seleccionar la columna A y obtener el rango de celdas
			Range columnRange = ws.Range["A1:F" + ws.UsedRange.Rows.Count.ToString()];
			//Si necesito mas de una columna es necesario pasar un arreglo
			//object cols = new object[]{4,2};
			columnRange.RemoveDuplicates(4, XlYesNoGuess.xlYes);
			// Guardar el libro y cerrar Excel
			excelApp.ActiveWorkbook.Save();
			Report.Success("Se removieron correctamente los datos duplicados en " + nombreHoja);
		}
		public static void EliminarDato(string nombrePersona, string nombreHoja)
		{
			// Abrir una instancia de Excel y obtener una referencia a la primera hoja de cálculo
			var excelApp = getExelInstance();
			Workbook wb = excelApp.Workbooks.Open(@"C:\TEMP\OutputEmpresas.xlsx");
			Worksheet ws = (Worksheet)wb.Sheets[nombreHoja];

			// Buscar el dato en la hoja de Excel
			Range buscarRango = ws.UsedRange;
			Range resultado = buscarRango.Find(nombrePersona);

			// Verificar si se encontró el dato
			if (resultado != null)
			{
				// Seleccionar toda la fila correspondiente al resultado encontrado
				Range filaARemover = (Range)ws.Rows[resultado.Row];

				// Eliminar la fila
				filaARemover.Delete();

				Report.Success("Se eliminó correctamente la fila correspondiente al dato '" + nombrePersona + "' de la hoja '" + nombreHoja + "'");
			}
			
		}

		/// <summary>
		/// This is a placeholder text. Please describe the purpose of the
		/// user code method here. The method is published to the user code library
		/// within a user code collection.
		/// </summary>
		[UserCodeMethod]
		public static void CerrarExcel()
		{
			Microsoft.Office.Interop.Excel.Application excelApp = RepoVar.getExelInstance();
			excelApp.ActiveWorkbook.Close();
			RepoVar.setExelInstance(null);
			Report.Success("Excel cerrado correctamente");
		}

		private static Application GetExcelApplicationObject()
		{
			try
			{
				// Obtener una referencia a la aplicación Excel en ejecución
				var excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
				return excelApp;
			}
			catch (Exception e)
			{
				// Excel no está en ejecución
				return null;
			}
		}
	}
}
