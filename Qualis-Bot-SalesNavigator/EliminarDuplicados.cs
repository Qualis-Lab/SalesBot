/*
 * Created by Ranorex
 * User: Nicolas Risso
 * Date: 1/5/2023
 * Time: 11:11 AM
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
    /// Description of EliminarDuplicados.
    /// </summary>
    [TestModule("218D891F-26DB-49D2-BC19-80A202D49D44", ModuleType.UserCode, 1)]
    public class EliminarDuplicados : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public EliminarDuplicados()
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
            
            XlsWorker.EliminarDuplicados("", "DatosObtenidos");
            XlsWorker.EliminarDuplicados("", "DatosFiltrados");
          
            // Abrir una instancia de Excel y obtener una referencia a la primera hoja de cálculo
            /*var excelApp = new Application();
            var workbook = excelApp.Workbooks.Open("C:\\Example.xlsx");
            var worksheet = workbook.Sheets[1];
            // Seleccionar la columna A y obtener el rango de celdas
            var columnRange = worksheet.Range["A:A"];
            // Eliminar las filas duplicadas en el rango de celdas
            columnRange.RemoveDuplicates(XlYesNoGuess.xlYes, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Guardar el libro y cerrar Excel
            workbook.Save();
            excelApp.Quit();*/
        
        }
    }
}
