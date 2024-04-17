/*
 * Created by Ranorex
 * User: Juan Cruz Diana
 * Date: 12/4/2024
 * Time: 18:04
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
    /// Description of ValidacionEmpresaIndustria1.
    /// </summary>
    [TestModule("2ACE9F6D-173F-4C98-A671-B328B3FC698F", ModuleType.UserCode, 1)]
    public class ValidacionEmpresaIndustria : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public ValidacionEmpresaIndustria()
        {
            // Do not delete - a parameterless constructor is required!
        }

string _encontroEmpresa = "";
[TestVariable("d9b1a078-c908-45e9-8fb8-2f5a19576795")]
public string encontroEmpresa
{
	get { return _encontroEmpresa; }
	set { _encontroEmpresa = value; }
}

string _industria = "";
[TestVariable("48077a9d-b51a-4593-ac49-1e8679062432")]
public string industria
{
	get { return _industria; }
	set { _industria = value; }
}

string _nombreEmpresa = "";
[TestVariable("06a280a2-d523-4572-bb21-70befb726500")]
public string nombreEmpresa
{
	get { return _nombreEmpresa; }
	set { _nombreEmpresa = value; }
}

public static Qualis_Bot_SalesNavigatorRepository repo = Qualis_Bot_SalesNavigatorRepository.Instance;



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
        	
            try {
            	Report.Log(ReportLevel.Info, "Validation", "(Optional Action)\r\nValidating AttributeEqual (InnerText=$nombreEmpresa) on item 'InicioSalesNavigator.ValidacionNombreEmpresaIndustria'.", repo.InicioSalesNavigator.ValidacionNombreEmpresaIndustriaInfo, new RecordItemIndex(0));
                Validate.AttributeEqual(repo.InicioSalesNavigator.ValidacionNombreEmpresaIndustriaInfo, "InnerText", nombreEmpresa, null, new Validate.Options(){ReportLevelOnFailure=ReportLevel.Info});
             
                    // Aquí entra si el elemento se encuentra
   				Report.Info($"El nombre de la empresa es: {nombreEmpresa}");
    			encontroEmpresa = "True";
    			Report.Info("El valor del parametro encontroEmpresa es: " + encontroEmpresa);
           

            } catch(Exception ex) {
            	
					//Aquí entra si el elemento no se encuentra
			    Report.Info("La variable empresa no se encuentra.");
			    encontroEmpresa = "False";
			    Report.Log(ReportLevel.Warn, "Module", "(Optional Action) " + ex.Message, new RecordItemIndex(0)); 
			    Report.Info("El valor del parametro encontroEmpresa es: " + encontroEmpresa);}
       		}
        }
    }