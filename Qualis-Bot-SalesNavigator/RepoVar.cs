/*
 * Created by Ranorex
 * User: Nicolas Risso
 * Date: 1/4/2023
 * Time: 5:28 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Repository;
using Ranorex.Core.Testing;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Qualis_Bot_SalesNavigator
{
	/// <summary>
	/// Description of RepoVar.
	/// </summary>
	public static class RepoVar
	{
		private static Microsoft.Office.Interop.Excel.Application exelApp;
		
		public static Application getExelInstance(){
			return exelApp;
		}
		public static void setExelInstance(Microsoft.Office.Interop.Excel.Application instanExel){
			exelApp = instanExel;
		}
	}
}
