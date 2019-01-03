using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
	public partial class ThisAddIn
	{
		private UserControl1 UserControl1;
		private static Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;


		public void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			UserControl1 = new UserControl1();
			myCustomTaskPane = this.CustomTaskPanes.Add(UserControl1, "My Plugin");
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
		}

		public static void iniciarPlugin()
		{
			myCustomTaskPane.Visible = !myCustomTaskPane.Visible;
		}

		#region Código gerado por VSTO

		/// <summary>
		/// Método necessário para suporte ao Designer - não modifique 
		/// o conteúdo deste método com o editor de código.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
