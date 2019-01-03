using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Windows.Forms.ComponentModel;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;


namespace WordAddIn1
{
	public partial class Ribbon2
	{
		private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
		{
			//thisAddin = new ThisAddIn();
		}

		public void button1_Click(object sender, RibbonControlEventArgs e)
		{

			ThisAddIn.iniciarPlugin();

		}
	}
}
