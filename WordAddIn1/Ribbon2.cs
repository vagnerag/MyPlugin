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
		public Form1 fm;
		private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
		{
			fm = new Form1();
		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{
				
			//fm.DoDragDrop();
			fm.Show();
			
			
			
		}
	}
}
