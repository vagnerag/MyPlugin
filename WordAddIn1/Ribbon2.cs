using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace WordAddIn1
{
	public partial class Ribbon2
	{
		private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{
			Form1 fm = new Form1();
			fm.Show();
			
		}
	}
}
