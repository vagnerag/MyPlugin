using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;


namespace WordAddIn1
{
	public partial class UserControl1 : UserControl
	{
		public UserControl1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Word.Selection selecao = Globals.ThisAddIn.Application.Selection;
			Word.Range rg = selecao.Range;


			string condicao = "pertence(" + Lista_Teste.SelectedItem.ToString() + ", " + Lista_Teste.Name + ")";

			rg.InsertBefore("[");
			rg.InsertAfter("]");
			rg.Start = rg.Start + 1;
			rg.Select();
			rg.InsertBefore(condicao);
			rg.End = rg.Start + condicao.Length;
			rg.Select();
			rg.Application.Selection.Font.Subscript = - 1;
		}

		private void button2_Click(object sender, EventArgs e)
		{
			Word.Selection selecao = Globals.ThisAddIn.Application.Selection;
			Word.Range rg = selecao.Range;


			string condicao = "ListaTeste = " + Lista_Teste.SelectedItem.ToString();

			rg.InsertBefore("[");
			rg.InsertAfter("]");
			rg.Start = rg.Start + 1;
			rg.Select();
			rg.InsertBefore(condicao);
			rg.End = rg.Start + condicao.Length;
			rg.Select();
			rg.Application.Selection.Font.Subscript = -1;
		}

		protected override void OnDragDrop(DragEventArgs e)
		{
			e.Effect = DragDropEffects.Copy;
			Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
			doc.Save();

			//base.OnDragDrop(e);
		}

		private void comboBox1_DragDrop(object sender, DragEventArgs e)
		{
			OnDragDrop(e);
		}

		/*private void comboBox1_DragEnter(object sender, DragEventArgs e)
		{
			e.Effect = DragDropEffects.All;
		}

		private void comboBox1_DragLeave(object sender, EventArgs e)
		{
			Word.Range rg = Globals.ThisAddIn.Application.Selection.Range;
			MessageBox.Show(e.ToString());
		}*/
	}
}
