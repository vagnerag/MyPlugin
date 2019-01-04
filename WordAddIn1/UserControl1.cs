using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;


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

			string condicao = "pertence(\"" + Lista_Teste.SelectedItem.ToString() + "\", " + Lista_Teste.Name + ")";

			rg.InsertBefore("[");
			rg.InsertAfter("]");
			rg.Select();
			rg.Application.Selection.Font.Color = Word.WdColor.wdColorRed;
			rg.Start = rg.Start + 1;
			rg.Select();
			rg.InsertBefore(condicao);
			rg.End = rg.Start + condicao.Length;
			rg.Select();
			rg.Application.Selection.Font.Subscript = -1;
			//rg = selecao.Range;
			//rg.Application.Selection.Font.Color = Word.WdColor.wdColorRed;
		}

		private void button2_Click(object sender, EventArgs e)
		{
			Word.Selection selecao = Globals.ThisAddIn.Application.Selection;
			Word.Range rg = selecao.Range;

			string condicao = "ListaTeste = \"" + Lista_Teste.SelectedItem.ToString() + "\"";

			rg.InsertBefore("[");
			rg.InsertAfter("]");
			rg.Select();
			rg.Application.Selection.Font.Color = Word.WdColor.wdColorRed;
			rg.Start = rg.Start + 1;
			rg.Select();
			rg.InsertBefore(condicao);
			rg.End = rg.Start + condicao.Length;
			rg.Select();
			rg.Application.Selection.Font.Subscript = -1;
			//rg = selecao.Range;
			//rg.Application.Selection.Font.Color = Word.WdColor.wdColorRed;
		}

		protected override void OnDragDrop(DragEventArgs e)
		{
			e.Effect = DragDropEffects.Copy;
			Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
			doc.Save();

			//base.OnDragDrop(e);
		}

		protected override void OnMouseMove(MouseEventArgs e)
		{
			base.OnMouseMove(e);
			if (e.Button == MouseButtons.Left)
			{
				// Package the data.
				DataObject data = new DataObject();
				data.SetData(DataFormats.StringFormat, Lista_Teste.SelectedItem.ToString());
				//data.SetData("Double", circleUI.Height);
				//data.SetData("Object", this);

				// Inititate the drag-and-drop operation.
				this.Lista_Teste.DoDragDrop(data, DragDropEffects.Copy | DragDropEffects.Move);
			}
		}

		private void Lista_Teste_MouseMove(object sender, MouseEventArgs e)
		{
			OnMouseMove(e);
		}

		private void Lista_Teste_Click(object sender, EventArgs e)
		{
			this.Lista_Teste.DroppedDown = true;
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
