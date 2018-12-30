using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Ribbon;


namespace WordAddIn1
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();

		}

		private void Form1_Activated(object sender, EventArgs e)
		{
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Word.Selection selecao = Globals.ThisAddIn.Application.Selection;


			//	MessageBox.Show(selecao.Text);
			comboBox1.Items.Add(selecao.Text);

			//Word.Document doc = Globals.ThisAddIn.Application.Documents.Add();



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
		/*
				private void comboBox1_DragEnter(object sender, DragEventArgs e)
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
