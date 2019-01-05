namespace WordAddIn1
{
	partial class UserControl1
	{
		/// <summary> 
		/// Variável de designer necessária.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary> 
		/// Limpar os recursos que estão sendo usados.
		/// </summary>
		/// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Código gerado pelo Designer de Componentes

		/// <summary> 
		/// Método necessário para suporte ao Designer - não modifique 
		/// o conteúdo deste método com o editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.button2 = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.Lista_Teste = new System.Windows.Forms.ComboBox();
			this.button3 = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// button2
			// 
			this.button2.BackColor = System.Drawing.SystemColors.Control;
			this.button2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.button2.Location = new System.Drawing.Point(37, 74);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(120, 23);
			this.button2.TabIndex = 6;
			this.button2.Text = "Condicionar Radio";
			this.button2.UseVisualStyleBackColor = false;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// button1
			// 
			this.button1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.button1.Location = new System.Drawing.Point(37, 29);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(120, 23);
			this.button1.TabIndex = 5;
			this.button1.Text = "Condicionar Check";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// Lista_Teste
			// 
			this.Lista_Teste.AccessibleRole = System.Windows.Forms.AccessibleRole.Document;
			this.Lista_Teste.AllowDrop = true;
			this.Lista_Teste.BackColor = System.Drawing.SystemColors.ActiveCaption;
			this.Lista_Teste.Cursor = System.Windows.Forms.Cursors.Default;
			this.Lista_Teste.FormattingEnabled = true;
			this.Lista_Teste.Items.AddRange(new object[] {
            "a",
            "b",
            "c",
            "d"});
			this.Lista_Teste.Location = new System.Drawing.Point(36, 167);
			this.Lista_Teste.Name = "Lista_Teste";
			this.Lista_Teste.Size = new System.Drawing.Size(121, 21);
			this.Lista_Teste.TabIndex = 4;
			this.Lista_Teste.Click += new System.EventHandler(this.Lista_Teste_Click);
			this.Lista_Teste.MouseMove += new System.Windows.Forms.MouseEventHandler(this.Lista_Teste_MouseMove);
			// 
			// button3
			// 
			this.button3.Location = new System.Drawing.Point(36, 120);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(121, 23);
			this.button3.TabIndex = 7;
			this.button3.Text = "Colchete Antes";
			this.button3.UseVisualStyleBackColor = true;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// UserControl1
			// 
			this.AllowDrop = true;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.SystemColors.Desktop;
			this.Controls.Add(this.button3);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.Lista_Teste);
			this.Name = "UserControl1";
			this.Size = new System.Drawing.Size(252, 408);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.ComboBox Lista_Teste;
		private System.Windows.Forms.Button button3;
	}
}
