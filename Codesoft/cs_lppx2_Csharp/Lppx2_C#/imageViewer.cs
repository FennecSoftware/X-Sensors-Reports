using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Lppx2
{
	/// <summary>
	/// Description résumée de imageViewer.
	/// </summary>
	public class imageViewer : System.Windows.Forms.Form
	{
		private System.Windows.Forms.PictureBox bigImage;
		private System.Windows.Forms.HScrollBar hScrollBar1;
		private System.Windows.Forms.VScrollBar vScrollBar1;
		public Form opener ;
		private System.Windows.Forms.Button b_close;
		

		public Image imageToDisplay ;

		public imageViewer()
		{
			//
			// Requis pour la prise en charge du Concepteur Windows Forms
			//
			InitializeComponent();

			//
			// TODO : ajoutez le code du constructeur après l'appel à InitializeComponent
			//
		}

		/// <summary>
		/// Nettoyage des ressources utilisées.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			
			base.Dispose( disposing );
		}

		#region Code généré par le Concepteur Windows Form
		/// <summary>
		/// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
		/// le contenu de cette méthode avec l'éditeur de code.
		/// </summary>
		private void InitializeComponent()
		{
			this.bigImage = new System.Windows.Forms.PictureBox();
			this.hScrollBar1 = new System.Windows.Forms.HScrollBar();
			this.vScrollBar1 = new System.Windows.Forms.VScrollBar();
			this.b_close = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// bigImage
			// 
			this.bigImage.AccessibleRole = System.Windows.Forms.AccessibleRole.ScrollBar;
			this.bigImage.Location = new System.Drawing.Point(0, 0);
			this.bigImage.Name = "bigImage";
			this.bigImage.Size = new System.Drawing.Size(568, 576);
			this.bigImage.TabIndex = 0;
			this.bigImage.TabStop = false;
			this.bigImage.MouseDown += new System.Windows.Forms.MouseEventHandler(this.bigImage_MouseDown);
			// 
			// hScrollBar1
			// 
			this.hScrollBar1.Location = new System.Drawing.Point(0, 576);
			this.hScrollBar1.Name = "hScrollBar1";
			this.hScrollBar1.Size = new System.Drawing.Size(568, 16);
			this.hScrollBar1.TabIndex = 1;
			this.hScrollBar1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.ScrollBar1_Scroll);
			// 
			// vScrollBar1
			// 
			this.vScrollBar1.Location = new System.Drawing.Point(568, 0);
			this.vScrollBar1.Name = "vScrollBar1";
			this.vScrollBar1.Size = new System.Drawing.Size(16, 576);
			this.vScrollBar1.TabIndex = 2;
			this.vScrollBar1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.ScrollBar1_Scroll);
			// 
			// b_close
			// 
			this.b_close.Location = new System.Drawing.Point(120, 280);
			this.b_close.Name = "b_close";
			this.b_close.Size = new System.Drawing.Size(88, 24);
			this.b_close.TabIndex = 4;
			this.b_close.Text = "Close Preview";
			this.b_close.Click += new System.EventHandler(this.b_close_Click);
			// 
			// imageViewer
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(584, 592);
			this.Controls.Add(this.b_close);
			this.Controls.Add(this.vScrollBar1);
			this.Controls.Add(this.hScrollBar1);
			this.Controls.Add(this.bigImage);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "imageViewer";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "imageViewer";
			this.Load += new System.EventHandler(this.imageViewer_Load);
			this.Closed += new System.EventHandler(this.imageViewer_Closed);
			this.ResumeLayout(false);

		}
		#endregion

		private void imageViewer_Load(object sender, System.EventArgs e)
		{
			// Dealing with control placement
			this.Left = 0 ;
			this.Top = 0 ;
			this.Width = Screen.GetBounds(this).Width - 5 ;
			this.Height = Screen.GetBounds(this).Height - 80 ;
			this.bigImage.Width = this.Width - this.vScrollBar1.Width ;
			this.bigImage.Height = this.Height - this.hScrollBar1.Height ;
			this.vScrollBar1.Left = this.bigImage.Width ;
			this.hScrollBar1.Top = this.bigImage.Height ;
			this.vScrollBar1.Height = this.bigImage.Height;
			this.hScrollBar1.Width = this.bigImage.Width;
			bool hScrollDisabled = false ;
			bool vScrollDisabled = false ;

			if (this.imageToDisplay.Width < this.bigImage.Width)
			{
				this.bigImage.Width = this.imageToDisplay.Width ;
				this.hScrollBar1.Width = this.bigImage.Width ;
				this.vScrollBar1.Left = this.bigImage.Width ;
				this.Width = this.bigImage.Width + this.vScrollBar1.Width ;
				hScrollDisabled = true ;
			}
			if (this.imageToDisplay.Height < this.bigImage.Height)
			{
				this.bigImage.Height = this.imageToDisplay.Height ;
				this.vScrollBar1.Height = this.bigImage.Height ;
				this.hScrollBar1.Top = this.bigImage.Height ;
				this.Height = this.bigImage.Height + this.hScrollBar1.Height ;
				vScrollDisabled = true ;
			}
			bigImage.Image = this.imageToDisplay ;
			this.hScrollBar1.Minimum = 0 ;
			this.vScrollBar1.Minimum = 0 ;
			if (hScrollDisabled)
			{
				this.hScrollBar1.Enabled = false ;
			}
			else
			{
				this.hScrollBar1.Maximum = this.imageToDisplay.Width - this.bigImage.Width ;
			}
			if (vScrollDisabled)
			{
				this.vScrollBar1.Enabled = false ;
			}
			else
			{
				this.vScrollBar1.Maximum = this.imageToDisplay.Height - this.bigImage.Height ;
			}
			this.b_close.Top = this.Height + 2 ;
			this.Height += 46 ;
			this.Width += 5 ;
			this.b_close.Left = (int)((this.Width - this.b_close.Width)/2);
			this.Left = (int)((Screen.GetBounds(this).Width - this.Width)/2);
			this.Top = (int)((Screen.GetBounds(this).Height - this.Height)/2) ;

		}

		private void ScrollBar1_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e)
		{
			Graphics g = bigImage.CreateGraphics();
			Rectangle dest = new Rectangle (new System.Drawing.Point (0,0),new System.Drawing.Size (bigImage.Width,bigImage.Height)) ;
			Rectangle src = new Rectangle (new System.Drawing.Point (this.hScrollBar1.Value,this.vScrollBar1.Value),new System.Drawing.Size (bigImage.Width,bigImage.Height)) ;
			g.DrawImage (this.imageToDisplay,dest,src, GraphicsUnit.Pixel);
			bigImage.Update () ;
		}


		private void imageViewer_Closed(object sender, System.EventArgs e)
		{
			opener.Show() ;
		}

		private void bigImage_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
				this.Close () ;
		}

		private void b_close_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}




	}
}
