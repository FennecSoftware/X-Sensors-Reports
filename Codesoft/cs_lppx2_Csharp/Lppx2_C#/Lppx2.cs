using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Microsoft.Win32;


namespace Lppx2
{
	/// <summary>
	/// Description résumée de Form1.
	/// </summary>
	/// 
    public class Form1 : System.Windows.Forms.Form
    {
        #region Members
        private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox cmbFolder;
		private System.Windows.Forms.Button btnAddFolder;
		private System.Windows.Forms.ListBox listboxLabels;
		private System.Windows.Forms.ComboBox cmbFileExtension;
		private System.String [][] varTab = new string [2][] ;
		private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox listboxVars;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnUpdateVar;
		private System.Windows.Forms.TextBox lblVarValue;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.TextBox txtQtyValue;
		private System.Windows.Forms.Button btnPageSetup;
		private System.Windows.Forms.Button btnPrint;
		private System.String csPath ;
		private System.Int32 picDefW ;
		private System.Int32 picDefH ;
		private Image currentImage ;
		public Image realSizeImage ;


		/// <summary>
		/// Variable nécessaire au concepteur.
		/// </summary>
		private System.ComponentModel.Container components = null;

		private Lppx2Manager m_Lppx2Manager = null;
		private System.Windows.Forms.PictureBox docPreview;
		private LabelManager2.Application CsApp = null;
	    private LabelManager2.Document ActiveDocument = null;
		private System.Windows.Forms.Label l_loading;
		private System.Windows.Forms.Button b_realSize;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.CheckBox chkEvents;
		private System.Windows.Forms.ListBox listBoxEvents;
		private Image.GetThumbnailImageAbort myCallback ;

        // Document Closed Delegate
        LabelManager2.INotifyApplicationEvent_DocumentClosedEventHandler _DocumentClosedDelegate;

        // Printing delegates
        LabelManager2.INotifyDocumentEvent_BeginPrintingEventHandler _BeginPrintingDelegate;
        LabelManager2.INotifyDocumentEvent_PausedPrintingEventHandler _PausedPrintingDelegate;
        LabelManager2.INotifyDocumentEvent_EndPrintingEventHandler _EndPrintingDelegate;

        // Used by Invoke call to update the UI from the Events thread
        // Necessary to avoid an inter-thread invalid operation exception
        private delegate void UpdateMessagesListEvent(string strMessage);
        private UpdateMessagesListEvent _UpdateMessagesList;
        
        // Need to use BeginInvoke instead of Invoke for Printing events
        // or the Printing thread is blocked by Invoke (synchronous call)
        private IAsyncResult _BeginPrintingEventRes;
        private IAsyncResult _PausedPrintingEventRes;
        private IAsyncResult _EndPrintingEventRes;
        private Label label3;
        #endregion

        public Form1()
		{
			//
			// Requis pour la prise en charge du Concepteur Windows Forms
			//
			InitializeComponent();

			// Create an instance of Lppx2Manager
			m_Lppx2Manager = new Lppx2Manager();
			CsApp = m_Lppx2Manager.GetApplication();

            l_loading.Text = "Please Wait.\nLoading in progress... " ;
            CsApp.PreloadUI();

            // Document Closed Delegate
            _DocumentClosedDelegate = new LabelManager2.INotifyApplicationEvent_DocumentClosedEventHandler(CsApp_DocumentClosed);

            // Printing Delegates
            _BeginPrintingDelegate = new LabelManager2.INotifyDocumentEvent_BeginPrintingEventHandler(ActiveDocument_BeginPrinting);
            _PausedPrintingDelegate = new LabelManager2.INotifyDocumentEvent_PausedPrintingEventHandler(ActiveDocument_PausedPrinting);
            _EndPrintingDelegate = new LabelManager2.INotifyDocumentEvent_EndPrintingEventHandler(ActiveDocument_EndPrinting);

            // UI Delegate
            _UpdateMessagesList = new UpdateMessagesListEvent(UpdateMessagesList);

            _BeginPrintingEventRes = null;
            _PausedPrintingEventRes = null;
            _EndPrintingEventRes = null;

            // Use Events by default
            chkEvents.CheckState = CheckState.Checked;
        }


		/// <summary>
		/// Nettoyage des ressources utilisées.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}

			base.Dispose( disposing );
		}

		#region Code généré par le Concepteur Windows Form
		/// <summary>
		/// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
		/// le contenu de cette méthode avec l'éditeur de code.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.cmbFolder = new System.Windows.Forms.ComboBox();
            this.btnAddFolder = new System.Windows.Forms.Button();
            this.listboxLabels = new System.Windows.Forms.ListBox();
            this.cmbFileExtension = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtQtyValue = new System.Windows.Forms.TextBox();
            this.listboxVars = new System.Windows.Forms.ListBox();
            this.lblVarValue = new System.Windows.Forms.TextBox();
            this.btnUpdateVar = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.b_realSize = new System.Windows.Forms.Button();
            this.l_loading = new System.Windows.Forms.Label();
            this.docPreview = new System.Windows.Forms.PictureBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnPrint = new System.Windows.Forms.Button();
            this.btnPageSetup = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.listBoxEvents = new System.Windows.Forms.ListBox();
            this.chkEvents = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.docPreview)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label1.Location = new System.Drawing.Point(7, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "Label Folder";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbFolder
            // 
            this.cmbFolder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFolder.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.cmbFolder.Location = new System.Drawing.Point(96, 34);
            this.cmbFolder.Name = "cmbFolder";
            this.cmbFolder.Size = new System.Drawing.Size(270, 21);
            this.cmbFolder.Sorted = true;
            this.cmbFolder.TabIndex = 1;
            this.cmbFolder.SelectedIndexChanged += new System.EventHandler(this.OnFolderChange);
            // 
            // btnAddFolder
            // 
            this.btnAddFolder.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnAddFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAddFolder.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.btnAddFolder.Location = new System.Drawing.Point(364, 19);
            this.btnAddFolder.Name = "btnAddFolder";
            this.btnAddFolder.Size = new System.Drawing.Size(84, 24);
            this.btnAddFolder.TabIndex = 2;
            this.btnAddFolder.Text = "Add Folder";
            this.btnAddFolder.UseVisualStyleBackColor = false;
            this.btnAddFolder.Click += new System.EventHandler(this.OnBtnAddFolder);
            // 
            // listboxLabels
            // 
            this.listboxLabels.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.listboxLabels.Location = new System.Drawing.Point(10, 84);
            this.listboxLabels.Name = "listboxLabels";
            this.listboxLabels.Size = new System.Drawing.Size(216, 160);
            this.listboxLabels.Sorted = true;
            this.listboxLabels.TabIndex = 3;
            this.listboxLabels.SelectedIndexChanged += new System.EventHandler(this.OnSelectLabel);
            // 
            // cmbFileExtension
            // 
            this.cmbFileExtension.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.cmbFileExtension.Items.AddRange(new object[] {
            "*.lab",
            "*.*"});
            this.cmbFileExtension.Location = new System.Drawing.Point(88, 49);
            this.cmbFileExtension.Name = "cmbFileExtension";
            this.cmbFileExtension.Size = new System.Drawing.Size(270, 21);
            this.cmbFileExtension.TabIndex = 4;
            this.cmbFileExtension.SelectedIndexChanged += new System.EventHandler(this.OnExtensionChange);
            this.cmbFileExtension.KeyUp += new System.Windows.Forms.KeyEventHandler(this.OnExtensionChange);
            // 
            // label2
            // 
            this.label2.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label2.Location = new System.Drawing.Point(60, 30);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 16);
            this.label2.TabIndex = 7;
            this.label2.Text = "Label qty :";
            // 
            // txtQtyValue
            // 
            this.txtQtyValue.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.txtQtyValue.Location = new System.Drawing.Point(125, 26);
            this.txtQtyValue.Name = "txtQtyValue";
            this.txtQtyValue.Size = new System.Drawing.Size(32, 20);
            this.txtQtyValue.TabIndex = 8;
            this.txtQtyValue.Text = "1";
            // 
            // listboxVars
            // 
            this.listboxVars.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.listboxVars.Location = new System.Drawing.Point(10, 19);
            this.listboxVars.Name = "listboxVars";
            this.listboxVars.Size = new System.Drawing.Size(216, 95);
            this.listboxVars.TabIndex = 9;
            this.listboxVars.SelectedIndexChanged += new System.EventHandler(this.DisplaySelectedVarValue);
            // 
            // lblVarValue
            // 
            this.lblVarValue.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.lblVarValue.Location = new System.Drawing.Point(10, 130);
            this.lblVarValue.Name = "lblVarValue";
            this.lblVarValue.Size = new System.Drawing.Size(150, 20);
            this.lblVarValue.TabIndex = 12;
            // 
            // btnUpdateVar
            // 
            this.btnUpdateVar.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnUpdateVar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnUpdateVar.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.btnUpdateVar.Location = new System.Drawing.Point(166, 126);
            this.btnUpdateVar.Name = "btnUpdateVar";
            this.btnUpdateVar.Size = new System.Drawing.Size(60, 24);
            this.btnUpdateVar.TabIndex = 13;
            this.btnUpdateVar.Text = "Update";
            this.btnUpdateVar.UseVisualStyleBackColor = true;
            this.btnUpdateVar.Click += new System.EventHandler(this.OnBtnUpdateVariable);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.b_realSize);
            this.groupBox1.Controls.Add(this.l_loading);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnAddFolder);
            this.groupBox1.Controls.Add(this.listboxLabels);
            this.groupBox1.Controls.Add(this.cmbFileExtension);
            this.groupBox1.Controls.Add(this.docPreview);
            this.groupBox1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.groupBox1.Location = new System.Drawing.Point(8, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(457, 256);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "File preview";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label3.Location = new System.Drawing.Point(10, 52);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "File Extension";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // b_realSize
            // 
            this.b_realSize.BackColor = System.Drawing.Color.LightSteelBlue;
            this.b_realSize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.b_realSize.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.b_realSize.Location = new System.Drawing.Point(232, 220);
            this.b_realSize.Name = "b_realSize";
            this.b_realSize.Size = new System.Drawing.Size(216, 24);
            this.b_realSize.TabIndex = 2;
            this.b_realSize.Text = "Real Size";
            this.b_realSize.UseVisualStyleBackColor = true;
            this.b_realSize.Visible = false;
            this.b_realSize.Click += new System.EventHandler(this.OnBtnDisplayRealSizePreview);
            // 
            // l_loading
            // 
            this.l_loading.Location = new System.Drawing.Point(250, 105);
            this.l_loading.Name = "l_loading";
            this.l_loading.Size = new System.Drawing.Size(183, 91);
            this.l_loading.TabIndex = 1;
            this.l_loading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // docPreview
            // 
            this.docPreview.Location = new System.Drawing.Point(232, 84);
            this.docPreview.Name = "docPreview";
            this.docPreview.Size = new System.Drawing.Size(216, 130);
            this.docPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.docPreview.TabIndex = 0;
            this.docPreview.TabStop = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.listboxVars);
            this.groupBox2.Controls.Add(this.lblVarValue);
            this.groupBox2.Controls.Add(this.btnUpdateVar);
            this.groupBox2.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.groupBox2.Location = new System.Drawing.Point(8, 274);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(234, 159);
            this.groupBox2.TabIndex = 15;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Variables settings";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnPrint);
            this.groupBox3.Controls.Add(this.btnPageSetup);
            this.groupBox3.Controls.Add(this.txtQtyValue);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.groupBox3.Location = new System.Drawing.Point(248, 274);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(217, 159);
            this.groupBox3.TabIndex = 16;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Printing settings";
            // 
            // btnPrint
            // 
            this.btnPrint.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPrint.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.btnPrint.Location = new System.Drawing.Point(63, 52);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(94, 24);
            this.btnPrint.TabIndex = 10;
            this.btnPrint.Text = "Print";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.OnBtnPrintDocument);
            // 
            // btnPageSetup
            // 
            this.btnPageSetup.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnPageSetup.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPageSetup.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.btnPageSetup.Location = new System.Drawing.Point(63, 126);
            this.btnPageSetup.Name = "btnPageSetup";
            this.btnPageSetup.Size = new System.Drawing.Size(94, 24);
            this.btnPageSetup.TabIndex = 9;
            this.btnPageSetup.Text = "Page setup";
            this.btnPageSetup.UseVisualStyleBackColor = true;
            this.btnPageSetup.Click += new System.EventHandler(this.OnBtnPageSetup);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.listBoxEvents);
            this.groupBox4.Controls.Add(this.chkEvents);
            this.groupBox4.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.groupBox4.Location = new System.Drawing.Point(8, 439);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(457, 144);
            this.groupBox4.TabIndex = 19;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Events";
            // 
            // listBoxEvents
            // 
            this.listBoxEvents.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.listBoxEvents.Location = new System.Drawing.Point(8, 40);
            this.listBoxEvents.Name = "listBoxEvents";
            this.listBoxEvents.Size = new System.Drawing.Size(440, 95);
            this.listBoxEvents.TabIndex = 20;
            this.listBoxEvents.DoubleClick += new System.EventHandler(this.listBoxEvents_DoubleClick);
            // 
            // chkEvents
            // 
            this.chkEvents.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.chkEvents.Location = new System.Drawing.Point(8, 16);
            this.chkEvents.Name = "chkEvents";
            this.chkEvents.Size = new System.Drawing.Size(168, 24);
            this.chkEvents.TabIndex = 19;
            this.chkEvents.Text = "Catch events";
            this.chkEvents.CheckedChanged += new System.EventHandler(this.chkEvents_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.BackColor = System.Drawing.Color.Silver;
            this.ClientSize = new System.Drawing.Size(473, 592);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.cmbFolder);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "C#_Lppx2.Application";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.OnExit);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.docPreview)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Point d'entrée principal de l'application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}
        

        #region Windows Form Events
        private void Form1_Load(object sender, System.EventArgs e)
		{
			// Reading codeSoft directory from register
			csPath = CsApp.DefaultFilePath ;
			cmbFolder.Items.Add (csPath) ;

			// Displaying variables list for the current Doc
			cmbFileExtension.SelectedIndex = 0 ;
			cmbFolder.SelectedIndex = 0 ;
			picDefW = docPreview.Width ;
			picDefH = docPreview.Height ;
			myCallback = new Image.GetThumbnailImageAbort(ThumbnailCallback);

            // Set the Form Caption
            this.Text = CsApp.ActivePrinterName;
            
            // Makes the application visible
            //CsApp.Visible = true;
		}

        private void OnExit(object sender, FormClosingEventArgs e)
        {
            // Remove Events Handlers
            chkEvents.CheckState = CheckState.Unchecked;
            DisposeImages();
            if (ActiveDocument != null)
            {
                try
                {
                    ActiveDocument.Close(false);
                }
                finally 
                {
                    Marshal.ReleaseComObject(ActiveDocument);
                }
            }
            m_Lppx2Manager.QuitLppx2();
        }
        #endregion

        #region Files List Update

        private void UpdateFilesList()
		{
			listboxLabels.Items.Clear () ;
			if ((cmbFolder.Text != "") && (cmbFileExtension.Text != ""))
			{
				string[] fileList = Directory.GetFiles(cmbFolder.Text,cmbFileExtension.Text) ;
				foreach (string currentFile in fileList) 
				{
					listboxLabels.Items.Add(currentFile.Substring (currentFile.LastIndexOf("\\")+1));
				}
			}
		}

        private void OnBtnAddFolder(object sender, System.EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.SelectedPath = csPath;

            if (FBD.ShowDialog() == DialogResult.OK)
            {
                string dirName = FBD.SelectedPath;
                cmbFolder.SelectedIndex = cmbFolder.Items.Add(dirName);
            }
        }

		private void OnFolderChange(object sender, System.EventArgs e)
		{
			UpdateFilesList () ;
		}

		private void OnExtensionChange(object sender, System.EventArgs e)
		{
			UpdateFilesList () ;
		}

		private void OnExtensionChange(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			UpdateFilesList () ;
		}

		#endregion

		#region Preview Management

		public bool ThumbnailCallback()
		{
			return false;
		}

		private void resizeIfNeeded ()
		{
			// If image width is greater than the pictureBox width 
			// OR
			// image height is greater than the pictureBox height
			// then it must be resized...
			if (currentImage != null)
			{
				if ((currentImage.Width > picDefW) || (currentImage.Height > picDefH))
				{
					// which one (width or height) must be used to resize ?
					if ((currentImage.Width / picDefW) > (currentImage.Height / picDefH))
					{
						// Must use width to resize
						double a = ((double)picDefW) / (((double)currentImage.Width)) ;
						currentImage = currentImage.GetThumbnailImage(picDefW, (int)((double)currentImage.Height*a), myCallback, IntPtr.Zero);
					}
					else
					{
						// Must use Height to resize
						double a = ((double)picDefH) / (((double)currentImage.Height)) ;
						currentImage = currentImage.GetThumbnailImage( (int)((double)currentImage.Width*a),picDefH, myCallback, IntPtr.Zero);

						//	currentImage = currentImage.GetThumbnailImage((int)(currentImage.Width/(currentImage.Height / picDefH)),picDefH, myCallback, IntPtr.Zero);
					}
				}
			}
		}

        private void DisposeImages()
        {
            if (realSizeImage != null)
            {
                realSizeImage.Dispose();
            }

            if (currentImage != null)
            {
                currentImage.Dispose();
            }

            if (docPreview.Image != null)
            {
                docPreview.Image.Dispose();
            }
        }
		
		#endregion

		#region Label Variables Management

		// Fill Variables Array with variable name and value
		private void FillVariablesArray ()
		{
            if(ActiveDocument == null)
                return;

			listboxVars.Items.Clear () ;
		    var docVariables = ActiveDocument.Variables;

            varTab[0] = new string[docVariables.Count + 1];
            varTab[1] = new string[docVariables.Count + 1];

            for (int i = 1; i <= docVariables.Count; i++)
            {
                var item = docVariables.Item(i);
				varTab[0][i-1] = (item.Name);
				varTab[1][i-1] = (item.Value);
                Marshal.ReleaseComObject(item);
            }

		    Marshal.ReleaseComObject(docVariables);
		}
		
		// Inject variables Name in the listBox
		private void FillVariablesListBox ()
		{
			listboxVars.Items.Clear () ;

			for (int i=0; i < varTab[0].Length-1;i++)
			{
				listboxVars.Items.Add (varTab[0][i]) ;
			}
		}

		// Display selected variable value
		private void DisplaySelectedVarValue(object sender, System.EventArgs e)
		{
			lblVarValue.Text = varTab[1][listboxVars.SelectedIndex] ;
			
		}


		private void OnBtnUpdateVariable(object sender, System.EventArgs e)
		{
            if(ActiveDocument == null)
                return;

			varTab[1][listboxVars.SelectedIndex] = lblVarValue.Text ;

		    var docVariables = ActiveDocument.Variables;

            for (int i = 0; i < varTab[0].Length-1; i++)
            {
                var item = docVariables.Item(i + 1);
				item.Name = varTab[0][i] ;
				item.Value = varTab[1][i] ;
                Marshal.ReleaseComObject(item);
            }

		    Marshal.ReleaseComObject(docVariables);

			DisplayPreview();
		}

		#endregion

        #region Open Selected Label
        private void OnSelectLabel(object sender, System.EventArgs e)
		{
            try
            {
                this.b_realSize.Visible = true;
                docPreview.Visible = false;
                l_loading.Visible = true;

                if (ActiveDocument != null)
                {
                    ActiveDocument.Close(false);
                    Marshal.ReleaseComObject(ActiveDocument);
                    ActiveDocument = null;
                }

                if (CsApp != null && listboxLabels.Text.Length > 0)
                {
                    var documents = CsApp.Documents;
                    ActiveDocument = documents.Open(cmbFolder.Text + "\\" + listboxLabels.Text, false);

                    Marshal.ReleaseComObject(documents);

                    // Set the Form Caption
                    this.Text = CsApp.ActivePrinterName; 
                }

                DisplayPreview();
            }
            catch (Exception ex)
            {
                if (ActiveDocument != null)
                {
                    Marshal.ReleaseComObject(ActiveDocument);
                    ActiveDocument = null;
                }
                System.Diagnostics.Debug.Assert(false, ex.Message);
            }
        }
        #endregion

        #region Display Label Preview
        private void DisplayPreview()
		{
			if (ActiveDocument != null)
			{
				docPreview.Visible = false ;
				l_loading.Visible = true ;
                DisposeImages();

				try
				{
 					//This method use the clipboard.
                    //CsApp.ActiveDocument.CopyToClipboardEx (false,0,256,0,100) ;
					//currentImage = (Image)(Clipboard.GetDataObject  ().GetData (System.Windows.Forms.DataFormats.Bitmap,true)) ;

                    //This method use a stream instead of writing to clipboard or file.
                    object obj = ActiveDocument.GetPreview(true, true, 100);
                    if (obj is System.Array)
                    {
                        byte[] data = (byte[])obj;

                        System.IO.MemoryStream pStream = new System.IO.MemoryStream(data);
                        currentImage = new Bitmap(pStream);
                    }

					realSizeImage = currentImage ;
					resizeIfNeeded () ;
					docPreview.Image = currentImage ;
				}
				catch(Exception)
				{
				}

				docPreview.Visible = true ;
				l_loading.Visible = false ;

				Refresh () ;

				FillVariablesArray () ;
				FillVariablesListBox () ;
				if (listboxVars.Items.Count > 0)
				{
					listboxVars.SelectedIndex = 0 ;
				}
			}
		}

        private void OnBtnDisplayRealSizePreview(object sender, System.EventArgs e)
        {
            imageViewer imViewForm = new imageViewer();
            imViewForm.TopMost = true;
            imViewForm.imageToDisplay = realSizeImage;
            this.Hide();
            imViewForm.opener = this;
            imViewForm.Show();
        }
        #endregion

        #region Page Setup
        private void OnBtnPageSetup(object sender, System.EventArgs e)
		{
			if (ActiveDocument == null)
			{
				MessageBox.Show ("A document must be opened to perform this action...") ;
			}
			else
			{
                var dialogs = CsApp.Dialogs;
                var dialogItem = dialogs.Item(LabelManager2.enumDialogType.lppxPageSetupDialog);
                dialogItem.Show(this.Handle);
			    Marshal.ReleaseComObject(dialogItem);
			    Marshal.ReleaseComObject(dialogs);
			}
        }
        #endregion

        #region Print Document
        private void OnBtnPrintDocument(object sender, System.EventArgs e)
		{
			if (ActiveDocument == null)
			{
				MessageBox.Show ("A document must be opened to print !") ;
				return ;
			}

			try
			{
                // Add Handlers for Printing Events for the Active Document if necessary
                AddPrintingHandlers();

                short errorCode = ActiveDocument.PrintDocument(int.Parse(txtQtyValue.Text));
                if (errorCode <= 0)
                {
                    errorCode = CsApp.GetLastError();
                    MessageBox.Show(CsApp.ErrorMessage(errorCode));
                }

                // Process the asynchronous printing events
                if (_BeginPrintingEventRes != null)
                {
                    listBoxEvents.EndInvoke(_BeginPrintingEventRes);
                    _BeginPrintingEventRes = null;
                }

                if (_PausedPrintingEventRes != null)
                {
                    listBoxEvents.EndInvoke(_PausedPrintingEventRes);
                    _PausedPrintingEventRes = null;
                }

                if (_EndPrintingEventRes != null)
                {
                    listBoxEvents.EndInvoke(_EndPrintingEventRes);
                    _EndPrintingEventRes = null;
                }

                // Remove Handlers for Printing Events for the Active Document if necessary
                RemovePrintingHandlers();
            }
			catch (System.FormatException error)
			{
				MessageBox.Show ("Label qty must be an integer...\n\n\n"+error.Message) ;
			}
        }
        #endregion

        #region Events Management
        private void chkEvents_CheckedChanged(object sender, System.EventArgs e)
		{
            if (chkEvents.CheckState == CheckState.Checked)
            {
                CsApp.EnableEvents = true;
                CsApp.DocumentClosed += _DocumentClosedDelegate;
            }
            else if (chkEvents.CheckState == CheckState.Unchecked)
            {
                CsApp.EnableEvents = false;
                CsApp.DocumentClosed -= _DocumentClosedDelegate;
            }
        }

        private void AddPrintingHandlers()
		{       
            if (ActiveDocument != null)
			{
                if (chkEvents.CheckState == CheckState.Checked)
                {
                    ActiveDocument.BeginPrinting += _BeginPrintingDelegate;
                    ActiveDocument.PausedPrinting += _PausedPrintingDelegate;
                    ActiveDocument.EndPrinting += _EndPrintingDelegate;
                }
            }
        }

        private void RemovePrintingHandlers()
		{     
            if (ActiveDocument != null)
			{
                if (chkEvents.CheckState == CheckState.Checked)
                {
                    ActiveDocument.BeginPrinting -= _BeginPrintingDelegate;
                    ActiveDocument.PausedPrinting -= _PausedPrintingDelegate;
                    ActiveDocument.EndPrinting -= _EndPrintingDelegate;
                }
            }
		}

		private void CsApp_DocumentClosed(string strDocTitle)
		{
			string Message = "The label \"" + strDocTitle + "\" has been closed";
            this.Invoke(this._UpdateMessagesList, new object[] { Message });
		}

        private void ActiveDocument_BeginPrinting(string strDocName)
        {
            string Message = "The label \"" + strDocName + "\" begins to print";

            // Triggers the asynchronous event
            _BeginPrintingEventRes = listBoxEvents.BeginInvoke(this._UpdateMessagesList, new object[] { Message });
        }

        private void ActiveDocument_PausedPrinting(LabelManager2.enumPausedReasonPrinting Reason, ref short Cancel)
        {
            string Message = "Printing has been paused for the following reason: " + Reason.ToString() + ".";

            if ( MessageBox.Show(Message, "PausedPrinting event",MessageBoxButtons.RetryCancel, MessageBoxIcon.Hand) == DialogResult.Cancel )
            {
                Message = "Printing has been cancelled for the following reason: " + Reason.ToString() + ".";
               // Triggers the asynchronous event
                _PausedPrintingEventRes = listBoxEvents.BeginInvoke(this._UpdateMessagesList, new object[] { Message });
                Cancel = -1;
            }
            else
                Cancel = 0;
        }

        private void ActiveDocument_EndPrinting(LabelManager2.enumEndPrinting Reason)
		{
            if(ActiveDocument == null)
                return;

			string Message = "The label \"" + ActiveDocument.Name + "\" has finished printing";
			switch(Reason)
			{
				case LabelManager2.enumEndPrinting.lppxEndOfJob :
					Message += " successfully";
					break;
				case LabelManager2.enumEndPrinting.lppxCancelled :
					Message += " due to cancellation";
					break;
				case LabelManager2.enumEndPrinting.lppxSystemFailure :
					Message += " with system failure";
					break;
			}

            // Triggers the asynchronous event
            _EndPrintingEventRes = listBoxEvents.BeginInvoke(this._UpdateMessagesList, new object[] { Message });
        }

        private void UpdateMessagesList(string strMessage)
        {
            try
            {
                listBoxEvents.Items.Add(strMessage);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.Assert(false,e.Message);
            }
        }
        #endregion

        private void listBoxEvents_DoubleClick(object sender, EventArgs e)
        {
            listBoxEvents.Items.Clear();
        }
    }
}
