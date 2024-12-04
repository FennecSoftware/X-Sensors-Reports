using System;
using System.Runtime.InteropServices;
using System.Windows.Forms; // For the MessageBox class

namespace Lppx2
{
	/// <summary>
	/// Wrapper class to manage the connection with CODESOFT ActiveX interface
	/// </summary>
	public class Lppx2Manager
	{
		private LabelManager2.Application m_CsApp;
		private bool bDeleteCsApp = false;
		private bool bUnableToLoad = false;
		private const string Lppx2ProgId = "Lppx2.Application";
		
		public Lppx2Manager()
		{
			
		}

		~Lppx2Manager()
		{
		}

		private void ConnectToLppx2()
		{
			if(m_CsApp != null)
				return;

			Object csObject;

			try
			{
				csObject = System.Runtime.InteropServices.Marshal.GetActiveObject(Lppx2ProgId);
			}
			catch
			{
				//No CODESOFT object Running !
				csObject = null;
			}

			try
			{
				if(csObject == null)
				{
					m_CsApp = new LabelManager2.Application();
					bDeleteCsApp = true;
				}
				else
				{
					m_CsApp = (LabelManager2.Application)csObject;
				}
				//m_CsApp.Visible = true;
			}
			catch(Exception e)
			{
				string szerror = e.Message.ToString();
				MessageBox.Show(szerror);
			}
				
			if(bUnableToLoad == false)
			{
				if( m_CsApp.IsEval )
				{
					MessageBox.Show(null,
						"DemoMode",
						"Quick Print",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show(null,
					"Unable to load CODESOFT",
					"Quick Print",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error);
			}
		}

		public void QuitLppx2()
		{
			if(m_CsApp != null)
			{
				if(bDeleteCsApp)
				{
                    var documents = m_CsApp.Documents;
					documents.CloseAll(false);
				    Marshal.ReleaseComObject(documents);

					m_CsApp.Quit();

                    bDeleteCsApp = false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(m_CsApp);
                m_CsApp = null;
            }
		}

		public LabelManager2.Application GetApplication()
		{
			ConnectToLppx2();
			return m_CsApp;
		}

		public LabelManager2.Document GetActiveDocument()
		{
			ConnectToLppx2();
			return m_CsApp.ActiveDocument;
		}

		public void SwitchPrinter(string PrinterName)
		{
			ConnectToLppx2();

			LabelManager2.Document ActiveDoc = GetActiveDocument();
			if(ActiveDoc != null)
			{
				LabelManager2.PrinterSystem PrnSystem = m_CsApp.PrinterSystem();
				LabelManager2.Strings PrintersName = PrnSystem.Printers(LabelManager2.enumKindOfPrinters.lppxAllPrinters);
				
				string CurrentPrinter = ActiveDoc.Printer.Name;
				if(CurrentPrinter != PrinterName)
				{
					short NbPrinters = PrintersName.Count;
					string FullPrinterName = "";
					for(short i=1;i<=NbPrinters;i++)
					{
						FullPrinterName = PrintersName.Item(i);
						int pos = FullPrinterName.LastIndexOf(',');
						if(pos != -1)
						{
							string CurrentPrinterName = FullPrinterName.Substring(0,pos);
							if(CurrentPrinterName == PrinterName)
							{
								bool bDirectAccess = false;
								string PortName = FullPrinterName.Substring(pos+1);
								if(PortName.StartsWith("->"))
								{
									PortName = PortName.Substring(2);
									bDirectAccess = true;
								}
                                var printer = ActiveDoc.Printer;
								printer.SwitchTo(PrinterName, PortName, bDirectAccess);
							    Marshal.ReleaseComObject(printer);
							}
						}
					}
					System.Runtime.InteropServices.Marshal.ReleaseComObject(PrintersName);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(PrnSystem);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ActiveDoc);
				}
			}
		}
	}
}
