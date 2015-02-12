using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;

namespace CloseWatcherAddin
{
	/// <summary>
	///	Main Add-in class.
	/// </summary>
	[ComVisible(true)]
	[Guid("10356B96-E4D5-4688-9437-61E008522FC5")]
	[ClassInterface(ClassInterfaceType.None)]
	[ProgId("CloseWatcherAddin.Connector")]
	public class Connector : Object, Extensibility.IDTExtensibility2
	{
		// File for document full path
		private StreamWriter _outfile;
		private IDocumentHandler _documentHandler; 

		/// <summary>
		///	Implements the OnConnection method of the IDTExtensibility2 interface.
		///	Receives notification that the Add-in is being loaded.
		/// </summary>
		/// <param name='officeAapplication'>
		///	Root object of the host application.
		/// </param>
		/// <param name='connectMode'>
		///	Describes how the Add-in is being loaded.
		/// </param>
		/// <param name='addInInst'>
		///	Object representing this Add-in.
		/// </param>
		/// <param name='custom'>
		///	Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object officeAapplication, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			try
			{
				// If host app is null, we can't do anything
				if (officeAapplication == null)
					throw new ArgumentNullException("officeAapplication");

				string filePath = Path.Combine(Path.GetTempPath(), "ClosedDocuments.txt");
				_outfile = new StreamWriter(filePath, true);

				// Create handler for specific office app host
				_documentHandler = DocumentHandlerFactory.CreateDocumentHandler(officeAapplication);
				// Connect event
				_documentHandler.DocumentClosed += HandleDocumentClosed;
			}
			catch (Exception e)
			{
				ErrorHandler.ReportError(e);
			}
		}

		/// <summary>
		///	Implements the OnDisconnection method of the IDTExtensibility2 interface.
		///	Receives notification that the Add-in is being unloaded.
		/// </summary>
		/// <param name='removeMode'>
		///	Describes how the Add-in is being unloaded.
		/// </param>
		/// <param name='custom'>
		///	Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
		{	
			// Force handler to shutdown all work
			_documentHandler.HandleDisconnection();

			_outfile.Dispose();
		}

		/// <summary>
		///	Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
		///	Receives notification that the collection of Add-ins has changed.
		/// </summary>
		/// <param name='custom'>
		///	Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnAddInsUpdate(ref Array custom) { }

		/// <summary>
		///	Implements the OnStartupComplete method of the IDTExtensibility2 interface.
		///	Receives notification that the host application has completed loading.
		/// </summary>
		/// <param name='custom'>
		///	Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom) { }

		/// <summary>
		///	Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
		///	Receives notification that the host application is being unloaded.
		/// </summary>
		/// <param name='custom'>
		///	Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom) { }
		
		/// <summary>
		/// Handle document closed event
		/// </summary>
		/// <param name="docFullPath">
		/// Document full path.
		/// </param>
		private void HandleDocumentClosed(string docFullPath)
		{
			Debug.Assert(!string.IsNullOrEmpty(docFullPath));

			// Write text to a file in UI thread is not a very good idea.
			// We can move this logic into background thread, but it will be overkill
			// for this simple case.

			try
			{
				// Write path to output file.
				_outfile.WriteLine(docFullPath);
				
				// Immediatly flush data.
				_outfile.Flush();
			}
			catch (Exception e)
			{
				ErrorHandler.ReportError(e);
			}
		}
    }
}
