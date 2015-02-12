using System;
using System.Diagnostics;

using WordEvents = Microsoft.Office.Interop.Word.ApplicationEvents4_Event;
using ExcelEvents = Microsoft.Office.Interop.Excel.AppEvents_Event;
using PowerPointEvents = Microsoft.Office.Interop.PowerPoint.EApplication_Event;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace CloseWatcherAddin
{
	/// <summary>
	/// Simple factory to encapsulate handler construction
	/// depending on office application type
	/// </summary>
	static class DocumentHandlerFactory
	{
		/// <summary>
		///	Creates specific implementation behind IDocumentHandler
		/// that can raise event after document close.
		/// </summary>
		/// <param name='officeAapplication'>
		/// Office host application object.
		/// </param>
		public static IDocumentHandler CreateDocumentHandler(object officeAapplication)
		{
			Debug.Assert(officeAapplication != null);

			// This code is a little ugly, but for this simple case that's ok.
			if (officeAapplication is WordEvents)
				return new WordDocumentHandler((WordEvents)officeAapplication);
			if (officeAapplication is ExcelEvents)
				return new ExcelDocumentHandler((ExcelEvents)officeAapplication);
			if (officeAapplication is PowerPointEvents)
				return new PowerPointDocumentHandler((PowerPoint.Application)officeAapplication);

			throw new ApplicationException("Unknown Office application type");
		}
	}
}
