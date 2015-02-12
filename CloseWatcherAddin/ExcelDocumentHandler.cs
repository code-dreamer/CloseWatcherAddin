using System;
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;
using ExcelEvents = Microsoft.Office.Interop.Excel.AppEvents_Event;

namespace CloseWatcherAddin
{
	/// <summary>
	/// Implements Excel-specific event handling
	/// </summary>
	class ExcelDocumentHandler : IDocumentHandler
	{
		private Excel.Workbook _closingWorkbook;

		public event ClosedHandler DocumentClosed;

		public void HandleDisconnection() { }

		public ExcelDocumentHandler(ExcelEvents events)
		{
			Debug.Assert(events != null);

			events.WorkbookBeforeClose += (Excel.Workbook workbook, ref bool cancel) => _closingWorkbook = workbook;
			events.WindowActivate += (doc, window) => _closingWorkbook = null;
			events.WindowDeactivate += (workbook, window) =>
			{
				// To raise event for clear or unsaved documents, remove this check.
				if (workbook.Path == String.Empty)
					return;

				if (workbook == _closingWorkbook)
				{
					if (DocumentClosed != null)
						DocumentClosed(workbook.FullName);

					_closingWorkbook = null;
				}
			};
		}
	}
}
