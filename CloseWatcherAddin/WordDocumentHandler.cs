using System;
using System.Diagnostics;

using Word = Microsoft.Office.Interop.Word;
using WordEvents = Microsoft.Office.Interop.Word.ApplicationEvents4_Event;

namespace CloseWatcherAddin
{
	/// <summary>
	/// Implements Word-specific event handling
	/// </summary>
	class WordDocumentHandler : IDocumentHandler
	{
		public event ClosedHandler DocumentClosed;

		private Word.Document _closingDoc;

		public void HandleDisconnection() {}

		public WordDocumentHandler(WordEvents events)
		{
			Debug.Assert(events != null);

			events.DocumentBeforeClose += (Word.Document doc, ref bool cancel) => _closingDoc = doc;
			events.WindowActivate += (doc, window) => _closingDoc = null;
			events.WindowDeactivate += (doc, window) =>
				{	
					// To raise event for clear or unsaved documents, remove this check.
					if (doc.Path == String.Empty)
						return;

					if (doc == _closingDoc)
					{
						if (DocumentClosed != null)
							DocumentClosed(doc.FullName);
						
						_closingDoc = null;
					}
				};
		}
	}
}
