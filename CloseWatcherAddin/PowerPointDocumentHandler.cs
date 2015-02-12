using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using PowerPointEvents = Microsoft.Office.Interop.PowerPoint.EApplication_Event;
using Timer = System.Threading.Timer;

namespace CloseWatcherAddin
{
	/// <summary>
	/// Implements PowerPoint-specific event handling
	/// </summary>
	class PowerPointDocumentHandler : IDocumentHandler
	{
		private static Timer _presentationsChecker;
		private const int PresentationsCheckTimeout = 500;

		private readonly PowerPoint.Application _powerPointApplication;

		// Collection of Power Point documents that can be closed
		private readonly List<string> _lastPresentationsSnapshot = new List<string>();

		public event ClosedHandler DocumentClosed;

		/// <summary>
		///	Alert about watched documents if any
		/// </summary>
		public void HandleDisconnection()
		{
			foreach (string presentationFullName in _lastPresentationsSnapshot)
			{
				OnDocumentClosed(presentationFullName);
			}
		}

		public PowerPointDocumentHandler(PowerPoint.Application powerPointApplication)
		{
			Debug.Assert(powerPointApplication != null);

			_powerPointApplication = powerPointApplication;

			// Watch for open or fresh-saved documents
			powerPointApplication.PresentationOpen += RegisterNewPresentation;
			powerPointApplication.PresentationSave += RegisterNewPresentation;
		}

		/// <summary>
		///	Add new presentation and start watch timer if needed.
		/// </summary>
		private void RegisterNewPresentation(PowerPoint.Presentation presentation)
		{
			Debug.Assert(presentation != null);

			bool isPresentationExist = _lastPresentationsSnapshot.Any(currPresentationFullName => presentation.FullName == currPresentationFullName);
			if (!isPresentationExist)
			{
				_lastPresentationsSnapshot.Add(presentation.FullName);

				if (_presentationsChecker == null)
					_presentationsChecker = new Timer(OnTimer, null, PresentationsCheckTimeout, PresentationsCheckTimeout);
			}
		}

		/// <summary>
		///	Check if known documents still exist.
		/// </summary>
		private void OnTimer(object param)
		{
			Debug.Assert(param == null);
			Debug.Assert(_powerPointApplication != null);

			// Check if some document not exist anymore,
			// alert DocumentClosed subscribers
			// and remove closed documents from collection.
			_lastPresentationsSnapshot.RemoveAll(presentationFullName =>
			{
				bool isPresentationExist =
					_powerPointApplication.Presentations.Cast<PowerPoint.Presentation>()
						.Any(currPresentation => currPresentation.FullName == presentationFullName);

				if (isPresentationExist)
					return false;

				OnDocumentClosed(presentationFullName);
				return true;
			});
		}

		private void OnDocumentClosed(string presentationFullName)
		{
			Debug.Assert(string.IsNullOrEmpty(presentationFullName));

			if (DocumentClosed != null)
				DocumentClosed(presentationFullName);
		}
	}
}
