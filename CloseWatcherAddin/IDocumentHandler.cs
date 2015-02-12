
namespace CloseWatcherAddin
{
	public delegate void ClosedHandler(string fullPath);

	/// <summary>
	/// Common interface to receive event after document was closed
	/// </summary>
	interface IDocumentHandler
	{	
		// In .NET 4.5 and higher event declaration can be:
		// public event EventHandler<string> DocumentClosed;
		event ClosedHandler DocumentClosed;

		/// <summary>
		/// Notify about disconnect.
		/// </summary>
		void HandleDisconnection();
	}
}
