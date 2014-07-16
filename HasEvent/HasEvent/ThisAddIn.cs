using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace HasEvent {

	/// The observer pattern is needed to do eventing in COM. That's because
	/// COM does not recognizes delegates and therefore does not recognize native 
	/// .NET events.
	/// El Zorko explains this better in 
	/// http://stackoverflow.com/questions/1985451/exposing-a-net-class-which-has-events-to-com
	/// <seealso cref="http://en.wikipedia.org/wiki/Observer_pattern"/>

	/// <summary>
	/// Third-party interested in the Save event
	/// </summary>
	[ComVisible(true)]
	[Guid("12810C22-3FF2-4fc2-A7FD-7E1034462EB0")]
	public interface IObserver {

		// Called by the first party when the first party handled the
		// Word interop BeforeSave event
		void AfterSave();
	}

	/// <summary>
	/// The first party needing to gain first-fire access to the BeforeSave event
	/// </summary>
	[ComVisible(true)]
	[Guid("02810C22-3FF2-4fc2-A7FD-7E1034462EB0")]
	public interface ISubject {
		void Listen(IObserver observer);
		void Fire();
	}

	[ComVisible(true)]
	[Guid("02810C22-3FF2-4fc2-A7FD-5E1034462EB0")]
	[ClassInterface(ClassInterfaceType.None)]
	public partial class ThisAddIn : ISubject {

		private readonly List<IObserver> _observers = new List<IObserver>();

		public void Listen(IObserver observer) {
			_observers.Add(observer);
		}

		private void ThisAddIn_Startup(object sender, System.EventArgs e) {
			Application.DocumentBeforeSave += Application_DocumentBeforeSave;
		}

		void Application_DocumentBeforeSave(Document Doc, ref bool SaveAsUI, ref bool Cancel) {
			/* do all kinds of stuff */
			// Merp();

			/* let other interested parties outside this addin know */
			Fire();
		}
		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)  {
			/* nothing */
		}

		public void Fire() {
			foreach (IObserver observer in _observers) {
				observer.AfterSave();
			}
		}

		/// <summary>
		/// Required (and undocumented on the COMAddin MSDN docs) override.
		/// </summary>
		/// <returns></returns>
		protected override object RequestComAddInAutomationService() {
			return this;
		}

		private void InternalStartup() {
			Startup += ThisAddIn_Startup;
			Shutdown += ThisAddIn_Shutdown;
		}

	}
}