using System;
using System.Windows.Forms;
using HasEvent;
using Office = Microsoft.Office.Core;

namespace ConsumeEvent {
	public partial class ThisAddIn : IObserver {

		private Office.COMAddIn hasEvent;

		public void AfterSave() {
			MessageBox.Show("Before save event fired");
		}

		private void ThisAddIn_Startup(object sender, EventArgs e) {
			Office.COMAddIn comAddin = FindComAddinByName("HasEvent");
			ISubject hd = ObtainSubject(comAddin);
			hd.Listen(this);
		}

		private Office.COMAddIn FindComAddinByName(string name) {
			foreach (Office.COMAddIn comAddIn in Application.COMAddIns) {
				if (comAddIn.Description == name) {
					return comAddIn;
				}
			}
			return null;
		}

		/// <summary>
		/// Obtains a reference to a typed ComAddin project.
		/// The type is determined in the other ComAddin project.
		/// </summary>
		/// <seealso cref="http://blogs.msdn.com/b/andreww/archive/2008/08/13/comaddins-race-condition.aspx"/>
		/// <param name="comAddin"></param>
		/// <returns></returns>
		private ISubject ObtainSubject(Office.COMAddIn comAddin) {
			object obj = null;
			int tries = 0;
			// 50 * 100 miliseconds = 5000 milliseconds == 5 seconds
			while (obj == null && tries < 50) {
				obj = comAddin.Object;
				System.Threading.Thread.Sleep(100);
				tries++;
			}
			ISubject subject = obj as ISubject;
			return subject;
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e) {
			/* nothing */
		}

		private void InternalStartup() {
			Startup += ThisAddIn_Startup;
			Shutdown += ThisAddIn_Shutdown;
		}

	}
}