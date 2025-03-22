using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Application = NetOffice.WordApi.Application;

namespace Bastet.OfficeAddin
{
	// //e4d7fe8e-f7d4-ccf6-319c-c66ba44d355c
	// [ComVisible(true)]
	// [Guid("e4d7fe8e-f7d4-ccf6-319c-c66ba44d355c")]
	// [ProgId("Bastet.OfficeAddin.Word.WordPanel")]
	[ProgId("Bastet.OfficeAddin.WordPanel")]
	[ClassInterface(ClassInterfaceType.AutoDispatch)]
	public partial class WordPanel : UserControl, NetOffice.WordApi.Tools.ITaskPane
	{
		public WordPanel()
		{
			InitializeComponent();
		}

		public void OnConnection(Application application, _CustomTaskPane parentPane, object[] customArguments)
		{
		}

		public void OnDisconnection()
		{
		}

		public void OnDockPositionChanged(MsoCTPDockPosition position)
		{
		}

		public void OnVisibleStateChanged(bool visible)
		{
		}
	}
}