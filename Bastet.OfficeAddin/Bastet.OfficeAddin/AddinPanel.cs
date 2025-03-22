using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;


namespace Bastet.OfficeAddin;

[ProgId("Bastet.OfficeAddin.AddinPanel")]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public partial class AddinPanel : UserControl
	, NetOffice.WordApi.Tools.ITaskPane
	, NetOffice.ExcelApi.Tools.ITaskPane
	, NetOffice.PowerPointApi.Tools.ITaskPane
{
	public AddinPanel()
	{
		InitializeComponent();
	}

	#region word相关

	public void OnConnection(NetOffice.WordApi.Application application, _CustomTaskPane parentPane,
		object[] customArguments)
	{
	}

	void NetOffice.WordApi.Tools.ITaskPane.OnDisconnection()
	{
	}

	void NetOffice.WordApi.Tools.ITaskPane.OnDockPositionChanged(MsoCTPDockPosition position)
	{
	}

	void NetOffice.WordApi.Tools.ITaskPane.OnVisibleStateChanged(bool visible)
	{
	}

	#endregion

	#region excel相关

	public void OnConnection(NetOffice.ExcelApi.Application application, _CustomTaskPane parentPane,
		object[] customArguments)
	{
	}

	void NetOffice.ExcelApi.Tools.ITaskPane.OnDockPositionChanged(MsoCTPDockPosition position)
	{
	}

	void NetOffice.ExcelApi.Tools.ITaskPane.OnVisibleStateChanged(bool visible)
	{
	}

	void NetOffice.ExcelApi.Tools.ITaskPane.OnDisconnection()
	{
	}

	#endregion

	#region ppt相关

	public void OnConnection(NetOffice.PowerPointApi.Application application, _CustomTaskPane parentPane,
		object[] customArguments)
	{
	}


	void NetOffice.PowerPointApi.Tools.ITaskPane.OnDockPositionChanged(MsoCTPDockPosition position)
	{
	}

	void NetOffice.PowerPointApi.Tools.ITaskPane.OnVisibleStateChanged(bool visible)
	{
	}

	void NetOffice.PowerPointApi.Tools.ITaskPane.OnDisconnection()
	{
	}

	#endregion
}