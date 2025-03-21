using System;
using System.Runtime.InteropServices;
using NetOffice.ExcelApi.Tools;
using NetOffice.Tools;


namespace Bastet.OfficeAddin;

//e4866de5-adb8-60d8-1d7c-463da09e39d8
[ComVisible(true)]
[Guid("e4866de5-adb8-60d8-1d7c-463da09e39d8")]
[ProgId("Bastet.OfficeAddin.MyExcelAddin")]
[COMAddin("Bastet Word Addin", "Wordaddin By Bastet", LoadBehavior.LoadAtStartup)]
public class MyExcelAddin : COMAddin
{
	public MyExcelAddin()
	{
		this.OnConnection += Addin_OnConnection;
		OnStartupComplete += Addin_OnStartupComplete;
		OnDisconnection += Addin_OnDisconnection;
	}

	private void Addin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
	{
	}

	private void Addin_OnStartupComplete(ref Array custom)
	{
		Console.WriteLine("Excel Version is {0}", Application.Version);
	}

	private void Addin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
	{
	}
}