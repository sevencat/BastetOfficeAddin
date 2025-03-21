using System;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using NetOffice.WordApi;
using NetOffice.WordApi.Tools;

namespace Bastet.OfficeAddin;

[ComVisible(true)]
[Guid("9777f77a-f286-49d6-817b-0500c95b89c2")]
[ProgId("Bastet.OfficeAddin.MyWordAddin")]
[COMAddin("Bastet Word Addin", "Wordaddin By Bastet", LoadBehavior.LoadAtStartup)]
public class MyWordAddin : COMAddin
{
	public MyWordAddin()
	{
		this.OnConnection += MyAddin_OnConnection;
	}

	private void MyAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst,
		ref Array custom)
	{
		this.Application.DocumentOpenEvent += Application_DocumentOpenEvent;
	}

	private void Application_DocumentOpenEvent(Document doc)
	{
		using (doc)
		{
			// start working with the document
		}
	}
}