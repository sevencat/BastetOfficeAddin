using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.Tools;
using NetOffice.WordApi;
using NetOffice.WordApi.Tools;

namespace Bastet.OfficeAddin;

[ComVisible(true)]
[Guid("9777f77a-f286-49d6-817b-0500c95b89c2")]
[ProgId("Bastet.OfficeAddin.MyWordAddin")]
[COMAddin("Bastet Word Addin", "Wordaddin By Bastet", LoadBehavior.LoadAtStartup)]
[CustomPane(typeof(WordPanel), "Simple Taskpane", true, PaneDockPosition.msoCTPDockPositionRight)]
public class MyWordAddin : COMAddin
{
	Application wordApp;
	//ICTPFactory ctpFactory;

	public MyWordAddin()
	{
		int xx = 1;
	}

	private void MyAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst,
		ref Array custom)
	{
		wordApp = new Application(null, application);
		this.Application.DocumentOpenEvent += Application_DocumentOpenEvent;
	}

	private void Application_DocumentOpenEvent(Document doc)
	{
		using (doc)
		{
			// start working with the document
		}
	}

	#region 回调

	//TODO 后面用zip文件统一管理资源。不依赖于vsproject
	public override string GetCustomUI(string RibbonID)
	{
		var ret = GetResourceText("Bastet.OfficeAddin.Word.WordRibbonUI.xml");
		return ret;
	}

	public void Ribbon_Load(IRibbonUI ribbonUI)
	{
		//这个会不会也有多份才对？
	}

	public void CommonWordFunc_Click(IRibbonControl control)
	{
		var id = control.Id;
		var tag = control.Tag;
		if (id == "aboutButton")
		{
			var panel = this.TaskPaneFactory.CreateCTP("Bastet.OfficeAddin.Word.WordPanel", "Example");
			panel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
			panel.Width = 550;
			panel.Visible = true;
		}
	}


	private static string GetResourceText(string resourceName)
	{
		Assembly asm = Assembly.GetExecutingAssembly();
		string[] resourceNames = asm.GetManifestResourceNames();
		for (int i = 0; i < resourceNames.Length; ++i)
		{
			if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
			{
				using (StreamReader resourceReader =
				       new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
				{
					if (resourceReader != null)
					{
						return resourceReader.ReadToEnd();
					}
				}
			}
		}

		return null;
	}

	#endregion
}