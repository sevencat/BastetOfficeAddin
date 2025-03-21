using System.Runtime.InteropServices;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;

namespace Bastet.OfficeAddin;

[ComVisible(true)]
[Guid("35b9ab06-e04b-58e8-12d5-829021765651")]
[ProgId("Bastet.OfficeAddin.MyPptAddin")]
[COMAddin("Bastet PowerPoint Addin", "PowerPointaddin By Bastet", LoadBehavior.LoadAtStartup)]
public class MyPptAddin : COMAddin
{
	public MyPptAddin()
	{
	}
}