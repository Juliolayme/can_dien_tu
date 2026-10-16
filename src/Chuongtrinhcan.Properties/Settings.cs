using System.CodeDom.Compiler;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;

namespace Chuongtrinhcan.Properties;

[CompilerGenerated]
[GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "9.0.0.0")]
internal sealed class Settings : ApplicationSettingsBase
{
	private static Settings defaultInstance = (Settings)(object)SettingsBase.Synchronized((SettingsBase)(object)new Settings());

	public static Settings Default => defaultInstance;

	[DefaultSettingValue("Arial, 12pt")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public Font font
	{
		get
		{
			//IL_000c: Unknown result type (might be due to invalid IL or missing references)
			//IL_0012: Expected O, but got Unknown
			return (Font)((SettingsBase)this)["font"];
		}
		set
		{
			((SettingsBase)this)["font"] = value;
		}
	}

	[DefaultSettingValue("Microsoft Sans Serif, 12pt")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public Font font12
	{
		get
		{
			//IL_000c: Unknown result type (might be due to invalid IL or missing references)
			//IL_0012: Expected O, but got Unknown
			return (Font)((SettingsBase)this)["font12"];
		}
		set
		{
			((SettingsBase)this)["font12"] = value;
		}
	}
}
