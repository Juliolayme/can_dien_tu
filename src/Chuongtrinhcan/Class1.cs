using System;
using System.Drawing;
using System.Windows.Forms;

namespace Chuongtrinhcan;

public class Class1
{
	public static string gMatkhau = "";

	public static DialogResult InputBox(string title, string promptText, ref string value)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Expected O, but got Unknown
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		//IL_00b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0174: Unknown result type (might be due to invalid IL or missing references)
		//IL_0179: Unknown result type (might be due to invalid IL or missing references)
		//IL_0183: Unknown result type (might be due to invalid IL or missing references)
		//IL_0185: Unknown result type (might be due to invalid IL or missing references)
		//IL_0189: Unknown result type (might be due to invalid IL or missing references)
		Form val = new Form();
		Label val2 = new Label();
		TextBox val3 = new TextBox();
		Button val4 = new Button();
		Button val5 = new Button();
		((Control)val).Text = title;
		((Control)val2).Text = promptText;
		((Control)val3).Text = value;
		((Control)val4).Text = "OK";
		((Control)val5).Text = "Cancel";
		val4.DialogResult = (DialogResult)1;
		val5.DialogResult = (DialogResult)2;
		((Control)val2).SetBounds(9, 20, 372, 13);
		((Control)val3).SetBounds(12, 36, 372, 20);
		((Control)val4).SetBounds(228, 72, 75, 23);
		((Control)val5).SetBounds(309, 72, 75, 23);
		((Control)val2).AutoSize = true;
		((Control)val3).Anchor = (AnchorStyles)((int)((Control)val3).Anchor | 8);
		((Control)val4).Anchor = (AnchorStyles)10;
		((Control)val5).Anchor = (AnchorStyles)10;
		val.ClientSize = new Size(396, 107);
		((Control)val).Controls.AddRange((Control[])(object)new Control[4]
		{
			(Control)val2,
			(Control)val3,
			(Control)val4,
			(Control)val5
		});
		val.ClientSize = new Size(Math.Max(300, ((Control)val2).Right + 10), val.ClientSize.Height);
		val.FormBorderStyle = (FormBorderStyle)3;
		val.StartPosition = (FormStartPosition)1;
		val.MinimizeBox = false;
		val.MaximizeBox = false;
		val.AcceptButton = (IButtonControl)(object)val4;
		val.CancelButton = (IButtonControl)(object)val5;
		DialogResult result = val.ShowDialog();
		value = ((Control)val3).Text;
		return result;
	}
}
