using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Chuongtrinhcan;

public class frmCopyright2 : Form
{
	private IContainer components = null;

	private TextBox textBox1;

	private Label label1;

	private Label label2;

	private Button button1;

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		((Form)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_0057: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		//IL_00d0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00da: Expected O, but got Unknown
		//IL_0154: Unknown result type (might be due to invalid IL or missing references)
		//IL_015e: Expected O, but got Unknown
		//IL_01ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d8: Expected O, but got Unknown
		//IL_02f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0303: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmCopyright2));
		textBox1 = new TextBox();
		label1 = new Label();
		label2 = new Label();
		button1 = new Button();
		((Control)this).SuspendLayout();
		((Control)textBox1).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)textBox1).Location = new Point(136, 53);
		((Control)textBox1).Name = "textBox1";
		((Control)textBox1).Size = new Size(248, 22);
		((Control)textBox1).TabIndex = 0;
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(22, 56);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(96, 16);
		((Control)label1).TabIndex = 1;
		((Control)label1).Text = "Nhập dãy số";
		((Control)label2).AutoSize = true;
		((Control)label2).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label2).Location = new Point(79, 9);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(325, 25);
		((Control)label2).TabIndex = 2;
		((Control)label2).Text = "NHẬP THÔNG SỐ HỆ THỐNG";
		((Control)button1).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)button1).Location = new Point(182, 94);
		((Control)button1).Name = "button1";
		((Control)button1).Size = new Size(118, 38);
		((Control)button1).TabIndex = 3;
		((Control)button1).Text = "OK";
		((ButtonBase)button1).UseVisualStyleBackColor = true;
		((Control)button1).Click += button1_Click;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(483, 158);
		((Control)this).Controls.Add((Control)(object)button1);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)textBox1);
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Control)this).Name = "frmCopyright";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Thong so he thong";
		((Form)this).Load += frmCopyright_Load;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}

	public frmCopyright2()
	{
		InitializeComponent();
	}

	private void button1_Click(object sender, EventArgs e)
	{
		//IL_00e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ca: Unknown result type (might be due to invalid IL or missing references)
		DateTime now = DateTime.Now;
		string text = "";
		text = now.Year + now.Month + now.Day.ToString("00") + now.Day % 2 + now.Day % 5 + now.Day % 7;
		if (((Control)textBox1).Text.Trim() == text)
		{
			Registry.SetValue("HKEY_CURRENT_USER", "cprDTH", "sondt");
			MessageBox.Show("Dãy số OK. Vui lòng chạy lại chương trình.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			((Form)this).Close();
		}
		else
		{
			MessageBox.Show("Dãy số không đúng. Vui lòng liên hệ nhà cung cấp.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
		}
	}

	private void frmCopyright_Load(object sender, EventArgs e)
	{
	}
}
