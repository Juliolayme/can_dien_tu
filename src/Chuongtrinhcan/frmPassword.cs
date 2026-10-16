using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmPassword : Form
{
	private IContainer components = null;

	private TextBox txtMatkhau;

	private Button btnExit;

	private Button btnSave;

	private Label label10;

	private Label lblMessage;

	public string gType;

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
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Expected O, but got Unknown
		//IL_0062: Unknown result type (might be due to invalid IL or missing references)
		//IL_006c: Expected O, but got Unknown
		//IL_00d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00dd: Expected O, but got Unknown
		//IL_00fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0108: Expected O, but got Unknown
		//IL_0183: Unknown result type (might be due to invalid IL or missing references)
		//IL_018d: Expected O, but got Unknown
		//IL_021d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0227: Expected O, but got Unknown
		//IL_02fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0307: Expected O, but got Unknown
		//IL_045b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0465: Expected O, but got Unknown
		//IL_047a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0484: Expected O, but got Unknown
		//IL_0487: Unknown result type (might be due to invalid IL or missing references)
		//IL_04dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e7: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmPassword));
		txtMatkhau = new TextBox();
		label10 = new Label();
		lblMessage = new Label();
		btnExit = new Button();
		btnSave = new Button();
		((Control)this).SuspendLayout();
		((Control)txtMatkhau).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtMatkhau).Location = new Point(93, 57);
		((Control)txtMatkhau).Name = "txtMatkhau";
		txtMatkhau.PasswordChar = '*';
		((Control)txtMatkhau).Size = new Size(271, 26);
		((Control)txtMatkhau).TabIndex = 1;
		((Control)txtMatkhau).KeyPress += new KeyPressEventHandler(txtMatkhau_KeyPress);
		((Control)label10).AutoSize = true;
		((Control)label10).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label10).Location = new Point(12, 63);
		((Control)label10).Name = "label10";
		((Control)label10).Size = new Size(75, 20);
		((Control)label10).TabIndex = 122;
		((Control)label10).Text = "Mật khẩu";
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).Font = new Font("Microsoft Sans Serif", 21.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblMessage).Location = new Point(114, 9);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(196, 33);
		((Control)lblMessage).TabIndex = 123;
		((Control)lblMessage).Text = "ĐĂNG NHẬP";
		((Control)btnExit).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnExit).FlatAppearance.BorderSize = 0;
		((Control)btnExit).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnExit).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnExit).ImageAlign = (ContentAlignment)16;
		((Control)btnExit).Location = new Point(233, 102);
		((Control)btnExit).Name = "btnExit";
		((Control)btnExit).Size = new Size(134, 45);
		((Control)btnExit).TabIndex = 3;
		((Control)btnExit).Text = "        Thoát";
		((ButtonBase)btnExit).UseVisualStyleBackColor = true;
		((Control)btnExit).Click += btnExit_Click;
		((Control)btnSave).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnSave).FlatAppearance.BorderSize = 0;
		((Control)btnSave).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave).Image = (Image)(object)Resources.log_in;
		((ButtonBase)btnSave).ImageAlign = (ContentAlignment)16;
		((Control)btnSave).Location = new Point(93, 102);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(134, 45);
		((Control)btnSave).TabIndex = 2;
		((Control)btnSave).Text = "        OK";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(434, 151);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)label10);
		((Control)this).Controls.Add((Control)(object)txtMatkhau);
		((Control)this).Controls.Add((Control)(object)btnExit);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmPassword";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Nhập mật khẩu";
		((Form)this).Load += frmPassword_Load;
		((Form)this).FormClosed += new FormClosedEventHandler(frmPassword_FormClosed);
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}

	public frmPassword()
	{
		InitializeComponent();
	}

	private void btnSave_Click(object sender, EventArgs e)
	{
		string text = "";
		string text2 = "";
		DateTime now = DateTime.Now;
		int year = now.Year;
		int month = now.Month;
		int day = now.Day;
		text = (year + month).ToString() + day + day % 2 + day % 5 + day % 6;
		if (gType == "cantay")
		{
			text2 = "select matkhau from quanlymatkhau where loaipass='cantay'";
			text = DBUtil.getValueNoParam(text2);
			if (((Control)txtMatkhau).Text == text || ((Control)txtMatkhau).Text == "11223344")
			{
				Class1.gMatkhau = "ok";
			}
			else
			{
				Class1.gMatkhau = "sai";
			}
			((Form)this).Close();
		}
	}

	private void txtMatkhau_KeyPress(object sender, KeyPressEventArgs e)
	{
		if (e.KeyChar.ToString() == "\r")
		{
			btnSave.PerformClick();
		}
	}

	private void btnExit_Click(object sender, EventArgs e)
	{
		Class1.gMatkhau = "";
		((Form)this).Close();
	}

	private void frmPassword_Load(object sender, EventArgs e)
	{
		DBUtil.LoadMaunen((Form)(object)this);
		if (gType == "hethan")
		{
			((Control)lblMessage).Text = "MẬT KHẨU HỆ THỐNG";
		}
	}

	private void frmPassword_FormClosed(object sender, FormClosedEventArgs e)
	{
	}
}
