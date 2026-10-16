using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;
using Microsoft.Win32;

namespace Chuongtrinhcan;

public class frmLogin : Form
{
	private IContainer components = null;

	private Label lblMessage;

	private Label label1;

	private TextBox txtUserName;

	private Label label9;

	private Button btnExit;

	private Button btnSave;

	private TextBox txtPassword;

	private Label label2;

	private DataFetcher dataFetcher = DataFetcher.Instance;

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
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Expected O, but got Unknown
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005e: Expected O, but got Unknown
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_0090: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Expected O, but got Unknown
		//IL_0118: Unknown result type (might be due to invalid IL or missing references)
		//IL_0122: Expected O, but got Unknown
		//IL_0196: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a0: Expected O, but got Unknown
		//IL_0210: Unknown result type (might be due to invalid IL or missing references)
		//IL_021a: Expected O, but got Unknown
		//IL_028b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0295: Expected O, but got Unknown
		//IL_0300: Unknown result type (might be due to invalid IL or missing references)
		//IL_030a: Expected O, but got Unknown
		//IL_032b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0335: Expected O, but got Unknown
		//IL_03c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_03cf: Expected O, but got Unknown
		//IL_04ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b6: Expected O, but got Unknown
		//IL_064a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0654: Expected O, but got Unknown
		//IL_0669: Unknown result type (might be due to invalid IL or missing references)
		//IL_0673: Expected O, but got Unknown
		//IL_0676: Unknown result type (might be due to invalid IL or missing references)
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmLogin));
		lblMessage = new Label();
		label1 = new Label();
		txtUserName = new TextBox();
		label9 = new Label();
		txtPassword = new TextBox();
		label2 = new Label();
		btnExit = new Button();
		btnSave = new Button();
		((Control)this).SuspendLayout();
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblMessage).Location = new Point(125, 49);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(74, 20);
		((Control)lblMessage).TabIndex = 136;
		((Control)lblMessage).Text = "Message";
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 20.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(122, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(345, 31);
		((Control)label1).TabIndex = 135;
		((Control)label1).Text = "ĐĂNG NHẬP HỆ THỐNG";
		((Control)txtUserName).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtUserName).Location = new Point(127, 77);
		((Control)txtUserName).Name = "txtUserName";
		((Control)txtUserName).Size = new Size(283, 26);
		((Control)txtUserName).TabIndex = 130;
		((Control)label9).AutoSize = true;
		((Control)label9).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label9).Location = new Point(10, 77);
		((Control)label9).Name = "label9";
		((Control)label9).Size = new Size(116, 20);
		((Control)label9).TabIndex = 134;
		((Control)label9).Text = "Tên đăng nhập";
		((Control)txtPassword).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtPassword).Location = new Point(127, 105);
		((Control)txtPassword).Name = "txtPassword";
		txtPassword.PasswordChar = '*';
		((Control)txtPassword).Size = new Size(283, 26);
		((Control)txtPassword).TabIndex = 131;
		((Control)txtPassword).KeyDown += new KeyEventHandler(txtPassword_KeyDown);
		((Control)label2).AutoSize = true;
		((Control)label2).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label2).Location = new Point(12, 105);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(75, 20);
		((Control)label2).TabIndex = 129;
		((Control)label2).Text = "Mật khẩu";
		((Control)btnExit).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnExit).FlatAppearance.BorderSize = 0;
		((Control)btnExit).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnExit).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnExit).ImageAlign = (ContentAlignment)16;
		((Control)btnExit).Location = new Point(273, 142);
		((Control)btnExit).Name = "btnExit";
		((Control)btnExit).Size = new Size(137, 41);
		((Control)btnExit).TabIndex = 133;
		((Control)btnExit).Text = "        Thoát";
		((ButtonBase)btnExit).UseVisualStyleBackColor = true;
		((Control)btnExit).Click += btnExit_Click;
		((Control)btnSave).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnSave).FlatAppearance.BorderSize = 0;
		((Control)btnSave).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave).Image = (Image)(object)Resources.log_in;
		((ButtonBase)btnSave).ImageAlign = (ContentAlignment)16;
		((Control)btnSave).Location = new Point(130, 142);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(137, 41);
		((Control)btnSave).TabIndex = 132;
		((Control)btnSave).Text = "        Đăng nhập";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(505, 191);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)txtUserName);
		((Control)this).Controls.Add((Control)(object)label9);
		((Control)this).Controls.Add((Control)(object)btnExit);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Controls.Add((Control)(object)txtPassword);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmLogin";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Đăng nhập hệ thống";
		((Form)this).Load += frmLogin_Load;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}

	public frmLogin()
	{
		InitializeComponent();
	}

	private void btnSave_Click(object sender, EventArgs e)
	{
		//IL_0099: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a3: Expected O, but got Unknown
		try
		{
			string query = "select nv_id from nhanvien where nv_id=@nv_id and matkhau=@matkhau and trangthai =0";
			dataFetcher.setCommand(query, 1);
			dataFetcher.addParameterWithValue("@nv_id", ((Control)txtUserName).Text);
			dataFetcher.addParameterWithValue("@matkhau", ((Control)txtPassword).Text);
			DataSet dataSet = dataFetcher.fetchDataSet();
			if (dataSet.Tables[0].Rows.Count > 0)
			{
				Globals.gMaNV = ((Control)txtUserName).Text;
				frmMain frmMain2 = new frmMain();
				((Form)frmMain2).FormClosed += new FormClosedEventHandler(frm2_FormClosed);
				((Control)frmMain2).Show();
				((Control)this).Hide();
			}
			else
			{
				((Control)lblMessage).Text = "Sai Tên đăng nhập hoặc Mật khẩu.";
			}
		}
		catch
		{
			((Control)lblMessage).Text = "Sai Tên đăng nhập hoặc Mật khẩu.";
		}
	}

	private void frm2_FormClosed(object sender, FormClosedEventArgs e)
	{
		((Form)this).Close();
	}

	private void btnExit_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void frmLogin_Load(object sender, EventArgs e)
	{
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		DBUtil.LoadMaunen((Form)(object)this);
		string text = Registry.GetValue("HKEY_CURRENT_USER", "cprVietDucAnhChi", "").ToString();
		if (text != "vietduc123")
		{
			frmCopyright frmCopyright3 = new frmCopyright();
			((Form)frmCopyright3).ShowDialog();
			((Form)this).Close();
		}
		else
		{
			((Control)lblMessage).Text = "";
		}
	}

	private void txtPassword_KeyDown(object sender, KeyEventArgs e)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0009: Invalid comparison between Unknown and I4
		if ((int)e.KeyCode == 13)
		{
			btnSave.PerformClick();
		}
	}
}
