using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmChangePass : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private IContainer components = null;

	private Label lblMessage;

	private Label label1;

	private TextBox txtOldPass;

	private Label label9;

	private Button btnExit;

	private Button btnSave;

	private TextBox txtNewpass;

	private Label label2;

	private TextBox txtConfirmPass;

	private Label label3;

	public frmChangePass()
	{
		InitializeComponent();
	}

	private void btnSave_Click(object sender, EventArgs e)
	{
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		string valueNoParam = DBUtil.getValueNoParam("select matkhau from nhanvien where nv_id='" + Globals.gMaNV + "'");
		if (valueNoParam != ((Control)txtOldPass).Text)
		{
			MessageBox.Show("Mật khẩu hiện tại không đúng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			return;
		}
		if (((Control)txtNewpass).Text == "")
		{
			MessageBox.Show("Mật khẩu không được rỗng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			return;
		}
		if (((Control)txtNewpass).Text != ((Control)txtConfirmPass).Text)
		{
			MessageBox.Show("Mật khẩu mới và Nhập lại mật khẩu mới không giống nhau.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			return;
		}
		string query = "UPDATE nhanvien SET matkhau = @matkhau WHERE nv_id = @id";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@matkhau", ((Control)txtNewpass).Text);
		dataFetcher.addParameterWithValue("@id", Globals.gMaNV);
		int num = dataFetcher.execute();
		if (num == -1)
		{
			((Control)lblMessage).Text = "Không lưu được.";
		}
		else
		{
			((Control)lblMessage).Text = "Đã lưu.";
		}
	}

	private void btnExit_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void frmChangePass_Load(object sender, EventArgs e)
	{
		DBUtil.LoadMaunen((Form)(object)this);
		((Control)lblMessage).Text = "";
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		base.Dispose(disposing);
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
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_00a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b0: Expected O, but got Unknown
		//IL_0131: Unknown result type (might be due to invalid IL or missing references)
		//IL_013b: Expected O, but got Unknown
		//IL_01b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01bc: Expected O, but got Unknown
		//IL_023d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0247: Expected O, but got Unknown
		//IL_02ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c4: Expected O, but got Unknown
		//IL_0345: Unknown result type (might be due to invalid IL or missing references)
		//IL_034f: Expected O, but got Unknown
		//IL_03de: Unknown result type (might be due to invalid IL or missing references)
		//IL_03e8: Expected O, but got Unknown
		//IL_04c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_04cf: Expected O, but got Unknown
		//IL_058d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0597: Expected O, but got Unknown
		//IL_061b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0625: Expected O, but got Unknown
		//IL_0795: Unknown result type (might be due to invalid IL or missing references)
		//IL_079f: Expected O, but got Unknown
		//IL_07b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_07be: Expected O, but got Unknown
		//IL_07c1: Unknown result type (might be due to invalid IL or missing references)
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmChangePass));
		lblMessage = new Label();
		label1 = new Label();
		txtOldPass = new TextBox();
		label9 = new Label();
		txtNewpass = new TextBox();
		label2 = new Label();
		btnExit = new Button();
		btnSave = new Button();
		txtConfirmPass = new TextBox();
		label3 = new Label();
		((Control)this).SuspendLayout();
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblMessage).Location = new Point(171, 50);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(74, 20);
		((Control)lblMessage).TabIndex = 144;
		((Control)lblMessage).Text = "Message";
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 20.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(168, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(222, 31);
		((Control)label1).TabIndex = 143;
		((Control)label1).Text = "ĐỔI MẬT KHẨU";
		((Control)txtOldPass).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtOldPass).Location = new Point(174, 77);
		((Control)txtOldPass).Name = "txtOldPass";
		txtOldPass.PasswordChar = '*';
		((Control)txtOldPass).Size = new Size(281, 26);
		((Control)txtOldPass).TabIndex = 138;
		((Control)label9).AutoSize = true;
		((Control)label9).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label9).Location = new Point(2, 77);
		((Control)label9).Name = "label9";
		((Control)label9).Size = new Size(130, 20);
		((Control)label9).TabIndex = 142;
		((Control)label9).Text = "Mật khẩu hiện tại";
		((Control)txtNewpass).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtNewpass).Location = new Point(174, 106);
		((Control)txtNewpass).Name = "txtNewpass";
		txtNewpass.PasswordChar = '*';
		((Control)txtNewpass).Size = new Size(281, 26);
		((Control)txtNewpass).TabIndex = 139;
		((Control)label2).AutoSize = true;
		((Control)label2).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label2).Location = new Point(4, 106);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(104, 20);
		((Control)label2).TabIndex = 137;
		((Control)label2).Text = "Mật khẩu mới";
		((Control)btnExit).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnExit).FlatAppearance.BorderSize = 0;
		((Control)btnExit).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnExit).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnExit).ImageAlign = (ContentAlignment)16;
		((Control)btnExit).Location = new Point(318, 170);
		((Control)btnExit).Name = "btnExit";
		((Control)btnExit).Size = new Size(137, 41);
		((Control)btnExit).TabIndex = 141;
		((Control)btnExit).Text = "        Thoát";
		((ButtonBase)btnExit).UseVisualStyleBackColor = true;
		((Control)btnExit).Click += btnExit_Click;
		((Control)btnSave).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnSave).FlatAppearance.BorderSize = 0;
		((Control)btnSave).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave).Image = (Image)(object)Resources.Save;
		((ButtonBase)btnSave).ImageAlign = (ContentAlignment)16;
		((Control)btnSave).Location = new Point(175, 170);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(137, 41);
		((Control)btnSave).TabIndex = 140;
		((Control)btnSave).Text = "        Lưu";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		((Control)txtConfirmPass).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtConfirmPass).Location = new Point(174, 135);
		((Control)txtConfirmPass).Name = "txtConfirmPass";
		txtConfirmPass.PasswordChar = '*';
		((Control)txtConfirmPass).Size = new Size(281, 26);
		((Control)txtConfirmPass).TabIndex = 146;
		((Control)label3).AutoSize = true;
		((Control)label3).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label3).Location = new Point(4, 135);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(165, 20);
		((Control)label3).TabIndex = 145;
		((Control)label3).Text = "Nhập lại mật khẩu mới";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(495, 223);
		((Control)this).Controls.Add((Control)(object)txtConfirmPass);
		((Control)this).Controls.Add((Control)(object)label3);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)txtOldPass);
		((Control)this).Controls.Add((Control)(object)label9);
		((Control)this).Controls.Add((Control)(object)btnExit);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Controls.Add((Control)(object)txtNewpass);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmChangePass";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Đổi mật khẩu";
		((Form)this).Load += frmChangePass_Load;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
