using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmCaidatMatkhau : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private string tong;

	private string cantay;

	private string sua;

	private string xoa;

	private string khunggia;

	private string xacxe;

	private string dssoxe;

	private string khoaxacxe;

	private string klbi;

	private string loaihang;

	private string dstienkhunggia;

	private string guimail;

	private string hienthicty;

	private IContainer components = null;

	private Label label1;

	private Button btnExit;

	private Button btnSave;

	private Label lblMessage;

	private TextBox txtPassmoi;

	private Label label8;

	private TextBox txtNhaplaiPass;

	private Label label2;

	public frmCaidatMatkhau()
	{
		InitializeComponent();
	}

	private void btnSave_Click(object sender, EventArgs e)
	{
		if (((Control)txtPassmoi).Text != ((Control)txtNhaplaiPass).Text)
		{
			((Control)lblMessage).Text = "Pass mới và Nhập lại pass mới không giống nhau.";
			return;
		}
		capNhatmatkhau(((Control)txtPassmoi).Text, "cantay");
		((Control)lblMessage).Text = "Đã lưu dữ liệu.";
	}

	private void capNhatmatkhau(string matKhau, string loaiPass)
	{
		try
		{
			string text = "";
			text = "UPDATE quanlymatkhau set matkhau=@matkhau where loaipass=@loaipass ";
			dataFetcher.setCommand(text, 1);
			dataFetcher.addParameterWithValue("@matkhau", matKhau);
			dataFetcher.addParameterWithValue("@loaipass", loaiPass);
			int num = dataFetcher.execute();
		}
		catch
		{
			((Control)lblMessage).Text = "Dữ liệu nhập không đúng, vui lòng nhập lại";
		}
	}

	private void btnExit_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void frmCaidatMatkhau_Load(object sender, EventArgs e)
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
		//IL_0129: Unknown result type (might be due to invalid IL or missing references)
		//IL_0133: Expected O, but got Unknown
		//IL_020c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0216: Expected O, but got Unknown
		//IL_02dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e7: Expected O, but got Unknown
		//IL_0369: Unknown result type (might be due to invalid IL or missing references)
		//IL_0373: Expected O, but got Unknown
		//IL_03f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_03fa: Expected O, but got Unknown
		//IL_046b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0475: Expected O, but got Unknown
		//IL_04f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fc: Expected O, but got Unknown
		//IL_0646: Unknown result type (might be due to invalid IL or missing references)
		//IL_0650: Expected O, but got Unknown
		//IL_0665: Unknown result type (might be due to invalid IL or missing references)
		//IL_066f: Expected O, but got Unknown
		//IL_0672: Unknown result type (might be due to invalid IL or missing references)
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmCaidatMatkhau));
		label1 = new Label();
		btnExit = new Button();
		btnSave = new Button();
		lblMessage = new Label();
		txtPassmoi = new TextBox();
		label8 = new Label();
		txtNhaplaiPass = new TextBox();
		label2 = new Label();
		((Control)this).SuspendLayout();
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 20.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(48, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(505, 31);
		((Control)label1).TabIndex = 0;
		((Control)label1).Text = "ĐỔI MẬT KHẨU NHẬP TL BẰNG TAY";
		((Control)btnExit).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnExit).FlatAppearance.BorderSize = 0;
		((Control)btnExit).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnExit).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnExit).ImageAlign = (ContentAlignment)16;
		((Control)btnExit).Location = new Point(351, 138);
		((Control)btnExit).Name = "btnExit";
		((Control)btnExit).Size = new Size(134, 45);
		((Control)btnExit).TabIndex = 3;
		((Control)btnExit).Text = "        Thoát";
		((ButtonBase)btnExit).UseVisualStyleBackColor = true;
		((Control)btnExit).Click += btnExit_Click;
		((Control)btnSave).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnSave).FlatAppearance.BorderSize = 0;
		((Control)btnSave).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave).Image = (Image)(object)Resources.Save;
		((ButtonBase)btnSave).ImageAlign = (ContentAlignment)16;
		((Control)btnSave).Location = new Point(211, 138);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(134, 45);
		((Control)btnSave).TabIndex = 2;
		((Control)btnSave).Text = "        Lưu";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblMessage).ForeColor = Color.Red;
		((Control)lblMessage).Location = new Point(195, 40);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(74, 20);
		((Control)lblMessage).TabIndex = 115;
		((Control)lblMessage).Text = "Message";
		((Control)txtPassmoi).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtPassmoi).Location = new Point(210, 74);
		((Control)txtPassmoi).Name = "txtPassmoi";
		txtPassmoi.PasswordChar = '*';
		((Control)txtPassmoi).Size = new Size(275, 26);
		((Control)txtPassmoi).TabIndex = 0;
		((Control)label8).AutoSize = true;
		((Control)label8).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label8).Location = new Point(39, 77);
		((Control)label8).Name = "label8";
		((Control)label8).Size = new Size(104, 20);
		((Control)label8).TabIndex = 134;
		((Control)label8).Text = "Mật khẩu mới";
		((Control)txtNhaplaiPass).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtNhaplaiPass).Location = new Point(210, 106);
		((Control)txtNhaplaiPass).Name = "txtNhaplaiPass";
		txtNhaplaiPass.PasswordChar = '*';
		((Control)txtNhaplaiPass).Size = new Size(275, 26);
		((Control)txtNhaplaiPass).TabIndex = 1;
		((Control)label2).AutoSize = true;
		((Control)label2).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label2).Location = new Point(39, 109);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(165, 20);
		((Control)label2).TabIndex = 136;
		((Control)label2).Text = "Nhập lại Mật khẩu mới";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(560, 197);
		((Control)this).Controls.Add((Control)(object)txtNhaplaiPass);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)txtPassmoi);
		((Control)this).Controls.Add((Control)(object)label8);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)btnExit);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmCaidatMatkhau";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Cài đặt Mật khẩu";
		((Form)this).Load += frmCaidatMatkhau_Load;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
