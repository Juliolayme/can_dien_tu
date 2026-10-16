using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmCaidat : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private IContainer components = null;

	private Label label1;

	private Label label2;

	private Label label3;

	private Label label4;

	private Label label5;

	private Label label6;

	private Label label7;

	private Label label8;

	private TextBox txtTencty;

	private TextBox txtDiachi;

	private TextBox txtDienthoai;

	private TextBox txtKyten;

	private TextBox txtChieucaogiay;

	private ComboBox cboSotrangin;

	private ComboBox cboKieuin;

	private Button btnExit;

	private Button btnSave;

	private Label lblMessage;

	public frmCaidat()
	{
		InitializeComponent();
	}

	private void btnSave_Click(object sender, EventArgs e)
	{
		try
		{
			string query = "UPDATE giaotiep SET tencty=@tencty,diachi=@diachi,dienthoai=@dienthoai,nguoican=@nguoican,chieucaogiay=@cao,sotrangin=@trang,kieuin=@kieu";
			dataFetcher.setCommand(query, 1);
			dataFetcher.addParameterWithValue("@tencty", ((Control)txtTencty).Text);
			dataFetcher.addParameterWithValue("@diachi", ((Control)txtDiachi).Text);
			dataFetcher.addParameterWithValue("@dienthoai", ((Control)txtDienthoai).Text);
			dataFetcher.addParameterWithValue("@nguoican", ((Control)txtKyten).Text);
			dataFetcher.addParameterWithValue("@cao", ((Control)txtChieucaogiay).Text);
			dataFetcher.addParameterWithValue("@trang", ((Control)cboSotrangin).Text);
			dataFetcher.addParameterWithValue("@kieu", ((Control)cboKieuin).Text);
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
		catch
		{
			((Control)lblMessage).Text = "Dữ liệu nhập không đúng, vui lòng nhập lại";
		}
	}

	private void frmCaidat_Load(object sender, EventArgs e)
	{
		DBUtil.LoadMaunen((Form)(object)this);
		((Control)lblMessage).Text = "";
		try
		{
			((Control)txtTencty).Text = DBUtil.getValue("select tencty from giaotiep where id=@param", "1");
			((Control)txtDiachi).Text = DBUtil.getValue("select diachi from giaotiep where id=@param", "1");
			((Control)txtDienthoai).Text = DBUtil.getValue("select dienthoai from giaotiep where id=@param", "1");
			((Control)txtKyten).Text = DBUtil.getValue("select nguoican from giaotiep where id=@param", "1");
			((Control)txtChieucaogiay).Text = DBUtil.getValue("select chieucaogiay from giaotiep where id=@param", "1");
			((Control)cboSotrangin).Text = DBUtil.getValue("select sotrangin from giaotiep where id=@param", "1");
			((Control)cboKieuin).Text = DBUtil.getValue("select kieuin from giaotiep where id=@param", "1");
		}
		catch
		{
		}
	}

	private void btnExit_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
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
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_008a: Expected O, but got Unknown
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b6: Expected O, but got Unknown
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c1: Expected O, but got Unknown
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cc: Expected O, but got Unknown
		//IL_00cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d7: Expected O, but got Unknown
		//IL_00fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0108: Expected O, but got Unknown
		//IL_01fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0208: Expected O, but got Unknown
		//IL_0282: Unknown result type (might be due to invalid IL or missing references)
		//IL_028c: Expected O, but got Unknown
		//IL_0306: Unknown result type (might be due to invalid IL or missing references)
		//IL_0310: Expected O, but got Unknown
		//IL_045f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0469: Expected O, but got Unknown
		//IL_0537: Unknown result type (might be due to invalid IL or missing references)
		//IL_0541: Expected O, but got Unknown
		//IL_05a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_05ae: Expected O, but got Unknown
		//IL_0611: Unknown result type (might be due to invalid IL or missing references)
		//IL_061b: Expected O, but got Unknown
		//IL_077b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0785: Expected O, but got Unknown
		//IL_0840: Unknown result type (might be due to invalid IL or missing references)
		//IL_084a: Expected O, but got Unknown
		//IL_0924: Unknown result type (might be due to invalid IL or missing references)
		//IL_092e: Expected O, but got Unknown
		//IL_09f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a00: Expected O, but got Unknown
		//IL_0bfb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c05: Expected O, but got Unknown
		//IL_0c1a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c24: Expected O, but got Unknown
		//IL_0c27: Unknown result type (might be due to invalid IL or missing references)
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmCaidat));
		label1 = new Label();
		label2 = new Label();
		label3 = new Label();
		label4 = new Label();
		label5 = new Label();
		label6 = new Label();
		label7 = new Label();
		label8 = new Label();
		txtTencty = new TextBox();
		txtDiachi = new TextBox();
		txtDienthoai = new TextBox();
		txtKyten = new TextBox();
		txtChieucaogiay = new TextBox();
		cboSotrangin = new ComboBox();
		cboKieuin = new ComboBox();
		btnExit = new Button();
		btnSave = new Button();
		lblMessage = new Label();
		((Control)this).SuspendLayout();
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 20.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(315, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(290, 31);
		((Control)label1).TabIndex = 0;
		((Control)label1).Text = "CÀI ĐẶT THÔNG SỐ";
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(10, 319);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(78, 16);
		((Control)label2).TabIndex = 1;
		((Control)label2).Text = "Tên công ty";
		((Control)label2).Visible = false;
		((Control)label3).AutoSize = true;
		((Control)label3).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label3).Location = new Point(10, 76);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(57, 20);
		((Control)label3).TabIndex = 2;
		((Control)label3).Text = "Địa chỉ";
		((Control)label4).AutoSize = true;
		((Control)label4).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label4).Location = new Point(10, 104);
		((Control)label4).Name = "label4";
		((Control)label4).Size = new Size(81, 20);
		((Control)label4).TabIndex = 3;
		((Control)label4).Text = "Điện thoại";
		((Control)label5).AutoSize = true;
		((Control)label5).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label5).Location = new Point(10, 133);
		((Control)label5).Name = "label5";
		((Control)label5).Size = new Size(53, 20);
		((Control)label5).TabIndex = 4;
		((Control)label5).Text = "Ký tên";
		((Control)label6).AutoSize = true;
		((Control)label6).Location = new Point(10, 262);
		((Control)label6).Name = "label6";
		((Control)label6).Size = new Size(122, 16);
		((Control)label6).TabIndex = 5;
		((Control)label6).Text = "Chiều cao khổ giấy";
		((Control)label7).AutoSize = true;
		((Control)label7).Location = new Point(10, 290);
		((Control)label7).Name = "label7";
		((Control)label7).Size = new Size(71, 16);
		((Control)label7).TabIndex = 6;
		((Control)label7).Text = "Số trang in";
		((Control)label8).AutoSize = true;
		((Control)label8).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label8).Location = new Point(11, 164);
		((Control)label8).Name = "label8";
		((Control)label8).Size = new Size(56, 20);
		((Control)label8).TabIndex = 7;
		((Control)label8).Text = "Kiểu in";
		((Control)txtTencty).Location = new Point(146, 313);
		((Control)txtTencty).Name = "txtTencty";
		((Control)txtTencty).Size = new Size(520, 22);
		((Control)txtTencty).TabIndex = 8;
		((Control)txtTencty).Visible = false;
		((Control)txtDiachi).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtDiachi).Location = new Point(146, 70);
		((Control)txtDiachi).Name = "txtDiachi";
		((Control)txtDiachi).Size = new Size(783, 26);
		((Control)txtDiachi).TabIndex = 9;
		((Control)txtDienthoai).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtDienthoai).Location = new Point(146, 98);
		((Control)txtDienthoai).Name = "txtDienthoai";
		((Control)txtDienthoai).Size = new Size(783, 26);
		((Control)txtDienthoai).TabIndex = 10;
		((Control)txtKyten).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtKyten).Location = new Point(146, 127);
		((Control)txtKyten).Name = "txtKyten";
		((Control)txtKyten).Size = new Size(783, 26);
		((Control)txtKyten).TabIndex = 11;
		((Control)txtChieucaogiay).Location = new Point(146, 256);
		((Control)txtChieucaogiay).Name = "txtChieucaogiay";
		((Control)txtChieucaogiay).Size = new Size(196, 22);
		((Control)txtChieucaogiay).TabIndex = 12;
		((Control)txtChieucaogiay).Text = "6";
		cboSotrangin.DropDownStyle = (ComboBoxStyle)2;
		((ListControl)cboSotrangin).FormattingEnabled = true;
		cboSotrangin.Items.AddRange(new object[1] { "1" });
		((Control)cboSotrangin).Location = new Point(146, 282);
		((Control)cboSotrangin).Name = "cboSotrangin";
		((Control)cboSotrangin).Size = new Size(196, 24);
		((Control)cboSotrangin).TabIndex = 13;
		cboKieuin.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboKieuin).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboKieuin).FormattingEnabled = true;
		cboKieuin.Items.AddRange(new object[2] { "Bằng tay", "Tự động" });
		((Control)cboKieuin).Location = new Point(146, 156);
		((Control)cboKieuin).Name = "cboKieuin";
		((Control)cboKieuin).Size = new Size(196, 28);
		((Control)cboKieuin).TabIndex = 14;
		((Control)btnExit).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnExit).FlatAppearance.BorderSize = 0;
		((Control)btnExit).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnExit).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnExit).ImageAlign = (ContentAlignment)16;
		((Control)btnExit).Location = new Point(468, 190);
		((Control)btnExit).Name = "btnExit";
		((Control)btnExit).Size = new Size(137, 41);
		((Control)btnExit).TabIndex = 114;
		((Control)btnExit).Text = "        Thoát";
		((ButtonBase)btnExit).UseVisualStyleBackColor = true;
		((Control)btnExit).Click += btnExit_Click;
		((Control)btnSave).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnSave).FlatAppearance.BorderSize = 0;
		((Control)btnSave).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave).Image = (Image)(object)Resources.Save;
		((ButtonBase)btnSave).ImageAlign = (ContentAlignment)16;
		((Control)btnSave).Location = new Point(325, 190);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(137, 41);
		((Control)btnSave).TabIndex = 113;
		((Control)btnSave).Text = "        Lưu";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblMessage).Location = new Point(143, 51);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(74, 20);
		((Control)lblMessage).TabIndex = 115;
		((Control)lblMessage).Text = "Message";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(941, 235);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)btnExit);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Controls.Add((Control)(object)cboKieuin);
		((Control)this).Controls.Add((Control)(object)cboSotrangin);
		((Control)this).Controls.Add((Control)(object)txtChieucaogiay);
		((Control)this).Controls.Add((Control)(object)txtKyten);
		((Control)this).Controls.Add((Control)(object)txtDienthoai);
		((Control)this).Controls.Add((Control)(object)txtDiachi);
		((Control)this).Controls.Add((Control)(object)txtTencty);
		((Control)this).Controls.Add((Control)(object)label8);
		((Control)this).Controls.Add((Control)(object)label7);
		((Control)this).Controls.Add((Control)(object)label6);
		((Control)this).Controls.Add((Control)(object)label5);
		((Control)this).Controls.Add((Control)(object)label4);
		((Control)this).Controls.Add((Control)(object)label3);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmCaidat";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Cài đặt thông số";
		((Form)this).Load += frmCaidat_Load;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
