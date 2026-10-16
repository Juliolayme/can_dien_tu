using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmTemplate : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private IContainer components = null;

	private ComboBox cboMau;

	private Label label1;

	private Label label10;

	private Button btnExcel;

	private Button btnThoat;

	private Label lblMessage;

	public frmTemplate()
	{
		InitializeComponent();
	}

	private void btnThoat_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void btnExcel_Click(object sender, EventArgs e)
	{
		try
		{
			string query = "update thongsohethong set maunen=@maunen ";
			dataFetcher.setCommand(query, 1);
			dataFetcher.addParameterWithValue("@maunen", ((Control)cboMau).Text);
			dataFetcher.execute();
			((Control)lblMessage).Text = "Đã lưu!";
		}
		catch
		{
			((Control)lblMessage).Text = "Không lưu được";
		}
	}

	private void frmTemplate_Load(object sender, EventArgs e)
	{
		DBUtil.LoadMaunen((Form)(object)this);
		((Control)lblMessage).Text = "";
		cboMau.Items.Add((object)"Mặc định");
		cboMau.Items.Add((object)"Đỏ");
		cboMau.Items.Add((object)"Vàng");
		cboMau.Items.Add((object)"Xanh");
		cboMau.Items.Add((object)"Trắng");
		cboMau.Items.Add((object)"Cam");
		cboMau.Items.Add((object)"Hồng");
		((Control)cboMau).Text = DBUtil.getValueNoParam("select maunen from thongsohethong");
		if (((Control)cboMau).Text == "")
		{
			((Control)cboMau).Text = "Mặc định";
		}
	}

	private void cboMau_SelectedValueChanged(object sender, EventArgs e)
	{
		try
		{
			switch (((Control)cboMau).Text)
			{
			case "Đỏ":
				((Control)this).BackColor = Color.Red;
				break;
			case "Vàng":
				((Control)this).BackColor = Color.Yellow;
				break;
			case "Xanh":
				((Control)this).BackColor = Color.Green;
				break;
			case "Trắng":
				((Control)this).BackColor = Color.White;
				break;
			case "Cam":
				((Control)this).BackColor = Color.Orange;
				break;
			case "Hồng":
				((Control)this).BackColor = Color.Pink;
				break;
			default:
				((Control)this).BackColor = Color.FromArgb(0, 192, 192);
				break;
			}
		}
		catch
		{
		}
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
		//IL_007a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0084: Expected O, but got Unknown
		//IL_0115: Unknown result type (might be due to invalid IL or missing references)
		//IL_011f: Expected O, but got Unknown
		//IL_019d: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a7: Expected O, but got Unknown
		//IL_0234: Unknown result type (might be due to invalid IL or missing references)
		//IL_023e: Expected O, but got Unknown
		//IL_0312: Unknown result type (might be due to invalid IL or missing references)
		//IL_031c: Expected O, but got Unknown
		//IL_03e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_03eb: Expected O, but got Unknown
		//IL_050b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0515: Expected O, but got Unknown
		//IL_052a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0534: Expected O, but got Unknown
		//IL_0537: Unknown result type (might be due to invalid IL or missing references)
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmTemplate));
		cboMau = new ComboBox();
		label1 = new Label();
		label10 = new Label();
		btnExcel = new Button();
		btnThoat = new Button();
		lblMessage = new Label();
		((Control)this).SuspendLayout();
		cboMau.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboMau).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboMau).FormattingEnabled = true;
		((Control)cboMau).Location = new Point(108, 82);
		((Control)cboMau).Name = "cboMau";
		((Control)cboMau).Size = new Size(253, 28);
		((Control)cboMau).TabIndex = 5;
		((ListControl)cboMau).SelectedValueChanged += cboMau_SelectedValueChanged;
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 20.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(112, 9);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(239, 31);
		((Control)label1).TabIndex = 118;
		((Control)label1).Text = "CHỌN MÀU NỀN";
		((Control)label10).AutoSize = true;
		((Control)label10).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label10).Location = new Point(16, 87);
		((Control)label10).Name = "label10";
		((Control)label10).Size = new Size(82, 20);
		((Control)label10).TabIndex = 119;
		((Control)label10).Text = "Chọn màu";
		((Control)btnExcel).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnExcel).FlatAppearance.BorderSize = 0;
		((Control)btnExcel).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnExcel).Image = (Image)(object)Resources.Save;
		((ButtonBase)btnExcel).ImageAlign = (ContentAlignment)16;
		((Control)btnExcel).Location = new Point(95, 123);
		((Control)btnExcel).Name = "btnExcel";
		((Control)btnExcel).Size = new Size(130, 41);
		((Control)btnExcel).TabIndex = 120;
		((Control)btnExcel).Text = "        Lưu";
		((ButtonBase)btnExcel).UseVisualStyleBackColor = true;
		((Control)btnExcel).Click += btnExcel_Click;
		((Control)btnThoat).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnThoat).FlatAppearance.BorderSize = 0;
		((Control)btnThoat).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnThoat).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnThoat).ImageAlign = (ContentAlignment)16;
		((Control)btnThoat).Location = new Point(231, 123);
		((Control)btnThoat).Name = "btnThoat";
		((Control)btnThoat).Size = new Size(130, 41);
		((Control)btnThoat).TabIndex = 121;
		((Control)btnThoat).Text = "        Thoát";
		((ButtonBase)btnThoat).UseVisualStyleBackColor = true;
		((Control)btnThoat).Click += btnThoat_Click;
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblMessage).Location = new Point(105, 53);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(74, 20);
		((Control)lblMessage).TabIndex = 122;
		((Control)lblMessage).Text = "Message";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(456, 176);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)btnExcel);
		((Control)this).Controls.Add((Control)(object)btnThoat);
		((Control)this).Controls.Add((Control)(object)label10);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)cboMau);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Control)this).Name = "frmTemplate";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Cài đặt giao diện";
		((Form)this).Load += frmTemplate_Load;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
