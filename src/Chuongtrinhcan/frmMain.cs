using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmMain : Form
{
	private IContainer components = null;

	private MenuStrip menuStrip1;

	private ToolStripMenuItem mnuCaidat;

	private ToolStripMenuItem mnuLoaihang;

	private ToolStripMenuItem mnuChangePass;

	private ToolStripMenuItem mnuThuchien;

	private ToolStripMenuItem mnuThongke;

	private ToolStripMenuItem mnuDSNhanvien;

	private ToolStripMenuItem mnuCanhang;

	private ToolStripMenuItem mnuMaunen;

	private ToolStripMenuItem mnuDSKH;

	private ToolStripMenuItem mnuKieuin;

	private ToolStripMenuItem mnuDoipassnhaptay;

	private DataFetcher dataFetcher = DataFetcher.Instance;

	private Class1 cl = new Class1();

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
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_008a: Expected O, but got Unknown
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Expected O, but got Unknown
		//IL_0198: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a2: Expected O, but got Unknown
		//IL_0533: Unknown result type (might be due to invalid IL or missing references)
		//IL_053d: Expected O, but got Unknown
		//IL_06c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d2: Expected O, but got Unknown
		//IL_06df: Unknown result type (might be due to invalid IL or missing references)
		//IL_06e9: Expected O, but got Unknown
		//IL_06ec: Unknown result type (might be due to invalid IL or missing references)
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmMain));
		menuStrip1 = new MenuStrip();
		mnuCaidat = new ToolStripMenuItem();
		mnuLoaihang = new ToolStripMenuItem();
		mnuDSNhanvien = new ToolStripMenuItem();
		mnuDSKH = new ToolStripMenuItem();
		mnuMaunen = new ToolStripMenuItem();
		mnuKieuin = new ToolStripMenuItem();
		mnuDoipassnhaptay = new ToolStripMenuItem();
		mnuChangePass = new ToolStripMenuItem();
		mnuThuchien = new ToolStripMenuItem();
		mnuCanhang = new ToolStripMenuItem();
		mnuThongke = new ToolStripMenuItem();
		((Control)menuStrip1).SuspendLayout();
		((Control)this).SuspendLayout();
		((ToolStrip)menuStrip1).Items.AddRange((ToolStripItem[])(object)new ToolStripItem[2]
		{
			(ToolStripItem)mnuCaidat,
			(ToolStripItem)mnuThuchien
		});
		((Control)menuStrip1).Location = new Point(0, 0);
		((Control)menuStrip1).Name = "menuStrip1";
		((Control)menuStrip1).Size = new Size(1121, 28);
		((Control)menuStrip1).TabIndex = 7;
		((Control)menuStrip1).Text = "menuStrip1";
		((ToolStripDropDownItem)mnuCaidat).DropDownItems.AddRange((ToolStripItem[])(object)new ToolStripItem[7]
		{
			(ToolStripItem)mnuLoaihang,
			(ToolStripItem)mnuDSNhanvien,
			(ToolStripItem)mnuDSKH,
			(ToolStripItem)mnuMaunen,
			(ToolStripItem)mnuKieuin,
			(ToolStripItem)mnuDoipassnhaptay,
			(ToolStripItem)mnuChangePass
		});
		((ToolStripItem)mnuCaidat).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ToolStripItem)mnuCaidat).Image = (Image)(object)Resources.setting_tools_icon;
		((ToolStripItem)mnuCaidat).Name = "mnuCaidat";
		mnuCaidat.ShortcutKeyDisplayString = "Ctrl+M";
		mnuCaidat.ShortcutKeys = (Keys)131149;
		((ToolStripItem)mnuCaidat).Size = new Size(87, 24);
		((ToolStripItem)mnuCaidat).Text = "Cài đặt";
		((ToolStripItem)mnuLoaihang).Image = (Image)(object)Resources.list_edit;
		((ToolStripItem)mnuLoaihang).Name = "mnuLoaihang";
		mnuLoaihang.ShortcutKeyDisplayString = "";
		((ToolStripItem)mnuLoaihang).Size = new Size(259, 24);
		((ToolStripItem)mnuLoaihang).Text = "Danh sách Sản phẩm";
		((ToolStripItem)mnuLoaihang).Click += mnuLoaihang_Click;
		((ToolStripItem)mnuDSNhanvien).Image = (Image)(object)Resources.people;
		((ToolStripItem)mnuDSNhanvien).Name = "mnuDSNhanvien";
		mnuDSNhanvien.ShortcutKeyDisplayString = "";
		((ToolStripItem)mnuDSNhanvien).Size = new Size(259, 24);
		((ToolStripItem)mnuDSNhanvien).Text = "Danh sách Nhân viên";
		((ToolStripItem)mnuDSNhanvien).Click += mnuDSNhanvien_Click;
		((ToolStripItem)mnuDSKH).Image = (Image)(object)Resources.people;
		((ToolStripItem)mnuDSKH).Name = "mnuDSKH";
		((ToolStripItem)mnuDSKH).Size = new Size(259, 24);
		((ToolStripItem)mnuDSKH).Text = "Danh sách Khách hàng";
		((ToolStripItem)mnuDSKH).Click += mnuDSKH_Click;
		((ToolStripItem)mnuMaunen).Image = (Image)(object)Resources.Actions_office_chart_pie_icon;
		((ToolStripItem)mnuMaunen).Name = "mnuMaunen";
		((ToolStripItem)mnuMaunen).Size = new Size(259, 24);
		((ToolStripItem)mnuMaunen).Text = "Màu nền";
		((ToolStripItem)mnuMaunen).Click += mnuMaunen_Click;
		((ToolStripItem)mnuKieuin).Image = (Image)(object)Resources.printer_icon;
		((ToolStripItem)mnuKieuin).Name = "mnuKieuin";
		((ToolStripItem)mnuKieuin).Size = new Size(259, 24);
		((ToolStripItem)mnuKieuin).Text = "Thông tin in phiếu";
		((ToolStripItem)mnuKieuin).Click += mnuKieuin_Click;
		((ToolStripItem)mnuDoipassnhaptay).Image = (Image)(object)Resources.Edit;
		((ToolStripItem)mnuDoipassnhaptay).Name = "mnuDoipassnhaptay";
		((ToolStripItem)mnuDoipassnhaptay).Size = new Size(259, 24);
		((ToolStripItem)mnuDoipassnhaptay).Text = "Đổi mật khẩu nhập TL tay";
		((ToolStripItem)mnuDoipassnhaptay).Click += mnuDoipassnhaptay_Click;
		((ToolStripItem)mnuChangePass).Image = (Image)(object)Resources.password;
		((ToolStripItem)mnuChangePass).Name = "mnuChangePass";
		mnuChangePass.ShortcutKeyDisplayString = "";
		((ToolStripItem)mnuChangePass).Size = new Size(259, 24);
		((ToolStripItem)mnuChangePass).Text = "Đổi mật khẩu";
		((ToolStripItem)mnuChangePass).Click += mnuChangePass_Click;
		((ToolStripDropDownItem)mnuThuchien).DropDownItems.AddRange((ToolStripItem[])(object)new ToolStripItem[2]
		{
			(ToolStripItem)mnuCanhang,
			(ToolStripItem)mnuThongke
		});
		((ToolStripItem)mnuThuchien).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ToolStripItem)mnuThuchien).Image = (Image)(object)Resources.Actions_office_chart_pie_icon;
		((ToolStripItem)mnuThuchien).Name = "mnuThuchien";
		((ToolStripItem)mnuThuchien).Size = new Size(106, 24);
		((ToolStripItem)mnuThuchien).Text = "Thực hiện";
		((ToolStripItem)mnuCanhang).Image = (Image)(object)Resources.Science_Weight_icon;
		((ToolStripItem)mnuCanhang).Name = "mnuCanhang";
		((ToolStripItem)mnuCanhang).Size = new Size(194, 24);
		((ToolStripItem)mnuCanhang).Text = "Cân hàng";
		((ToolStripItem)mnuCanhang).Click += mnuCanhang_Click;
		((ToolStripItem)mnuThongke).Image = (Image)(object)Resources.report_32;
		((ToolStripItem)mnuThongke).Name = "mnuThongke";
		mnuThongke.ShortcutKeyDisplayString = "";
		((ToolStripItem)mnuThongke).Size = new Size(194, 24);
		((ToolStripItem)mnuThongke).Text = "Thống kê dữ liệu";
		((ToolStripItem)mnuThongke).Click += mnuThongke_Click_1;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(1121, 594);
		((Control)this).Controls.Add((Control)(object)menuStrip1);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Control)this).Name = "frmMain";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Chương trình quản lý cân gà - CÔNG TY CÂN ĐIỆN TỬ VIỆT ĐỨC - Hotline: 0936.024.024";
		((Form)this).WindowState = (FormWindowState)2;
		((Form)this).Load += frmMain_Load;
		((Control)this).Resize += frmMain_Resize;
		((Control)menuStrip1).ResumeLayout(false);
		((Control)menuStrip1).PerformLayout();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}

	public frmMain()
	{
		InitializeComponent();
	}

	private void frmMain_Load(object sender, EventArgs e)
	{
		DBUtil.LoadMaunen((Form)(object)this);
		if (Globals.gMaNV == "admin")
		{
			((ToolStripItem)mnuDSNhanvien).Visible = true;
			((ToolStripItem)mnuDoipassnhaptay).Visible = true;
		}
		else
		{
			((ToolStripItem)mnuDSNhanvien).Visible = false;
			((ToolStripItem)mnuDoipassnhaptay).Visible = false;
		}
	}

	private void mnuDSNhanvien_Click(object sender, EventArgs e)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		frmNhanvien frmNhanvien2 = new frmNhanvien();
		((Form)frmNhanvien2).ShowDialog();
	}

	private void mnuChangePass_Click(object sender, EventArgs e)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		frmChangePass frmChangePass2 = new frmChangePass();
		((Form)frmChangePass2).ShowDialog();
	}

	private void mnuLoaihang_Click(object sender, EventArgs e)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		frmHang frmHang2 = new frmHang();
		((Form)frmHang2).ShowDialog();
	}

	private void mnuThongke_Click_1(object sender, EventArgs e)
	{
		try
		{
			frmThongke frmThongke2 = new frmThongke();
			((Control)frmThongke2).Show();
		}
		catch
		{
		}
	}

	private void mnuCanhang_Click(object sender, EventArgs e)
	{
		try
		{
			frmCan frmCan2 = new frmCan();
			((Control)frmCan2).Show();
		}
		catch
		{
		}
	}

	private void mnuMaunen_Click(object sender, EventArgs e)
	{
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			frmTemplate frmTemplate2 = new frmTemplate();
			((Form)frmTemplate2).ShowDialog();
		}
		catch
		{
		}
	}

	private void mnuDSKH_Click(object sender, EventArgs e)
	{
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			frmKhachhang frmKhachhang2 = new frmKhachhang();
			((Form)frmKhachhang2).ShowDialog();
		}
		catch
		{
		}
	}

	private void mnuKieuin_Click(object sender, EventArgs e)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		frmCaidat frmCaidat2 = new frmCaidat();
		((Form)frmCaidat2).ShowDialog();
	}

	private void frmMain_Resize(object sender, EventArgs e)
	{
	}

	private void mnuDoipassnhaptay_Click(object sender, EventArgs e)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		frmCaidatMatkhau frmCaidatMatkhau2 = new frmCaidatMatkhau();
		((Form)frmCaidatMatkhau2).ShowDialog();
	}
}
