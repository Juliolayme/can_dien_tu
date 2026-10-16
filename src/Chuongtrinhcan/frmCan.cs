using System;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Drawing;
using System.IO.Ports;
using System.Media;
using System.Threading;
using System.Windows.Forms;
using Label = System.Windows.Forms.Label;
using TextBox = System.Windows.Forms.TextBox;
using CheckBox = System.Windows.Forms.CheckBox;
using Timer = System.Windows.Forms.Timer;
using Padding = System.Windows.Forms.Padding;
using Chuongtrinhcan.Properties;
using DataDynamics.ActiveReports;

namespace Chuongtrinhcan;

public class frmCan : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private string data;

	private string data2;

	private double gSolancan;

	private double gMaxIDPhieucan;

	private int flagSaveCan;

	private int counterOndinhCan;

	private int flagSaveCan2;

	private int counterOndinhCan2;

	private double gTLOndinhCan;

	private double gTLOndinhCan2;

	private int gFlagLoadOK = 0;

	private Class1 cl = new Class1();

	private IContainer components = null;

	private Button btnSave;

	private DataGridView dataGrid;

	private Label lblDaucan;

	private TextBox txtSoluong;

	private SerialPort serialPort1;

	private Timer timer1;

	private Timer timerChekOndinh;

	private ComboBox cboMahang;

	private Label label1;

	private ComboBox cboTenhang;

	private ComboBox cboTenKH;

	private Label label2;

	private ComboBox cboMaKH;

	private Label label3;

	private TextBox txtNguoinhan;

	private Label label4;

	private ComboBox cboSoxe;

	private ComboBox cboXuatnhap;

	private Label label5;

	private Button btnKetthuc;

	private Label label6;

	private Label label7;

	private Label lblTotal;

	private Label lblSolancan;

	private Button btnDelete;

	private DataGridView dataGridPhieucan;

	private Button btnTaophieu;

	private GroupBox groupBox1;

	private Label label10;

	private DateTimePicker txtStartDate;

	private DateTimePicker txtEndDate;

	private Label label9;

	private Panel panel1;

	private Button btnInphieu;

	private Label label8;

	private Label lblSophieu;

	private TextBox txtSolong;

	private Label label11;

	private Panel panel2;

	private CheckBox chkCanbi;

	private TextBox txtSocon1long;

	private Label label12;

	private DataGridViewTextBoxColumn stt;

	private DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;

	private DataGridViewTextBoxColumn sophieu;

	private DataGridViewTextBoxColumn soxe;

	private DataGridViewTextBoxColumn tenkh;

	private DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;

	private DataGridViewTextBoxColumn xuatnhap;

	private DataGridViewTextBoxColumn ghichu;

	private Label lblTongsolong;

	private Label label14;

	private Label lblTongsocon;

	private Label label16;

	private Label lblTLTrungbinh;

	private Label label15;

	private DataGridView dataGridTemp;

	private DataGridViewTextBoxColumn macan1;

	private DataGridViewTextBoxColumn tl1;

	private DataGridViewTextBoxColumn solong1;

	private DataGridViewTextBoxColumn socon1;

	private DataGridViewTextBoxColumn macan2;

	private DataGridViewTextBoxColumn tl2;

	private DataGridViewTextBoxColumn solong2;

	private DataGridViewTextBoxColumn socon2;

	private DataGridViewTextBoxColumn macan3;

	private DataGridViewTextBoxColumn tl3;

	private DataGridViewTextBoxColumn solong3;

	private DataGridViewTextBoxColumn socon3;

	private DataGridViewTextBoxColumn macan4;

	private DataGridViewTextBoxColumn tl4;

	private DataGridViewTextBoxColumn solong4;

	private DataGridViewTextBoxColumn socon4;

	private Label lblOndinhKytu;

	private GroupBox groupBox2;

	private Label label13;

	private TextBox txtTLLuutu;

	private Label label17;

	private TextBox txtTGOndinh;

	private CheckBox chkCantudong;

	private Label lblDaucan2;

	private TextBox txtSoluong2;

	private Label label18;

	private ComboBox cboCongCom1;

	private GroupBox groupBox3;

	private Label label19;

	private ComboBox cboCongCom2;

	private Label label20;

	private TextBox txtTLLuutu2;

	private Label label21;

	private TextBox txtTGOndinh2;

	private Timer timerChekOndinh2;

	private Timer timer2;

	private Button btnSave2;

	private CheckBox chkCantudong2;

	private TextBox txtSolongbi;

	private Label label22;

	private Label lblSolancanbi;

	private Label label24;

	private Label lblSolongcanbi;

	private Label label25;

	private SerialPort serialPort2;

	private CheckBox chkCanbi2;

	private DataGridViewTextBoxColumn id;

	private DataGridViewTextBoxColumn STT2;

	private DataGridViewTextBoxColumn mathang;

	private DataGridViewTextBoxColumn soluong;

	private DataGridViewTextBoxColumn solong;

	private DataGridViewTextBoxColumn socon1long;

	private DataGridViewTextBoxColumn tongsocon;

	private DataGridViewTextBoxColumn hinhthuccan;

	private DataGridViewTextBoxColumn ngaycan;

	public frmCan()
	{
		InitializeComponent();
	}

	public bool IsNumeric(string text)
	{
		double result;
		return double.TryParse(text, out result);
	}

	private void formatDataGrid()
	{
		for (int i = 0; i < ((BaseCollection)dataGrid.Columns).Count; i++)
		{
			((DataGridViewCell)dataGrid.Columns[i].HeaderCell).Style.Alignment = (DataGridViewContentAlignment)32;
		}
	}

	private void frmCan_Resize(object sender, EventArgs e)
	{
		try
		{
			((Control)dataGrid).Height = ((Control)this).Height - 182;
			((Control)dataGrid).Width = ((Control)this).Width - 680;
			((Control)dataGridPhieucan).Height = ((Control)this).Height - ((Control)panel1).Height - 153;
		}
		catch
		{
		}
	}

	private void frmCan_Load(object sender, EventArgs e)
	{
		//IL_01f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0203: Expected O, but got Unknown
		DBUtil.LoadMaunen((Form)(object)this);
		txtStartDate.Value = DateTime.Now;
		txtStartDate.CustomFormat = "dd/MM/yyyy";
		txtEndDate.Value = DateTime.Now;
		txtEndDate.CustomFormat = "dd/MM/yyyy";
		gSolancan = 0.0;
		string text = "";
		text = "select id,mahang as fullname from hanghoa order by mahang";
		DBUtil.loadComboValue(cboMahang, text, "fullname", "id");
		text = "select id,tenhang as fullname from hanghoa order by tenhang";
		DBUtil.loadComboValue(cboTenhang, text, "fullname", "id");
		text = "select id,makh as fullname from khachhang order by makh";
		DBUtil.loadComboValue(cboMaKH, text, "fullname", "id");
		text = "select id,tenkh as fullname from khachhang order by tenkh";
		DBUtil.loadComboValue(cboTenKH, text, "fullname", "id");
		text = "select distinct soxe from phieucan where soxe <>'' order by soxe";
		DBUtil.loadCombo(cboSoxe, text, "soxe");
		((ListControl)cboMahang).SelectedIndex = -1;
		((ListControl)cboTenhang).SelectedIndex = -1;
		((ListControl)cboMaKH).SelectedIndex = -1;
		((ListControl)cboTenKH).SelectedIndex = -1;
		((Control)cboXuatnhap).Text = "";
		((Control)cboSoxe).Text = "";
		((Control)txtNguoinhan).Text = "";
		showDataPhieuchuacan();
		((Control)txtTGOndinh).Text = DBUtil.getValueNoParam("select tgondinh from thongsohethong where id='1'");
		((Control)txtTLLuutu).Text = DBUtil.getValueNoParam("select tlluutu from thongsohethong where id='1'");
		((Control)txtTGOndinh2).Text = DBUtil.getValueNoParam("select tgondinh from thongsohethong where id='2'");
		((Control)txtTLLuutu2).Text = DBUtil.getValueNoParam("select tlluutu from thongsohethong where id='2'");
		string text2 = (((Control)cboCongCom1).Text = DBUtil.getValueNoParam("select congcom from thongsohethong where id='1'"));
		string text3 = text2;
		text2 = (((Control)cboCongCom2).Text = DBUtil.getValueNoParam("select congcom from thongsohethong where id='2'"));
		string text4 = text2;
		gFlagLoadOK = 1;
		try
		{
			serialPort1 = new SerialPort(text3, 9600, (Parity)0, 8, (StopBits)1);
			serialPort1.ReceivedBytesThreshold = 1;
			if (!serialPort1.IsOpen)
			{
				serialPort1.Open();
			}
			timer1.Enabled = true;
		}
		catch
		{
		}
	}

	private void btnSave_Click(object sender, EventArgs e)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		if (!IsNumeric(((Control)lblDaucan).Text))
		{
			MessageBox.Show("Dữ liệu cân không đúng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			return;
		}
		if (((Control)cboMahang).Text == "")
		{
			MessageBox.Show("Vui lòng nhập loại hàng để cân.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			((Control)cboMahang).Focus();
			return;
		}
		if (chkCanbi.Checked)
		{
			insertDataPhieucanChitietCanbi();
		}
		else
		{
			insertDataPhieucanChitiet();
		}
		Thread.Sleep(700);
		showData();
	}

	private void btnSave2_Click(object sender, EventArgs e)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		if (!IsNumeric(((Control)lblDaucan2).Text))
		{
			MessageBox.Show("Dữ liệu cân không đúng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
		}
		else if (((Control)cboMahang).Text == "")
		{
			MessageBox.Show("Vui lòng nhập loại hàng để cân.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			((Control)cboMahang).Focus();
		}
		else
		{
			insertDataPhieucanChitietCanbi();
			Thread.Sleep(700);
			showData();
		}
	}

	private void timer1_Tick(object sender, EventArgs e)
	{
		try
		{
			string text = "";
			data += serialPort1.ReadExisting();
			if (data.Length >= 20)
			{
				text = ReadScale(data).ToString();
				((Control)lblDaucan).Text = double.Parse(text).ToString();
				data = "";
			}
		}
		catch
		{
			data = "";
		}
	}

	private void timer2_Tick(object sender, EventArgs e)
	{
		try
		{
			string text = "";
			data2 += serialPort2.ReadExisting();
			if (data2.Length >= 20)
			{
				text = ReadScale2(data2).ToString();
				((Control)lblDaucan2).Text = double.Parse(text).ToString();
				data2 = "";
			}
		}
		catch
		{
			data2 = "";
		}
	}

	public string Reverse(string text)
	{
		if (text == null)
		{
			return null;
		}
		char[] array = text.ToCharArray();
		Array.Reverse((Array)array);
		return new string(array);
	}

	private double ReadScale(string data)
	{
		double num = 0.0;
		string value = "=";
		int num2 = 0;
		num2 = data.IndexOf(value, 0);
		return double.Parse(data.Substring(num2 + 1, 7).Replace(" ", ""));
	}

	private double ReadScale2(string data2)
	{
		double num = 0.0;
		string value = ">";
		int num2 = 0;
		num2 = data2.IndexOf(value, 0);
		num = double.Parse(data2.Substring(num2 + 3, 7).Replace(" ", "").Replace(".", ""));
		return num / 10.0;
	}

	private bool inFmtSoluong = false;

	private bool inFmtSoluong2 = false;

	// Nhap lien tuc cac chu so, chu so cuoi luon la phan le (1 so le),
	// dau cham duoc tu dong chen truoc chu so cuoi. VD: 125 -> 12.5
	private string FormatCanTay(string input)
	{
		string d = "";
		if (input != null)
		{
			foreach (char c in input)
			{
				if (c >= '0' && c <= '9')
				{
					d += c;
				}
			}
		}
		int i = 0;
		while (i < d.Length - 1 && d[i] == '0')
		{
			i++;
		}
		d = d.Substring(i);
		if (d.Length == 0)
		{
			return "0";
		}
		if (d.Length == 1)
		{
			return "0." + d;
		}
		return d.Substring(0, d.Length - 1) + "." + d.Substring(d.Length - 1);
	}

	private void txtSoluong_TextChanged(object sender, EventArgs e)
	{
		if (inFmtSoluong)
		{
			return;
		}
		inFmtSoluong = true;
		try
		{
			string s = FormatCanTay(((Control)txtSoluong).Text);
			((Control)txtSoluong).Text = s;
			txtSoluong.SelectionStart = s.Length;
			((Control)lblDaucan).Text = s;
		}
		finally
		{
			inFmtSoluong = false;
		}
	}

	private void txtSoluong2_TextChanged(object sender, EventArgs e)
	{
		if (inFmtSoluong2)
		{
			return;
		}
		inFmtSoluong2 = true;
		try
		{
			string s = FormatCanTay(((Control)txtSoluong2).Text);
			((Control)txtSoluong2).Text = s;
			txtSoluong2.SelectionStart = s.Length;
			((Control)lblDaucan2).Text = s;
		}
		finally
		{
			inFmtSoluong2 = false;
		}
	}

	// Xoa het o nhap can de nguoi can go so moi ngay lap tuc
	private void XoaONhapCan()
	{
		inFmtSoluong = true;
		((Control)txtSoluong).Text = "";
		inFmtSoluong = false;
		inFmtSoluong2 = true;
		((Control)txtSoluong2).Text = "";
		inFmtSoluong2 = false;
		((Control)lblDaucan).Text = "0";
		((Control)lblDaucan2).Text = "0";
		if (((Control)txtSoluong).Visible)
		{
			((Control)txtSoluong).Focus();
		}
		else if (((Control)txtSoluong2).Visible)
		{
			((Control)txtSoluong2).Focus();
		}
	}

	private void button1_Click(object sender, EventArgs e)
	{
		serialPort1.DtrEnable = true;
	}

	private void button2_Click(object sender, EventArgs e)
	{
		serialPort1.DtrEnable = false;
	}

	private void cboMahang_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			((ListControl)cboTenhang).SelectedValue = ((ListControl)cboMahang).SelectedValue;
		}
		catch
		{
		}
	}

	private void cboTenhang_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			((ListControl)cboMahang).SelectedValue = ((ListControl)cboTenhang).SelectedValue;
		}
		catch
		{
		}
	}

	private void cboMaKH_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (gFlagLoadOK == 0)
		{
			return;
		}
		try
		{
			((ListControl)cboTenKH).SelectedValue = ((ListControl)cboMaKH).SelectedValue;
			try
			{
				if (!(gMaxIDPhieucan <= 0.0))
				{
					capNhatPhieucan();
					Thread.Sleep(700);
					showDataPhieuchuacan();
				}
			}
			catch
			{
			}
		}
		catch
		{
		}
	}

	private void cboTenKH_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (gFlagLoadOK == 0)
		{
			return;
		}
		try
		{
			((ListControl)cboMaKH).SelectedValue = ((ListControl)cboTenKH).SelectedValue;
			try
			{
				if (!(gMaxIDPhieucan <= 0.0))
				{
					capNhatPhieucan();
					Thread.Sleep(700);
					showDataPhieuchuacan();
				}
			}
			catch
			{
			}
		}
		catch
		{
		}
	}

	private void cboMaKH_Click(object sender, EventArgs e)
	{
	}

	private void cboTenKH_Click(object sender, EventArgs e)
	{
	}

	private void btnKetthuc_Click(object sender, EventArgs e)
	{
		try
		{
			gSolancan = 0.0;
			gMaxIDPhieucan = 0.0;
			((Control)cboSoxe).Text = "";
			((ListControl)cboMaKH).SelectedIndex = -1;
			((ListControl)cboTenKH).SelectedIndex = -1;
			((Control)cboXuatnhap).Text = "";
			((Control)txtNguoinhan).Text = "";
			((ListControl)cboMahang).SelectedIndex = -1;
			((ListControl)cboTenhang).SelectedIndex = -1;
			showData();
			((Control)btnTaophieu).Enabled = true;
			((Control)btnKetthuc).Enabled = false;
			((Control)btnInphieu).Enabled = false;
			((Control)btnSave).Enabled = false;
			((Control)btnSave2).Enabled = false;
		}
		catch
		{
		}
	}

	private void frmCan_KeyDown(object sender, KeyEventArgs e)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		switch (((object)e.KeyCode).ToString())
		{
		case "F1":
			btnSave.PerformClick();
			((Control)btnSave).Focus();
			break;
		case "F2":
			btnSave2.PerformClick();
			((Control)btnSave2).Focus();
			break;
		case "Escape":
			XoaONhapCan();
			e.Handled = true;
			e.SuppressKeyPress = true;
			break;
		}
	}

	private void frmCan_FormClosed(object sender, FormClosedEventArgs e)
	{
		try
		{
			serialPort1.Close();
			serialPort2.Close();
			string query = "update thongsohethong set tgondinh=@tg,tlluutu=@tl,congcom=@congcom where id=@id ";
			dataFetcher.setCommand(query, 1);
			dataFetcher.addParameterWithValue("@tg", ((Control)txtTGOndinh).Text);
			dataFetcher.addParameterWithValue("@tl", ((Control)txtTLLuutu).Text);
			dataFetcher.addParameterWithValue("@congcom", ((Control)cboCongCom1).Text);
			dataFetcher.addParameterWithValue("@id", "1");
			int num = dataFetcher.execute();
			Thread.Sleep(1000);
			query = "update thongsohethong set tgondinh=@tg,tlluutu=@tl,congcom=@congcom where id=@id ";
			dataFetcher.setCommand(query, 1);
			dataFetcher.addParameterWithValue("@tg", ((Control)txtTGOndinh2).Text);
			dataFetcher.addParameterWithValue("@tl", ((Control)txtTLLuutu2).Text);
			dataFetcher.addParameterWithValue("@congcom", ((Control)cboCongCom2).Text);
			dataFetcher.addParameterWithValue("@id", "2");
			num = dataFetcher.execute();
		}
		catch
		{
		}
	}

	private void btnDelete_Click(object sender, EventArgs e)
	{
		//IL_00e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0100: Unknown result type (might be due to invalid IL or missing references)
		//IL_005b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Invalid comparison between Unknown and I4
		try
		{
			if (dataGrid.Rows.Count > 0 && (int)MessageBox.Show("Bạn có muốn lượt cân có số thứ tự [" + dataGrid.CurrentRow.Cells["stt2"].Value.ToString() + "] không?", "Thông báo", (MessageBoxButtons)4, (MessageBoxIcon)32) == 6)
			{
				try
				{
					string text = "";
					text = "DELETE * FROM chitietphieucan WHERE id =" + dataGrid.CurrentRow.Cells["id"].Value.ToString();
					dataFetcher.setCommand(text, 1);
					int num = dataFetcher.execute();
					Thread.Sleep(700);
					showData();
					return;
				}
				catch
				{
					MessageBox.Show("Không xóa được.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
					return;
				}
			}
		}
		catch
		{
			MessageBox.Show("Chọn lượt cân cần xóa.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
		}
	}

	private void btnTaophieu_Click(object sender, EventArgs e)
	{
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ad: Unknown result type (might be due to invalid IL or missing references)
		if (((Control)cboSoxe).Text == "")
		{
			MessageBox.Show("Vui lòng nhập số xe.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			((Control)cboSoxe).Focus();
			return;
		}
		if (((Control)cboMaKH).Text == "")
		{
			MessageBox.Show("Vui lòng nhập Khách hàng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			((Control)cboMaKH).Focus();
			return;
		}
		if (((Control)cboMahang).Text == "")
		{
			MessageBox.Show("Vui lòng nhập Loại hàng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			((Control)cboMahang).Focus();
			return;
		}
		((Control)lblSophieu).Text = LaySophieu();
		taoPhieucan();
		Thread.Sleep(700);
		gMaxIDPhieucan = double.Parse(DBUtil.getValueNoParam("select max(stt) from phieucan"));
		showDataPhieuchuacan();
		((Control)btnTaophieu).Enabled = false;
		((Control)btnKetthuc).Enabled = true;
		((Control)btnInphieu).Enabled = true;
		((Control)btnSave).Enabled = true;
		((Control)btnSave2).Enabled = true;
	}

	private string LaySophieu()
	{
		string text = "";
		string text2 = "";
		DateTime now = DateTime.Now;
		text = "select top 1 sophieu from phieucan ";
		object obj = text;
		text = string.Concat(obj, " where day(ngaycan)= ", now.Day, " and month(ngaycan)= ", now.Month, " and year(ngaycan)= ", now.Year);
		text += " order by stt desc";
		try
		{
			text2 = DBUtil.getValueNoParam(text);
			return (double.Parse(text2) + 1.0).ToString();
		}
		catch
		{
			return "1";
		}
	}

	private void taoPhieucan()
	{
		//IL_00c6: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string query = "INSERT INTO PHIEUCAN(SOPHIEU,SOXE,KH_ID,NV_ID,GHICHU) VALUES(@SOPHIEU,@SOXE,@KH_ID,@NV_ID,@GHICHU)";
			dataFetcher.setCommand(query, 1);
			dataFetcher.addParameterWithValue("@SOPHIEU", ((Control)lblSophieu).Text);
			dataFetcher.addParameterWithValue("@SOXE", ((Control)cboSoxe).Text);
			dataFetcher.addParameterWithValue("@KH_ID", ((ListControl)cboMaKH).SelectedValue.ToString());
			dataFetcher.addParameterWithValue("@NV_ID", Globals.gMaNV);
			dataFetcher.addParameterWithValue("@GHICHU", ((Control)txtNguoinhan).Text);
			int num = dataFetcher.execute();
			if (num == -1)
			{
				MessageBox.Show("Không tạo được phiếu.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			}
		}
		catch
		{
		}
	}

	private void capNhatPhieucan()
	{
		//IL_00c6: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string query = "update PHIEUCAN SET SOXE=@SOXE, KH_ID=@KH_ID,NV_ID=@NV_ID,GHICHU=@GHICHU WHERE stt=@stt";
			dataFetcher.setCommand(query, 1);
			dataFetcher.addParameterWithValue("@SOXE", ((Control)cboSoxe).Text);
			dataFetcher.addParameterWithValue("@KH_ID", ((ListControl)cboMaKH).SelectedValue.ToString());
			dataFetcher.addParameterWithValue("@NV_ID", Globals.gMaNV);
			dataFetcher.addParameterWithValue("@GHICHU", ((Control)txtNguoinhan).Text);
			dataFetcher.addParameterWithValue("@stt", gMaxIDPhieucan.ToString());
			int num = dataFetcher.execute();
			if (num == -1)
			{
				MessageBox.Show("Không cập nhật được phiếu.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			}
		}
		catch
		{
		}
	}

	private void insertDataPhieucanChitiet()
	{
		//IL_0139: Unknown result type (might be due to invalid IL or missing references)
		string query = "INSERT INTO chitietphieucan(phieu_id,trongluong,hang_id,solong,socon1long) VALUES(@phieu_id,@trongluong,@hang_id,@solong,@socon1long)";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@phieu_id", gMaxIDPhieucan.ToString());
		if (chkCanbi.Checked)
		{
			dataFetcher.addParameterWithValue("@trongluong", (double.Parse(((Control)lblDaucan).Text) * -1.0).ToString());
		}
		else
		{
			dataFetcher.addParameterWithValue("@trongluong", double.Parse(((Control)lblDaucan).Text).ToString());
		}
		dataFetcher.addParameterWithValue("@hang_id", ((ListControl)cboMahang).SelectedValue.ToString());
		dataFetcher.addParameterWithValue("@solong", double.Parse(((Control)txtSolong).Text).ToString());
		dataFetcher.addParameterWithValue("@socon1long", double.Parse(((Control)txtSocon1long).Text).ToString());
		int num = dataFetcher.execute();
		if (num == -1)
		{
			MessageBox.Show("Không lưu được.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			return;
		}
		playSound("beep-02.wav");
		serialPort1.DtrEnable = true;
		Thread.Sleep(400);
		serialPort1.DtrEnable = false;
	}

	private void insertDataPhieucanChitietCanbi()
	{
		//IL_013c: Unknown result type (might be due to invalid IL or missing references)
		string query = "INSERT INTO chitietphieucan(phieu_id,trongluong,hang_id,solong,socon1long,loaican) VALUES(@phieu_id,@trongluong,@hang_id,@solong,@socon1long,@loaican)";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@phieu_id", gMaxIDPhieucan.ToString());
		if (chkCanbi2.Checked)
		{
			dataFetcher.addParameterWithValue("@trongluong", (double.Parse(((Control)lblDaucan).Text) * -1.0).ToString());
		}
		else
		{
			dataFetcher.addParameterWithValue("@trongluong", double.Parse(((Control)lblDaucan).Text).ToString());
		}
		dataFetcher.addParameterWithValue("@hang_id", ((ListControl)cboMahang).SelectedValue.ToString());
		dataFetcher.addParameterWithValue("@solong", double.Parse(((Control)txtSolongbi).Text).ToString());
		dataFetcher.addParameterWithValue("@socon1long", "0");
		dataFetcher.addParameterWithValue("@loaican", "1");
		int num = dataFetcher.execute();
		if (num == -1)
		{
			MessageBox.Show("Không lưu được.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			return;
		}
		playSound("beep-05.wav");
		serialPort2.DtrEnable = true;
		Thread.Sleep(400);
		serialPort2.DtrEnable = false;
	}

	private void playSound(string path)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		SoundPlayer val = new SoundPlayer();
		val.SoundLocation = path;
		val.Load();
		val.Play();
	}

	private void showDataPhieuchuacan()
	{
		try
		{
			dataGridPhieucan.RowCount = 0;
			string text = "";
			text = "select stt,sophieu,soxe,tenkh,xuatnhap,ghichu,ngaycan from phieucan p, khachhang k ";
			text += " where p.kh_id=k.id ";
			string text2 = text;
			text = text2 + " and ngaycan >=#" + txtStartDate.Value.ToString("MM/dd/yyyy") + " 00:00:01# and ngaycan <= #" + txtEndDate.Value.ToString("MM/dd/yyyy") + " 23:59:59# ";
			text += " order by ngaycan desc ";
			dataFetcher.setCommand(text, 1);
			DataSet dataSet = dataFetcher.fetchDataSet();
			int num = 0;
			foreach (DataRow row in dataSet.Tables[0].Rows)
			{
				num = dataGridPhieucan.Rows.Add();
				dataGridPhieucan.Rows[num].Cells[0].Value = row["stt"];
				dataGridPhieucan.Rows[num].Cells[1].Value = (num + 1).ToString();
				dataGridPhieucan.Rows[num].Cells[2].Value = row["sophieu"];
				dataGridPhieucan.Rows[num].Cells[3].Value = row["soxe"];
				dataGridPhieucan.Rows[num].Cells[4].Value = row["tenkh"];
				dataGridPhieucan.Rows[num].Cells[5].Value = row["ngaycan"];
				dataGridPhieucan.Rows[num].Cells[6].Value = row["xuatnhap"];
				dataGridPhieucan.Rows[num].Cells[7].Value = row["ghichu"];
			}
		}
		catch
		{
		}
	}

	private void showData()
	{
		try
		{
			dataGrid.RowCount = 0;
			string text = "";
			text = "select c.id,trongluong,ngaycanct,tenhang,solong,socon1long,solong*socon1long as tongsocon, ";
			text += " iif(loaican=0,'Cân gà','Cân bì') as hinhthuccan,loaican ";
			text += " from chitietphieucan c, hanghoa h  ";
			text = text + "where h.id=c.hang_id and phieu_id=" + gMaxIDPhieucan;
			text += " order by ngaycanct ";
			dataFetcher.setCommand(text, 1);
			DataSet dataSet = dataFetcher.fetchDataSet();
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			double num4 = 0.0;
			double num5 = 0.0;
			double num6 = 0.0;
			double num7 = 0.0;
			double num8 = 0.0;
			foreach (DataRow row in dataSet.Tables[0].Rows)
			{
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = row["id"];
				dataGrid.Rows[num].Cells[1].Value = (num + 1).ToString();
				dataGrid.Rows[num].Cells[2].Value = row["tenhang"];
				dataGrid.Rows[num].Cells[3].Value = row["trongluong"];
				dataGrid.Rows[num].Cells[4].Value = row["solong"];
				dataGrid.Rows[num].Cells[5].Value = row["socon1long"];
				dataGrid.Rows[num].Cells[6].Value = row["tongsocon"];
				dataGrid.Rows[num].Cells[7].Value = row["hinhthuccan"];
				dataGrid.Rows[num].Cells[8].Value = row["ngaycanct"];
				if (row["loaican"].ToString() == "0")
				{
					num4 += double.Parse(row["trongluong"].ToString());
					num6 += double.Parse(row["solong"].ToString());
					num8 += double.Parse(row["tongsocon"].ToString());
					num2++;
				}
				else
				{
					num5 += double.Parse(row["trongluong"].ToString());
					num7 += double.Parse(row["solong"].ToString());
					num3++;
				}
			}
			dataGrid.CurrentCell = dataGrid.Rows[num].Cells["mathang"];
			((Control)lblTotal).Text = (num4 - num5).ToString("#,##0.0");
			((Control)lblSolancan).Text = num2.ToString();
			((Control)lblSolancanbi).Text = num3.ToString();
			((Control)lblTongsocon).Text = num8.ToString("#,##0");
			((Control)lblTongsolong).Text = num6.ToString("#,##0");
			((Control)lblSolongcanbi).Text = num7.ToString("#,##0");
			((Control)lblTLTrungbinh).Text = ((num4 - num5 / num7 * num6) / num8).ToString("#,##0.000");
		}
		catch
		{
		}
		formatDataGrid();
		dataGrid.AutoResizeColumns();
	}

	private void dataGridPhieucan_DoubleClick(object sender, EventArgs e)
	{
		try
		{
			gFlagLoadOK = 0;
			gMaxIDPhieucan = double.Parse(dataGridPhieucan.CurrentRow.Cells["stt"].Value.ToString());
			((Control)lblSophieu).Text = dataGridPhieucan.CurrentRow.Cells["sophieu"].Value.ToString();
			((Control)cboSoxe).Text = dataGridPhieucan.CurrentRow.Cells["soxe"].Value.ToString();
			((Control)cboTenKH).Text = dataGridPhieucan.CurrentRow.Cells["tenkh"].Value.ToString();
			((ListControl)cboMaKH).SelectedValue = ((ListControl)cboTenKH).SelectedValue;
			((Control)cboXuatnhap).Text = dataGridPhieucan.CurrentRow.Cells["xuatnhap"].Value.ToString();
			((Control)txtNguoinhan).Text = dataGridPhieucan.CurrentRow.Cells["ghichu"].Value.ToString();
			showData();
			((Control)btnTaophieu).Enabled = false;
			((Control)btnKetthuc).Enabled = true;
			((Control)btnInphieu).Enabled = true;
			((Control)btnSave).Enabled = true;
			((Control)btnSave2).Enabled = true;
			gFlagLoadOK = 1;
		}
		catch
		{
		}
	}

	private void txtStartDate_ValueChanged(object sender, EventArgs e)
	{
		if (gFlagLoadOK != 0)
		{
			showDataPhieuchuacan();
		}
	}

	private void txtEndDate_ValueChanged(object sender, EventArgs e)
	{
		if (gFlagLoadOK != 0)
		{
			showDataPhieuchuacan();
		}
	}

	private void btnInphieu_Click(object sender, EventArgs e)
	{
		try
		{
			string text = "";
			text = "delete * from phieuin ";
			dataFetcher.setCommand(text, 1);
			int num = dataFetcher.execute();
			Thread.Sleep(1000);
			if (insertDataPhieuin(gMaxIDPhieucan.ToString(), 0) != -1 && insertDataPhieuin(gMaxIDPhieucan.ToString(), 1) != -1)
			{
				string valueNoParam = DBUtil.getValueNoParam("select kieuin from giaotiep");
				rptInphieu rptInphieu3 = new rptInphieu();
				rptInphieu3.rptSophieu = gMaxIDPhieucan.ToString();
				if (valueNoParam == "Bằng tay")
				{
					((ActiveReport3)rptInphieu3).Show();
					return;
				}
				((ActiveReport3)rptInphieu3).Run();
				((ActiveReport3)rptInphieu3).Document.Print(false, false, false);
			}
		}
		catch
		{
		}
	}

	private void KhoitaoGridIn(int soDong)
	{
		dataGridTemp.RowCount = 0;
		dataGridTemp.RowCount = soDong;
		for (int i = 0; i < soDong; i++)
		{
			dataGridTemp.Rows[i].Cells[0].Value = "0";
			dataGridTemp.Rows[i].Cells[1].Value = "0";
			dataGridTemp.Rows[i].Cells[2].Value = "0";
			dataGridTemp.Rows[i].Cells[3].Value = "0";
			dataGridTemp.Rows[i].Cells[4].Value = "0";
			dataGridTemp.Rows[i].Cells[5].Value = "0";
			dataGridTemp.Rows[i].Cells[6].Value = "0";
			dataGridTemp.Rows[i].Cells[7].Value = "0";
			dataGridTemp.Rows[i].Cells[8].Value = "0";
			dataGridTemp.Rows[i].Cells[9].Value = "0";
			dataGridTemp.Rows[i].Cells[10].Value = "0";
			dataGridTemp.Rows[i].Cells[11].Value = "0";
			dataGridTemp.Rows[i].Cells[12].Value = "0";
			dataGridTemp.Rows[i].Cells[13].Value = "0";
			dataGridTemp.Rows[i].Cells[14].Value = "0";
			dataGridTemp.Rows[i].Cells[15].Value = "0";
		}
	}

	private int insertDataPhieuin(string phieuid, int loaiCan)
	{
		//IL_0958: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0502: Expected O, but got Unknown
		string text = "";
		Thread.Sleep(1000);
		text = "select trongluong,solong,socon1long*solong as soconga,loaican ";
		text += " from chitietphieucan ";
		object obj = text;
		text = string.Concat(obj, " where phieu_id=", phieuid, " and loaican=", loaiCan);
		text += " order by loaican,ngaycanct";
		dataFetcher.setCommand(text, 1);
		DataSet dataSet = dataFetcher.fetchDataSet();
		int num = 6;
		int num2 = 0;
		int count = dataSet.Tables[0].Rows.Count;
		num2 = ((count % 20 != 0) ? ((count / 20 + 1) * num) : (count / 20 * num));
		KhoitaoGridIn(num2);
		int num3 = 0;
		int num4 = 0;
		int num5 = 0;
		int num6 = 0;
		int result = 0;
		int num7 = 0;
		double num8 = 0.0;
		double num9 = 0.0;
		double num10 = 0.0;
		foreach (DataRow row in dataSet.Tables[0].Rows)
		{
			num6++;
			dataGridTemp.Rows[num * num4 + num5].Cells[num7 * 4].Value = num6.ToString();
			dataGridTemp.Rows[num * num4 + num5].Cells[num7 * 4 + 1].Value = row["trongluong"].ToString();
			dataGridTemp.Rows[num * num4 + num5].Cells[num7 * 4 + 2].Value = row["solong"].ToString();
			dataGridTemp.Rows[num * num4 + num5].Cells[num7 * 4 + 3].Value = row["soconga"].ToString();
			num8 += double.Parse(row["trongluong"].ToString());
			num9 += double.Parse(row["solong"].ToString());
			num10 += double.Parse(row["soconga"].ToString());
			num5++;
			if (num5 == num - 1)
			{
				dataGridTemp.Rows[num * num4 + num5].Cells[num7 * 4].Value = "Tổng";
				dataGridTemp.Rows[num * num4 + num5].Cells[num7 * 4 + 1].Value = num8.ToString();
				dataGridTemp.Rows[num * num4 + num5].Cells[num7 * 4 + 2].Value = num9.ToString();
				dataGridTemp.Rows[num * num4 + num5].Cells[num7 * 4 + 3].Value = num10.ToString();
				num7++;
				num5 = 0;
				num8 = 0.0;
				num9 = 0.0;
				num10 = 0.0;
			}
			num3++;
			if (num3 % 20 == 0 && num3 > 0)
			{
				num4++;
				num7 = 0;
			}
		}
		try
		{
			if (dataGridTemp.Rows[num2 - 1].Cells[num7 * 4].Value.ToString() == "0")
			{
				dataGridTemp.Rows[num2 - 1].Cells[num7 * 4].Value = "Tổng";
				dataGridTemp.Rows[num2 - 1].Cells[num7 * 4 + 1].Value = num8.ToString();
				dataGridTemp.Rows[num2 - 1].Cells[num7 * 4 + 2].Value = num9.ToString();
				dataGridTemp.Rows[num2 - 1].Cells[num7 * 4 + 3].Value = num10.ToString();
			}
		}
		catch
		{
		}
		dataGridTemp.AutoResizeColumns();
		Thread.Sleep(1000);
		OleDbConnection val = new OleDbConnection();
		((DbConnection)(object)val).ConnectionString = dataFetcher.getConString();
		if (((DbConnection)(object)val).State == ConnectionState.Closed)
		{
			((DbConnection)(object)val).Open();
		}
		OleDbTransaction val2 = val.BeginTransaction();
		try
		{
			text = "insert into phieuin(macan1,tl1,solong1,socon1,macan2,tl2,solong2,socon2,macan3,tl3,solong3,socon3,macan4,tl4,solong4,socon4,loaican) ";
			text += " values(@macan1,@tl1,@solong1,@socon1,@macan2,@tl2,@solong2,@socon2,@macan3,@tl3,@solong3,@socon3,@macan4,@tl4,@solong4,@socon4,@loaican)";
			for (int i = 0; i < dataGridTemp.Rows.Count; i++)
			{
				dataFetcher.setTranCommand(text, 1);
				dataFetcher.addTranParameterWithValue("@macan1", dataGridTemp.Rows[i].Cells[0].Value.ToString());
				dataFetcher.addTranParameterWithValue("@tl1", dataGridTemp.Rows[i].Cells[1].Value.ToString());
				dataFetcher.addTranParameterWithValue("@solong1", dataGridTemp.Rows[i].Cells[2].Value.ToString());
				dataFetcher.addTranParameterWithValue("@socon1", dataGridTemp.Rows[i].Cells[3].Value.ToString());
				dataFetcher.addTranParameterWithValue("@macan2", dataGridTemp.Rows[i].Cells[4].Value.ToString());
				dataFetcher.addTranParameterWithValue("@tl2", dataGridTemp.Rows[i].Cells[5].Value.ToString());
				dataFetcher.addTranParameterWithValue("@solong2", dataGridTemp.Rows[i].Cells[6].Value.ToString());
				dataFetcher.addTranParameterWithValue("@socon2", dataGridTemp.Rows[i].Cells[7].Value.ToString());
				dataFetcher.addTranParameterWithValue("@macan3", dataGridTemp.Rows[i].Cells[8].Value.ToString());
				dataFetcher.addTranParameterWithValue("@tl3", dataGridTemp.Rows[i].Cells[9].Value.ToString());
				dataFetcher.addTranParameterWithValue("@solong3", dataGridTemp.Rows[i].Cells[10].Value.ToString());
				dataFetcher.addTranParameterWithValue("@socon3", dataGridTemp.Rows[i].Cells[11].Value.ToString());
				dataFetcher.addTranParameterWithValue("@macan4", dataGridTemp.Rows[i].Cells[12].Value.ToString());
				dataFetcher.addTranParameterWithValue("@tl4", dataGridTemp.Rows[i].Cells[13].Value.ToString());
				dataFetcher.addTranParameterWithValue("@solong4", dataGridTemp.Rows[i].Cells[14].Value.ToString());
				dataFetcher.addTranParameterWithValue("@socon4", dataGridTemp.Rows[i].Cells[15].Value.ToString());
				dataFetcher.addTranParameterWithValue("@loaican", loaiCan.ToString());
				dataFetcher.executeTransaction(val, val2);
			}
			((DbTransaction)(object)val2).Commit();
		}
		catch
		{
			((DbTransaction)(object)val2).Rollback();
			MessageBox.Show("Không in được. Đã có lỗi xảy ra!", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)16);
			result = -1;
		}
		((DbConnection)(object)val).Close();
		return result;
	}

	private void txtSolong_Leave(object sender, EventArgs e)
	{
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			double num = double.Parse(((Control)txtSolong).Text);
		}
		catch
		{
			MessageBox.Show("Nhập số lồng là số.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)16);
			((Control)txtSolong).Focus();
		}
	}

	private void txtSocon1long_Leave(object sender, EventArgs e)
	{
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			double num = double.Parse(((Control)txtSocon1long).Text);
		}
		catch
		{
			MessageBox.Show("Nhập số con trong 1 lồng là số.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)16);
			((Control)txtSocon1long).Focus();
		}
	}

	private void lblDaucan_DoubleClick(object sender, EventArgs e)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			frmPassword frmPassword2 = new frmPassword();
			frmPassword2.gType = "cantay";
			((Form)frmPassword2).ShowDialog();
			if (Class1.gMatkhau == "ok")
			{
				if (!((Control)txtSoluong).Visible)
				{
					((Control)txtSoluong).Visible = true;
					timer1.Enabled = false;
				}
				else
				{
					((Control)txtSoluong).Visible = false;
					timer1.Enabled = true;
				}
				Class1.gMatkhau = "";
			}
			else
			{
				if (Class1.gMatkhau != "")
				{
					MessageBox.Show("Mật khẩu không đúng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
				}
				Class1.gMatkhau = "";
			}
		}
		catch
		{
		}
	}

	private void lblDaucan2_DoubleClick(object sender, EventArgs e)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			frmPassword frmPassword2 = new frmPassword();
			frmPassword2.gType = "cantay";
			((Form)frmPassword2).ShowDialog();
			if (Class1.gMatkhau == "ok")
			{
				if (!((Control)txtSoluong2).Visible)
				{
					((Control)txtSoluong2).Visible = true;
					timer2.Enabled = false;
				}
				else
				{
					((Control)txtSoluong2).Visible = false;
					timer2.Enabled = true;
				}
				Class1.gMatkhau = "";
			}
			else
			{
				if (Class1.gMatkhau != "")
				{
					MessageBox.Show("Mật khẩu không đúng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
				}
				Class1.gMatkhau = "";
			}
		}
		catch
		{
		}
	}

	private void timerChekOndinh_Tick(object sender, EventArgs e)
	{
		try
		{
			double num = 0.0;
			double num2 = 0.0;
			double num3 = 0.0;
			num = double.Parse(((Control)lblDaucan).Text);
			num2 = double.Parse(((Control)txtTLLuutu).Text);
			num3 = double.Parse(((Control)txtTGOndinh).Text);
			if (num >= num2 && flagSaveCan == 0)
			{
				counterOndinhCan++;
				gTLOndinhCan += num;
				if ((double)counterOndinhCan >= num3 * 2.0)
				{
					if (num >= gTLOndinhCan / (double)counterOndinhCan - 0.1 || num <= gTLOndinhCan / (double)counterOndinhCan + 0.1)
					{
						if (flagSaveCan == 0)
						{
							btnSave.PerformClick();
							flagSaveCan = 1;
							counterOndinhCan = 0;
						}
					}
					else
					{
						counterOndinhCan = 0;
						gTLOndinhCan = 0.0;
					}
				}
			}
			if (num < num2 - 0.2)
			{
				flagSaveCan = 0;
				counterOndinhCan = 0;
				gTLOndinhCan = 0.0;
				serialPort1.DtrEnable = false;
			}
		}
		catch
		{
		}
	}

	private void timerChekOndinh2_Tick(object sender, EventArgs e)
	{
		try
		{
			double num = 0.0;
			double num2 = 0.0;
			double num3 = 0.0;
			num = double.Parse(((Control)lblDaucan2).Text);
			num2 = double.Parse(((Control)txtTLLuutu2).Text);
			num3 = double.Parse(((Control)txtTGOndinh2).Text);
			if (num >= num2 && flagSaveCan2 == 0)
			{
				counterOndinhCan2++;
				gTLOndinhCan2 += num;
				if ((double)counterOndinhCan2 >= num3 * 2.0)
				{
					if (num >= gTLOndinhCan2 / (double)counterOndinhCan2 - 0.1 || num <= gTLOndinhCan2 / (double)counterOndinhCan2 + 0.1)
					{
						if (flagSaveCan2 == 0)
						{
							btnSave2.PerformClick();
							flagSaveCan2 = 1;
							counterOndinhCan2 = 0;
						}
					}
					else
					{
						counterOndinhCan2 = 0;
						gTLOndinhCan2 = 0.0;
					}
				}
			}
			if (num < num2 - 0.2)
			{
				flagSaveCan2 = 0;
				counterOndinhCan2 = 0;
				gTLOndinhCan2 = 0.0;
				serialPort2.DtrEnable = false;
			}
		}
		catch
		{
		}
	}

	private void chkCantudong_CheckedChanged(object sender, EventArgs e)
	{
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		if (((Control)cboMahang).Text == "")
		{
			MessageBox.Show("Vui lòng chọn loại hàng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)48);
			((Control)cboMahang).Focus();
			chkCantudong.Checked = false;
		}
		else if (chkCantudong.Checked)
		{
			timerChekOndinh.Enabled = true;
		}
		else
		{
			timerChekOndinh.Enabled = false;
		}
	}

	private void chkCantudong2_CheckedChanged(object sender, EventArgs e)
	{
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		if (((Control)cboMahang).Text == "")
		{
			MessageBox.Show("Vui lòng chọn loại hàng.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)48);
			((Control)cboMahang).Focus();
			chkCantudong2.Checked = false;
		}
		else if (chkCantudong2.Checked)
		{
			timerChekOndinh2.Enabled = true;
		}
		else
		{
			timerChekOndinh2.Enabled = false;
		}
	}

	private void cboSoxe_Leave(object sender, EventArgs e)
	{
		try
		{
			if (!(gMaxIDPhieucan <= 0.0))
			{
				capNhatPhieucan();
				Thread.Sleep(700);
				showDataPhieuchuacan();
			}
		}
		catch
		{
		}
	}

	private void txtNguoinhan_Leave(object sender, EventArgs e)
	{
		try
		{
			if (!(gMaxIDPhieucan <= 0.0))
			{
				capNhatPhieucan();
				Thread.Sleep(700);
				showDataPhieuchuacan();
			}
		}
		catch
		{
		}
	}

	private void frmCan_DoubleClick(object sender, EventArgs e)
	{
		try
		{
			if (!((Control)groupBox2).Visible)
			{
				((Control)groupBox2).Visible = true;
			}
			else
			{
				((Control)groupBox2).Visible = false;
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
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Expected O, but got Unknown
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Expected O, but got Unknown
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Expected O, but got Unknown
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0040: Expected O, but got Unknown
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		//IL_005c: Expected O, but got Unknown
		//IL_005d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0067: Expected O, but got Unknown
		//IL_0068: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Expected O, but got Unknown
		//IL_0073: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0088: Expected O, but got Unknown
		//IL_0089: Unknown result type (might be due to invalid IL or missing references)
		//IL_0093: Expected O, but got Unknown
		//IL_0094: Unknown result type (might be due to invalid IL or missing references)
		//IL_009e: Expected O, but got Unknown
		//IL_009f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a9: Expected O, but got Unknown
		//IL_00aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b4: Expected O, but got Unknown
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bf: Expected O, but got Unknown
		//IL_00c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ca: Expected O, but got Unknown
		//IL_00cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d5: Expected O, but got Unknown
		//IL_00dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e6: Expected O, but got Unknown
		//IL_00ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f7: Expected O, but got Unknown
		//IL_00fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0108: Expected O, but got Unknown
		//IL_0109: Unknown result type (might be due to invalid IL or missing references)
		//IL_0113: Expected O, but got Unknown
		//IL_0114: Unknown result type (might be due to invalid IL or missing references)
		//IL_011e: Expected O, but got Unknown
		//IL_011f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0129: Expected O, but got Unknown
		//IL_012a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0134: Expected O, but got Unknown
		//IL_0135: Unknown result type (might be due to invalid IL or missing references)
		//IL_013f: Expected O, but got Unknown
		//IL_0140: Unknown result type (might be due to invalid IL or missing references)
		//IL_014a: Expected O, but got Unknown
		//IL_014b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0155: Expected O, but got Unknown
		//IL_0156: Unknown result type (might be due to invalid IL or missing references)
		//IL_0160: Expected O, but got Unknown
		//IL_0161: Unknown result type (might be due to invalid IL or missing references)
		//IL_016b: Expected O, but got Unknown
		//IL_016c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0176: Expected O, but got Unknown
		//IL_0177: Unknown result type (might be due to invalid IL or missing references)
		//IL_0181: Expected O, but got Unknown
		//IL_0182: Unknown result type (might be due to invalid IL or missing references)
		//IL_018c: Expected O, but got Unknown
		//IL_018d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0197: Expected O, but got Unknown
		//IL_0198: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a2: Expected O, but got Unknown
		//IL_01a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ad: Expected O, but got Unknown
		//IL_01ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b8: Expected O, but got Unknown
		//IL_01b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c3: Expected O, but got Unknown
		//IL_01c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ce: Expected O, but got Unknown
		//IL_01cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d9: Expected O, but got Unknown
		//IL_01da: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e4: Expected O, but got Unknown
		//IL_01e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ef: Expected O, but got Unknown
		//IL_01f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Expected O, but got Unknown
		//IL_01fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0205: Expected O, but got Unknown
		//IL_0206: Unknown result type (might be due to invalid IL or missing references)
		//IL_0210: Expected O, but got Unknown
		//IL_0211: Unknown result type (might be due to invalid IL or missing references)
		//IL_021b: Expected O, but got Unknown
		//IL_021c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0226: Expected O, but got Unknown
		//IL_0227: Unknown result type (might be due to invalid IL or missing references)
		//IL_0231: Expected O, but got Unknown
		//IL_0232: Unknown result type (might be due to invalid IL or missing references)
		//IL_023c: Expected O, but got Unknown
		//IL_023d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0247: Expected O, but got Unknown
		//IL_0248: Unknown result type (might be due to invalid IL or missing references)
		//IL_0252: Expected O, but got Unknown
		//IL_0253: Unknown result type (might be due to invalid IL or missing references)
		//IL_025d: Expected O, but got Unknown
		//IL_025e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0268: Expected O, but got Unknown
		//IL_0269: Unknown result type (might be due to invalid IL or missing references)
		//IL_0273: Expected O, but got Unknown
		//IL_0274: Unknown result type (might be due to invalid IL or missing references)
		//IL_027e: Expected O, but got Unknown
		//IL_027f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0289: Expected O, but got Unknown
		//IL_028a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0294: Expected O, but got Unknown
		//IL_0295: Unknown result type (might be due to invalid IL or missing references)
		//IL_029f: Expected O, but got Unknown
		//IL_02a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_02aa: Expected O, but got Unknown
		//IL_02ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b5: Expected O, but got Unknown
		//IL_02b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c0: Expected O, but got Unknown
		//IL_02c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_02cb: Expected O, but got Unknown
		//IL_02cc: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d6: Expected O, but got Unknown
		//IL_02d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e1: Expected O, but got Unknown
		//IL_02e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ec: Expected O, but got Unknown
		//IL_02ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f7: Expected O, but got Unknown
		//IL_02f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0302: Expected O, but got Unknown
		//IL_0303: Unknown result type (might be due to invalid IL or missing references)
		//IL_030d: Expected O, but got Unknown
		//IL_030e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0318: Expected O, but got Unknown
		//IL_0319: Unknown result type (might be due to invalid IL or missing references)
		//IL_0323: Expected O, but got Unknown
		//IL_0324: Unknown result type (might be due to invalid IL or missing references)
		//IL_032e: Expected O, but got Unknown
		//IL_032f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0339: Expected O, but got Unknown
		//IL_033a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0344: Expected O, but got Unknown
		//IL_0345: Unknown result type (might be due to invalid IL or missing references)
		//IL_034f: Expected O, but got Unknown
		//IL_0350: Unknown result type (might be due to invalid IL or missing references)
		//IL_035a: Expected O, but got Unknown
		//IL_035b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0365: Expected O, but got Unknown
		//IL_0366: Unknown result type (might be due to invalid IL or missing references)
		//IL_0370: Expected O, but got Unknown
		//IL_0371: Unknown result type (might be due to invalid IL or missing references)
		//IL_037b: Expected O, but got Unknown
		//IL_037c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0386: Expected O, but got Unknown
		//IL_0387: Unknown result type (might be due to invalid IL or missing references)
		//IL_0391: Expected O, but got Unknown
		//IL_0392: Unknown result type (might be due to invalid IL or missing references)
		//IL_039c: Expected O, but got Unknown
		//IL_039d: Unknown result type (might be due to invalid IL or missing references)
		//IL_03a7: Expected O, but got Unknown
		//IL_03a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b2: Expected O, but got Unknown
		//IL_03b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_03bd: Expected O, but got Unknown
		//IL_03be: Unknown result type (might be due to invalid IL or missing references)
		//IL_03c8: Expected O, but got Unknown
		//IL_03c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d3: Expected O, but got Unknown
		//IL_03d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03de: Expected O, but got Unknown
		//IL_03df: Unknown result type (might be due to invalid IL or missing references)
		//IL_03e9: Expected O, but got Unknown
		//IL_03ea: Unknown result type (might be due to invalid IL or missing references)
		//IL_03f4: Expected O, but got Unknown
		//IL_03f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ff: Expected O, but got Unknown
		//IL_0400: Unknown result type (might be due to invalid IL or missing references)
		//IL_040a: Expected O, but got Unknown
		//IL_040b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0415: Expected O, but got Unknown
		//IL_0416: Unknown result type (might be due to invalid IL or missing references)
		//IL_0420: Expected O, but got Unknown
		//IL_0421: Unknown result type (might be due to invalid IL or missing references)
		//IL_042b: Expected O, but got Unknown
		//IL_042c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0436: Expected O, but got Unknown
		//IL_0437: Unknown result type (might be due to invalid IL or missing references)
		//IL_0441: Expected O, but got Unknown
		//IL_0442: Unknown result type (might be due to invalid IL or missing references)
		//IL_044c: Expected O, but got Unknown
		//IL_044d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0457: Expected O, but got Unknown
		//IL_0458: Unknown result type (might be due to invalid IL or missing references)
		//IL_0462: Expected O, but got Unknown
		//IL_0463: Unknown result type (might be due to invalid IL or missing references)
		//IL_046d: Expected O, but got Unknown
		//IL_046e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0478: Expected O, but got Unknown
		//IL_0479: Unknown result type (might be due to invalid IL or missing references)
		//IL_0483: Expected O, but got Unknown
		//IL_048a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0494: Expected O, but got Unknown
		//IL_049b: Unknown result type (might be due to invalid IL or missing references)
		//IL_04a5: Expected O, but got Unknown
		//IL_04a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b0: Expected O, but got Unknown
		//IL_04b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_04bb: Expected O, but got Unknown
		//IL_04bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_04c6: Expected O, but got Unknown
		//IL_04c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d1: Expected O, but got Unknown
		//IL_04d2: Unknown result type (might be due to invalid IL or missing references)
		//IL_04dc: Expected O, but got Unknown
		//IL_04dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e7: Expected O, but got Unknown
		//IL_04e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f2: Expected O, but got Unknown
		//IL_04f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fd: Expected O, but got Unknown
		//IL_04fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0508: Expected O, but got Unknown
		//IL_0509: Unknown result type (might be due to invalid IL or missing references)
		//IL_0513: Expected O, but got Unknown
		//IL_0514: Unknown result type (might be due to invalid IL or missing references)
		//IL_051e: Expected O, but got Unknown
		//IL_051f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0529: Expected O, but got Unknown
		//IL_052a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0534: Expected O, but got Unknown
		//IL_053b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0545: Expected O, but got Unknown
		//IL_064d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0657: Expected O, but got Unknown
		//IL_0946: Unknown result type (might be due to invalid IL or missing references)
		//IL_0950: Expected O, but got Unknown
		//IL_09fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a04: Expected O, but got Unknown
		//IL_0b06: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b10: Expected O, but got Unknown
		//IL_0ba4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bae: Expected O, but got Unknown
		//IL_0c2e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c38: Expected O, but got Unknown
		//IL_0ccf: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cd9: Expected O, but got Unknown
		//IL_0d85: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d8f: Expected O, but got Unknown
		//IL_0e0c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e16: Expected O, but got Unknown
		//IL_0ebf: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ec9: Expected O, but got Unknown
		//IL_0f3c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f46: Expected O, but got Unknown
		//IL_0fcd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0fd7: Expected O, but got Unknown
		//IL_1065: Unknown result type (might be due to invalid IL or missing references)
		//IL_106f: Expected O, but got Unknown
		//IL_1111: Unknown result type (might be due to invalid IL or missing references)
		//IL_111b: Expected O, but got Unknown
		//IL_11ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_11d8: Expected O, but got Unknown
		//IL_1266: Unknown result type (might be due to invalid IL or missing references)
		//IL_1270: Expected O, but got Unknown
		//IL_12f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_12fd: Expected O, but got Unknown
		//IL_1390: Unknown result type (might be due to invalid IL or missing references)
		//IL_139a: Expected O, but got Unknown
		//IL_143b: Unknown result type (might be due to invalid IL or missing references)
		//IL_1445: Expected O, but got Unknown
		//IL_1561: Unknown result type (might be due to invalid IL or missing references)
		//IL_156b: Expected O, but got Unknown
		//IL_1879: Unknown result type (might be due to invalid IL or missing references)
		//IL_1883: Expected O, but got Unknown
		//IL_1b53: Unknown result type (might be due to invalid IL or missing references)
		//IL_1b5d: Expected O, but got Unknown
		//IL_1bee: Unknown result type (might be due to invalid IL or missing references)
		//IL_1bf8: Expected O, but got Unknown
		//IL_1ced: Unknown result type (might be due to invalid IL or missing references)
		//IL_1cf7: Expected O, but got Unknown
		//IL_1d7c: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d86: Expected O, but got Unknown
		//IL_1dec: Unknown result type (might be due to invalid IL or missing references)
		//IL_1df6: Expected O, but got Unknown
		//IL_1e8b: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e95: Expected O, but got Unknown
		//IL_1f78: Unknown result type (might be due to invalid IL or missing references)
		//IL_1f82: Expected O, but got Unknown
		//IL_201e: Unknown result type (might be due to invalid IL or missing references)
		//IL_2028: Expected O, but got Unknown
		//IL_20bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_20c9: Expected O, but got Unknown
		//IL_2158: Unknown result type (might be due to invalid IL or missing references)
		//IL_2162: Expected O, but got Unknown
		//IL_21f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_2200: Expected O, but got Unknown
		//IL_228f: Unknown result type (might be due to invalid IL or missing references)
		//IL_2299: Expected O, but got Unknown
		//IL_2330: Unknown result type (might be due to invalid IL or missing references)
		//IL_233a: Expected O, but got Unknown
		//IL_23c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_23d3: Expected O, but got Unknown
		//IL_253a: Unknown result type (might be due to invalid IL or missing references)
		//IL_2544: Expected O, but got Unknown
		//IL_280a: Unknown result type (might be due to invalid IL or missing references)
		//IL_2814: Expected O, but got Unknown
		//IL_29af: Unknown result type (might be due to invalid IL or missing references)
		//IL_29b9: Expected O, but got Unknown
		//IL_2a37: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a41: Expected O, but got Unknown
		//IL_2b35: Unknown result type (might be due to invalid IL or missing references)
		//IL_2b3f: Expected O, but got Unknown
		//IL_2baf: Unknown result type (might be due to invalid IL or missing references)
		//IL_2bb9: Expected O, but got Unknown
		//IL_2c36: Unknown result type (might be due to invalid IL or missing references)
		//IL_2c40: Expected O, but got Unknown
		//IL_2cad: Unknown result type (might be due to invalid IL or missing references)
		//IL_2cb7: Expected O, but got Unknown
		//IL_2d38: Unknown result type (might be due to invalid IL or missing references)
		//IL_2d42: Expected O, but got Unknown
		//IL_2e00: Unknown result type (might be due to invalid IL or missing references)
		//IL_2e0a: Expected O, but got Unknown
		//IL_2fcd: Unknown result type (might be due to invalid IL or missing references)
		//IL_2fd7: Expected O, but got Unknown
		//IL_3055: Unknown result type (might be due to invalid IL or missing references)
		//IL_305f: Expected O, but got Unknown
		//IL_3153: Unknown result type (might be due to invalid IL or missing references)
		//IL_315d: Expected O, but got Unknown
		//IL_31cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_31d7: Expected O, but got Unknown
		//IL_3254: Unknown result type (might be due to invalid IL or missing references)
		//IL_325e: Expected O, but got Unknown
		//IL_32cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_32d5: Expected O, but got Unknown
		//IL_33a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_33ae: Expected O, but got Unknown
		//IL_345a: Unknown result type (might be due to invalid IL or missing references)
		//IL_3464: Expected O, but got Unknown
		//IL_34e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_34f2: Expected O, but got Unknown
		//IL_3589: Unknown result type (might be due to invalid IL or missing references)
		//IL_3593: Expected O, but got Unknown
		//IL_3622: Unknown result type (might be due to invalid IL or missing references)
		//IL_362c: Expected O, but got Unknown
		//IL_36c3: Unknown result type (might be due to invalid IL or missing references)
		//IL_36cd: Expected O, but got Unknown
		//IL_375c: Unknown result type (might be due to invalid IL or missing references)
		//IL_3766: Expected O, but got Unknown
		//IL_3809: Unknown result type (might be due to invalid IL or missing references)
		//IL_3813: Expected O, but got Unknown
		//IL_390a: Unknown result type (might be due to invalid IL or missing references)
		//IL_3914: Expected O, but got Unknown
		//IL_39ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_39f8: Expected O, but got Unknown
		//IL_3ace: Unknown result type (might be due to invalid IL or missing references)
		//IL_3ad8: Expected O, but got Unknown
		//IL_3bb9: Unknown result type (might be due to invalid IL or missing references)
		//IL_3bc3: Expected O, but got Unknown
		//IL_3ca7: Unknown result type (might be due to invalid IL or missing references)
		//IL_3cb1: Expected O, but got Unknown
		//IL_41a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_41ac: Expected O, but got Unknown
		//IL_41ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_41c4: Expected O, but got Unknown
		//IL_41cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_4230: Unknown result type (might be due to invalid IL or missing references)
		//IL_423a: Expected O, but got Unknown
		//IL_4256: Unknown result type (might be due to invalid IL or missing references)
		//IL_4260: Expected O, but got Unknown
		components = new Container();
		DataGridViewCellStyle val = new DataGridViewCellStyle();
		DataGridViewCellStyle val2 = new DataGridViewCellStyle();
		DataGridViewCellStyle val3 = new DataGridViewCellStyle();
		DataGridViewCellStyle val4 = new DataGridViewCellStyle();
		DataGridViewCellStyle val5 = new DataGridViewCellStyle();
		DataGridViewCellStyle val6 = new DataGridViewCellStyle();
		DataGridViewCellStyle val7 = new DataGridViewCellStyle();
		DataGridViewCellStyle val8 = new DataGridViewCellStyle();
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmCan));
		dataGrid = new DataGridView();
		id = new DataGridViewTextBoxColumn();
		STT2 = new DataGridViewTextBoxColumn();
		mathang = new DataGridViewTextBoxColumn();
		soluong = new DataGridViewTextBoxColumn();
		solong = new DataGridViewTextBoxColumn();
		socon1long = new DataGridViewTextBoxColumn();
		tongsocon = new DataGridViewTextBoxColumn();
		hinhthuccan = new DataGridViewTextBoxColumn();
		ngaycan = new DataGridViewTextBoxColumn();
		lblDaucan = new Label();
		txtSoluong = new TextBox();
		serialPort1 = new SerialPort(components);
		timer1 = new Timer(components);
		timerChekOndinh = new Timer(components);
		cboMahang = new ComboBox();
		label1 = new Label();
		cboTenhang = new ComboBox();
		cboTenKH = new ComboBox();
		label2 = new Label();
		cboMaKH = new ComboBox();
		label3 = new Label();
		txtNguoinhan = new TextBox();
		label4 = new Label();
		cboSoxe = new ComboBox();
		cboXuatnhap = new ComboBox();
		label5 = new Label();
		label6 = new Label();
		label7 = new Label();
		lblTotal = new Label();
		lblSolancan = new Label();
		dataGridPhieucan = new DataGridView();
		stt = new DataGridViewTextBoxColumn();
		dataGridViewTextBoxColumn2 = new DataGridViewTextBoxColumn();
		sophieu = new DataGridViewTextBoxColumn();
		soxe = new DataGridViewTextBoxColumn();
		tenkh = new DataGridViewTextBoxColumn();
		dataGridViewTextBoxColumn4 = new DataGridViewTextBoxColumn();
		xuatnhap = new DataGridViewTextBoxColumn();
		ghichu = new DataGridViewTextBoxColumn();
		groupBox1 = new GroupBox();
		chkCanbi2 = new CheckBox();
		txtEndDate = new DateTimePicker();
		label9 = new Label();
		label10 = new Label();
		txtStartDate = new DateTimePicker();
		chkCanbi = new CheckBox();
		chkCantudong = new CheckBox();
		panel1 = new Panel();
		label8 = new Label();
		lblSophieu = new Label();
		txtSolong = new TextBox();
		label11 = new Label();
		panel2 = new Panel();
		txtSocon1long = new TextBox();
		label12 = new Label();
		lblTongsolong = new Label();
		label14 = new Label();
		lblTongsocon = new Label();
		label16 = new Label();
		lblTLTrungbinh = new Label();
		label15 = new Label();
		dataGridTemp = new DataGridView();
		macan1 = new DataGridViewTextBoxColumn();
		tl1 = new DataGridViewTextBoxColumn();
		solong1 = new DataGridViewTextBoxColumn();
		socon1 = new DataGridViewTextBoxColumn();
		macan2 = new DataGridViewTextBoxColumn();
		tl2 = new DataGridViewTextBoxColumn();
		solong2 = new DataGridViewTextBoxColumn();
		socon2 = new DataGridViewTextBoxColumn();
		macan3 = new DataGridViewTextBoxColumn();
		tl3 = new DataGridViewTextBoxColumn();
		solong3 = new DataGridViewTextBoxColumn();
		socon3 = new DataGridViewTextBoxColumn();
		macan4 = new DataGridViewTextBoxColumn();
		tl4 = new DataGridViewTextBoxColumn();
		solong4 = new DataGridViewTextBoxColumn();
		socon4 = new DataGridViewTextBoxColumn();
		lblOndinhKytu = new Label();
		groupBox2 = new GroupBox();
		label18 = new Label();
		cboCongCom1 = new ComboBox();
		label13 = new Label();
		txtTLLuutu = new TextBox();
		label17 = new Label();
		txtTGOndinh = new TextBox();
		lblDaucan2 = new Label();
		txtSoluong2 = new TextBox();
		groupBox3 = new GroupBox();
		label19 = new Label();
		cboCongCom2 = new ComboBox();
		label20 = new Label();
		txtTLLuutu2 = new TextBox();
		label21 = new Label();
		txtTGOndinh2 = new TextBox();
		timerChekOndinh2 = new Timer(components);
		timer2 = new Timer(components);
		chkCantudong2 = new CheckBox();
		txtSolongbi = new TextBox();
		label22 = new Label();
		lblSolancanbi = new Label();
		label24 = new Label();
		lblSolongcanbi = new Label();
		label25 = new Label();
		btnSave2 = new Button();
		btnInphieu = new Button();
		btnTaophieu = new Button();
		btnDelete = new Button();
		btnKetthuc = new Button();
		btnSave = new Button();
		serialPort2 = new SerialPort(components);
		((ISupportInitialize)dataGrid).BeginInit();
		((ISupportInitialize)dataGridPhieucan).BeginInit();
		((Control)groupBox1).SuspendLayout();
		((ISupportInitialize)dataGridTemp).BeginInit();
		((Control)groupBox2).SuspendLayout();
		((Control)groupBox3).SuspendLayout();
		((Control)this).SuspendLayout();
		dataGrid.AllowUserToAddRows = false;
		dataGrid.AllowUserToOrderColumns = true;
		dataGrid.ColumnHeadersHeightSizeMode = (DataGridViewColumnHeadersHeightSizeMode)2;
		dataGrid.Columns.AddRange((DataGridViewColumn[])(object)new DataGridViewColumn[9]
		{
			(DataGridViewColumn)id,
			(DataGridViewColumn)STT2,
			(DataGridViewColumn)mathang,
			(DataGridViewColumn)soluong,
			(DataGridViewColumn)solong,
			(DataGridViewColumn)socon1long,
			(DataGridViewColumn)tongsocon,
			(DataGridViewColumn)hinhthuccan,
			(DataGridViewColumn)ngaycan
		});
		((Control)dataGrid).DataBindings.Add(new Binding("Font", (object)Settings.Default, "font", true, (DataSourceUpdateMode)1));
		dataGrid.EditMode = (DataGridViewEditMode)4;
		((Control)dataGrid).Font = Settings.Default.font;
		((Control)dataGrid).Location = new Point(661, 141);
		((Control)dataGrid).Name = "dataGrid";
		dataGrid.RowTemplate.Height = 25;
		dataGrid.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)dataGrid).Size = new Size(532, 492);
		((Control)dataGrid).TabIndex = 11;
		((DataGridViewColumn)id).HeaderText = "ID";
		((DataGridViewColumn)id).Name = "id";
		((DataGridViewBand)id).Visible = false;
		((DataGridViewColumn)STT2).HeaderText = "STT";
		((DataGridViewColumn)STT2).Name = "STT2";
		((DataGridViewColumn)STT2).Width = 50;
		((DataGridViewColumn)mathang).HeaderText = "Loại hàng";
		((DataGridViewColumn)mathang).Name = "mathang";
		((DataGridViewColumn)mathang).Width = 150;
		val.Alignment = (DataGridViewContentAlignment)64;
		val.Format = "#,##0.0";
		val.NullValue = null;
		((DataGridViewBand)soluong).DefaultCellStyle = val;
		((DataGridViewColumn)soluong).HeaderText = "T.Lượng";
		((DataGridViewColumn)soluong).Name = "soluong";
		val2.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)solong).DefaultCellStyle = val2;
		((DataGridViewColumn)solong).HeaderText = "Số lồng";
		((DataGridViewColumn)solong).Name = "solong";
		val3.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)socon1long).DefaultCellStyle = val3;
		((DataGridViewColumn)socon1long).HeaderText = "Số con 1 lồng";
		((DataGridViewColumn)socon1long).Name = "socon1long";
		((DataGridViewColumn)socon1long).Width = 200;
		val4.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)tongsocon).DefaultCellStyle = val4;
		((DataGridViewColumn)tongsocon).HeaderText = "Tổng số con";
		((DataGridViewColumn)tongsocon).Name = "tongsocon";
		((DataGridViewColumn)tongsocon).Width = 140;
		((DataGridViewColumn)hinhthuccan).HeaderText = "Hình thức cân";
		((DataGridViewColumn)hinhthuccan).Name = "hinhthuccan";
		((DataGridViewColumn)hinhthuccan).Width = 150;
		val5.Format = "dd/MM/yyyy HH:mm:ss";
		val5.NullValue = null;
		((DataGridViewBand)ngaycan).DefaultCellStyle = val5;
		((DataGridViewColumn)ngaycan).HeaderText = "Ngày giờ cân";
		((DataGridViewColumn)ngaycan).Name = "ngaycan";
		((DataGridViewColumn)ngaycan).Width = 230;
		((Control)lblDaucan).BackColor = Color.Black;
		((Control)lblDaucan).Font = new Font("Arial", 48f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblDaucan).ForeColor = Color.Lime;
		((Control)lblDaucan).Location = new Point(9, 8);
		((Control)lblDaucan).Name = "lblDaucan";
		((Control)lblDaucan).Size = new Size(642, 72);
		((Control)lblDaucan).TabIndex = 143;
		((Control)lblDaucan).Text = "1";
		lblDaucan.TextAlign = (ContentAlignment)32;
		((Control)lblDaucan).DoubleClick += lblDaucan_DoubleClick;
		((Control)txtSoluong).Font = new Font("Arial", 20.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)txtSoluong).Location = new Point(13, 38);
		((Control)txtSoluong).Name = "txtSoluong";
		((Control)txtSoluong).Size = new Size(74, 39);
		((Control)txtSoluong).TabIndex = 144;
		((Control)txtSoluong).Text = "0";
		txtSoluong.TextAlign = (HorizontalAlignment)2;
		((Control)txtSoluong).Visible = false;
		((Control)txtSoluong).TextChanged += txtSoluong_TextChanged;
		timer1.Interval = 200;
		timer1.Tick += timer1_Tick;
		timerChekOndinh.Interval = 500;
		timerChekOndinh.Tick += timerChekOndinh_Tick;
		cboMahang.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboMahang).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboMahang).FormattingEnabled = true;
		((Control)cboMahang).Location = new Point(108, 184);
		((Control)cboMahang).Name = "cboMahang";
		((Control)cboMahang).Size = new Size(175, 28);
		((Control)cboMahang).TabIndex = 6;
		cboMahang.SelectedIndexChanged += cboMahang_SelectedIndexChanged;
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(8, 187);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(79, 20);
		((Control)label1).TabIndex = 171;
		((Control)label1).Text = "Loại hàng";
		cboTenhang.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboTenhang).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboTenhang).FormattingEnabled = true;
		((Control)cboTenhang).Location = new Point(289, 184);
		((Control)cboTenhang).Name = "cboTenhang";
		((Control)cboTenhang).Size = new Size(363, 28);
		((Control)cboTenhang).TabIndex = 7;
		cboTenhang.SelectedIndexChanged += cboTenhang_SelectedIndexChanged;
		cboTenKH.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboTenKH).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboTenKH).FormattingEnabled = true;
		((Control)cboTenKH).Location = new Point(289, 116);
		((Control)cboTenKH).Name = "cboTenKH";
		((Control)cboTenKH).Size = new Size(363, 28);
		((Control)cboTenKH).TabIndex = 3;
		cboTenKH.SelectedIndexChanged += cboTenKH_SelectedIndexChanged;
		((Control)cboTenKH).Click += cboTenKH_Click;
		((Control)label2).AutoSize = true;
		((Control)label2).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label2).Location = new Point(8, 119);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(94, 20);
		((Control)label2).TabIndex = 174;
		((Control)label2).Text = "Khách hàng";
		cboMaKH.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboMaKH).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboMaKH).FormattingEnabled = true;
		((Control)cboMaKH).Location = new Point(108, 116);
		((Control)cboMaKH).Name = "cboMaKH";
		((Control)cboMaKH).Size = new Size(175, 28);
		((Control)cboMaKH).TabIndex = 2;
		cboMaKH.SelectedIndexChanged += cboMaKH_SelectedIndexChanged;
		((Control)cboMaKH).Click += cboMaKH_Click;
		((Control)label3).AutoSize = true;
		((Control)label3).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label3).Location = new Point(8, 151);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(64, 20);
		((Control)label3).TabIndex = 176;
		((Control)label3).Text = "Ghi chú";
		((Control)txtNguoinhan).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtNguoinhan).Location = new Point(108, 148);
		((Control)txtNguoinhan).Name = "txtNguoinhan";
		((Control)txtNguoinhan).Size = new Size(544, 26);
		((Control)txtNguoinhan).TabIndex = 5;
		((Control)txtNguoinhan).Leave += txtNguoinhan_Leave;
		((Control)label4).AutoSize = true;
		((Control)label4).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label4).Location = new Point(8, 88);
		((Control)label4).Name = "label4";
		((Control)label4).Size = new Size(49, 20);
		((Control)label4).TabIndex = 178;
		((Control)label4).Text = "Số xe";
		cboSoxe.AutoCompleteMode = (AutoCompleteMode)3;
		cboSoxe.AutoCompleteSource = (AutoCompleteSource)256;
		((Control)cboSoxe).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboSoxe).FormattingEnabled = true;
		((Control)cboSoxe).Location = new Point(108, 85);
		((Control)cboSoxe).Name = "cboSoxe";
		((Control)cboSoxe).Size = new Size(176, 28);
		((Control)cboSoxe).TabIndex = 1;
		((Control)cboSoxe).Leave += cboSoxe_Leave;
		cboXuatnhap.AutoCompleteMode = (AutoCompleteMode)3;
		cboXuatnhap.AutoCompleteSource = (AutoCompleteSource)256;
		((Control)cboXuatnhap).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboXuatnhap).FormattingEnabled = true;
		cboXuatnhap.Items.AddRange(new object[2] { "Nhập", "Xuất" });
		((Control)cboXuatnhap).Location = new Point(566, 36);
		((Control)cboXuatnhap).Name = "cboXuatnhap";
		((Control)cboXuatnhap).Size = new Size(75, 28);
		((Control)cboXuatnhap).TabIndex = 4;
		((Control)cboXuatnhap).Visible = false;
		((Control)label5).AutoSize = true;
		((Control)label5).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label5).Location = new Point(566, 13);
		((Control)label5).Name = "label5";
		((Control)label5).Size = new Size(85, 20);
		((Control)label5).TabIndex = 180;
		((Control)label5).Text = "Xuất/Nhập";
		((Control)label5).Visible = false;
		((Control)label6).AutoSize = true;
		((Control)label6).Font = new Font("Microsoft Sans Serif", 14.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label6).Location = new Point(925, 8);
		((Control)label6).Name = "label6";
		((Control)label6).Size = new Size(138, 24);
		((Control)label6).TabIndex = 183;
		((Control)label6).Text = "Số lần cân gà";
		((Control)label7).AutoSize = true;
		((Control)label7).Font = new Font("Microsoft Sans Serif", 14.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label7).Location = new Point(658, 8);
		((Control)label7).Name = "label7";
		((Control)label7).Size = new Size(89, 24);
		((Control)label7).TabIndex = 184;
		((Control)label7).Text = "Tổng TL";
		((Control)lblTotal).BackColor = Color.FromArgb(224, 224, 224);
		((Control)lblTotal).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblTotal).ForeColor = Color.Red;
		((Control)lblTotal).Location = new Point(793, 2);
		((Control)lblTotal).Name = "lblTotal";
		((Control)lblTotal).Size = new Size(118, 31);
		((Control)lblTotal).TabIndex = 185;
		lblTotal.TextAlign = (ContentAlignment)32;
		((Control)lblSolancan).BackColor = Color.FromArgb(224, 224, 224);
		((Control)lblSolancan).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblSolancan).ForeColor = Color.Red;
		((Control)lblSolancan).Location = new Point(1093, 2);
		((Control)lblSolancan).Name = "lblSolancan";
		((Control)lblSolancan).Size = new Size(97, 31);
		((Control)lblSolancan).TabIndex = 186;
		lblSolancan.TextAlign = (ContentAlignment)32;
		dataGridPhieucan.AllowUserToAddRows = false;
		dataGridPhieucan.AllowUserToOrderColumns = true;
		dataGridPhieucan.ColumnHeadersHeightSizeMode = (DataGridViewColumnHeadersHeightSizeMode)2;
		dataGridPhieucan.Columns.AddRange((DataGridViewColumn[])(object)new DataGridViewColumn[8]
		{
			(DataGridViewColumn)stt,
			(DataGridViewColumn)dataGridViewTextBoxColumn2,
			(DataGridViewColumn)sophieu,
			(DataGridViewColumn)soxe,
			(DataGridViewColumn)tenkh,
			(DataGridViewColumn)dataGridViewTextBoxColumn4,
			(DataGridViewColumn)xuatnhap,
			(DataGridViewColumn)ghichu
		});
		((Control)dataGridPhieucan).DataBindings.Add(new Binding("Font", (object)Settings.Default, "font", true, (DataSourceUpdateMode)1));
		dataGridPhieucan.EditMode = (DataGridViewEditMode)4;
		((Control)dataGridPhieucan).Font = Settings.Default.font;
		((Control)dataGridPhieucan).Location = new Point(1, 401);
		((Control)dataGridPhieucan).Name = "dataGridPhieucan";
		dataGridPhieucan.RowTemplate.Height = 25;
		dataGridPhieucan.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)dataGridPhieucan).Size = new Size(654, 231);
		((Control)dataGridPhieucan).TabIndex = 188;
		((Control)dataGridPhieucan).DoubleClick += dataGridPhieucan_DoubleClick;
		((DataGridViewColumn)stt).HeaderText = "ID";
		((DataGridViewColumn)stt).Name = "stt";
		((DataGridViewBand)stt).Visible = false;
		((DataGridViewColumn)dataGridViewTextBoxColumn2).HeaderText = "STT";
		((DataGridViewColumn)dataGridViewTextBoxColumn2).Name = "dataGridViewTextBoxColumn2";
		((DataGridViewColumn)dataGridViewTextBoxColumn2).Width = 50;
		val6.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)sophieu).DefaultCellStyle = val6;
		((DataGridViewColumn)sophieu).HeaderText = "Số phiếu";
		((DataGridViewColumn)sophieu).Name = "sophieu";
		val7.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)soxe).DefaultCellStyle = val7;
		((DataGridViewColumn)soxe).HeaderText = "Số xe";
		((DataGridViewColumn)soxe).Name = "soxe";
		((DataGridViewColumn)soxe).Width = 120;
		((DataGridViewColumn)tenkh).HeaderText = "Khách hàng";
		((DataGridViewColumn)tenkh).Name = "tenkh";
		((DataGridViewColumn)tenkh).Width = 140;
		val8.Format = "dd/MM/yyyy HH:mm:ss";
		val8.NullValue = null;
		((DataGridViewBand)dataGridViewTextBoxColumn4).DefaultCellStyle = val8;
		((DataGridViewColumn)dataGridViewTextBoxColumn4).HeaderText = "Ngày giờ cân";
		((DataGridViewColumn)dataGridViewTextBoxColumn4).Name = "dataGridViewTextBoxColumn4";
		((DataGridViewColumn)dataGridViewTextBoxColumn4).Width = 230;
		((DataGridViewColumn)xuatnhap).HeaderText = "Xuất/Nhập";
		((DataGridViewColumn)xuatnhap).Name = "xuatnhap";
		((DataGridViewBand)xuatnhap).Visible = false;
		((DataGridViewColumn)xuatnhap).Width = 120;
		((DataGridViewColumn)ghichu).HeaderText = "Ghi chú";
		((DataGridViewColumn)ghichu).Name = "ghichu";
		((DataGridViewColumn)ghichu).Width = 120;
		((Control)groupBox1).Controls.Add((Control)(object)chkCanbi2);
		((Control)groupBox1).Controls.Add((Control)(object)txtEndDate);
		((Control)groupBox1).Controls.Add((Control)(object)label9);
		((Control)groupBox1).Controls.Add((Control)(object)label10);
		((Control)groupBox1).Controls.Add((Control)(object)txtStartDate);
		((Control)groupBox1).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)groupBox1).Location = new Point(6, 340);
		((Control)groupBox1).Name = "groupBox1";
		((Control)groupBox1).Size = new Size(647, 60);
		((Control)groupBox1).TabIndex = 190;
		groupBox1.TabStop = false;
		((Control)groupBox1).Text = "Hiển thị dữ liệu";
		((Control)chkCanbi2).AutoSize = true;
		((Control)chkCanbi2).Location = new Point(528, 25);
		((Control)chkCanbi2).Name = "chkCanbi2";
		((Control)chkCanbi2).Size = new Size(86, 24);
		((Control)chkCanbi2).TabIndex = 166;
		((Control)chkCanbi2).Text = "Cân bì 2";
		((ButtonBase)chkCanbi2).UseVisualStyleBackColor = true;
		((Control)chkCanbi2).Visible = false;
		txtEndDate.Format = (DateTimePickerFormat)8;
		((Control)txtEndDate).Location = new Point(294, 25);
		((Control)txtEndDate).Name = "txtEndDate";
		((Control)txtEndDate).Size = new Size(127, 26);
		((Control)txtEndDate).TabIndex = 13;
		txtEndDate.ValueChanged += txtEndDate_ValueChanged;
		((Control)label9).AutoSize = true;
		((Control)label9).Location = new Point(215, 29);
		((Control)label9).Name = "label9";
		((Control)label9).Size = new Size(77, 20);
		((Control)label9).TabIndex = 164;
		((Control)label9).Text = "Đến ngày";
		((Control)label10).AutoSize = true;
		((Control)label10).Location = new Point(12, 29);
		((Control)label10).Name = "label10";
		((Control)label10).Size = new Size(65, 20);
		((Control)label10).TabIndex = 162;
		((Control)label10).Text = "Từ ngày";
		txtStartDate.Format = (DateTimePickerFormat)8;
		((Control)txtStartDate).Location = new Point(78, 25);
		((Control)txtStartDate).Name = "txtStartDate";
		((Control)txtStartDate).Size = new Size(131, 26);
		((Control)txtStartDate).TabIndex = 12;
		txtStartDate.ValueChanged += txtStartDate_ValueChanged;
		((Control)chkCanbi).AutoSize = true;
		((Control)chkCanbi).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkCanbi).Location = new Point(554, 310);
		((Control)chkCanbi).Name = "chkCanbi";
		((Control)chkCanbi).Size = new Size(73, 24);
		((Control)chkCanbi).TabIndex = 165;
		((Control)chkCanbi).Text = "Cân bì";
		((ButtonBase)chkCanbi).UseVisualStyleBackColor = true;
		((Control)chkCantudong).AutoSize = true;
		((Control)chkCantudong).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkCantudong).Location = new Point(433, 310);
		((Control)chkCantudong).Name = "chkCantudong";
		((Control)chkCantudong).Size = new Size(115, 24);
		((Control)chkCantudong).TabIndex = 208;
		((Control)chkCantudong).Text = "Cân tự động";
		((ButtonBase)chkCantudong).UseVisualStyleBackColor = true;
		chkCantudong.CheckedChanged += chkCantudong_CheckedChanged;
		((Control)panel1).Location = new Point(1, 2);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(2, 289);
		((Control)panel1).TabIndex = 191;
		((Control)label8).AutoSize = true;
		((Control)label8).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label8).Location = new Point(292, 88);
		((Control)label8).Name = "label8";
		((Control)label8).Size = new Size(72, 20);
		((Control)label8).TabIndex = 192;
		((Control)label8).Text = "Số phiếu";
		((Control)lblSophieu).BackColor = Color.White;
		((Control)lblSophieu).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblSophieu).Location = new Point(382, 87);
		((Control)lblSophieu).Name = "lblSophieu";
		((Control)lblSophieu).Size = new Size(270, 25);
		((Control)lblSophieu).TabIndex = 193;
		((Control)txtSolong).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtSolong).Location = new Point(108, 215);
		((Control)txtSolong).Name = "txtSolong";
		((Control)txtSolong).Size = new Size(77, 26);
		((Control)txtSolong).TabIndex = 8;
		((Control)txtSolong).Text = "2";
		((Control)txtSolong).Leave += txtSolong_Leave;
		((Control)label11).AutoSize = true;
		((Control)label11).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label11).Location = new Point(8, 218);
		((Control)label11).Name = "label11";
		((Control)label11).Size = new Size(63, 20);
		((Control)label11).TabIndex = 195;
		((Control)label11).Text = "Số lồng";
		((Control)panel2).BackColor = Color.FromArgb(255, 192, 192);
		((Control)panel2).Location = new Point(5, 176);
		((Control)panel2).Name = "panel2";
		((Control)panel2).Size = new Size(646, 4);
		((Control)panel2).TabIndex = 196;
		((Control)txtSocon1long).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtSocon1long).Location = new Point(289, 215);
		((Control)txtSocon1long).Name = "txtSocon1long";
		((Control)txtSocon1long).Size = new Size(96, 26);
		((Control)txtSocon1long).TabIndex = 197;
		((Control)txtSocon1long).Text = "7";
		((Control)txtSocon1long).Leave += txtSocon1long_Leave;
		((Control)label12).AutoSize = true;
		((Control)label12).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label12).Location = new Point(187, 218);
		((Control)label12).Name = "label12";
		((Control)label12).Size = new Size(102, 20);
		((Control)label12).TabIndex = 198;
		((Control)label12).Text = "Số con1 lồng";
		((Control)lblTongsolong).BackColor = Color.FromArgb(224, 224, 224);
		((Control)lblTongsolong).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblTongsolong).ForeColor = Color.Red;
		((Control)lblTongsolong).Location = new Point(793, 36);
		((Control)lblTongsolong).Name = "lblTongsolong";
		((Control)lblTongsolong).Size = new Size(118, 31);
		((Control)lblTongsolong).TabIndex = 200;
		lblTongsolong.TextAlign = (ContentAlignment)32;
		((Control)label14).AutoSize = true;
		((Control)label14).Font = new Font("Microsoft Sans Serif", 14.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label14).Location = new Point(658, 42);
		((Control)label14).Name = "label14";
		((Control)label14).Size = new Size(111, 24);
		((Control)label14).TabIndex = 199;
		((Control)label14).Text = "Số lồng gà";
		((Control)lblTongsocon).BackColor = Color.FromArgb(224, 224, 224);
		((Control)lblTongsocon).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblTongsocon).ForeColor = Color.Red;
		((Control)lblTongsocon).Location = new Point(793, 70);
		((Control)lblTongsocon).Name = "lblTongsocon";
		((Control)lblTongsocon).Size = new Size(118, 31);
		((Control)lblTongsocon).TabIndex = 202;
		lblTongsocon.TextAlign = (ContentAlignment)32;
		((Control)label16).AutoSize = true;
		((Control)label16).Font = new Font("Microsoft Sans Serif", 14.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label16).Location = new Point(658, 76);
		((Control)label16).Name = "label16";
		((Control)label16).Size = new Size(128, 24);
		((Control)label16).TabIndex = 201;
		((Control)label16).Text = "Tổng số con";
		((Control)lblTLTrungbinh).BackColor = Color.FromArgb(224, 224, 224);
		((Control)lblTLTrungbinh).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblTLTrungbinh).ForeColor = Color.Red;
		((Control)lblTLTrungbinh).Location = new Point(793, 105);
		((Control)lblTLTrungbinh).Name = "lblTLTrungbinh";
		((Control)lblTLTrungbinh).Size = new Size(118, 31);
		((Control)lblTLTrungbinh).TabIndex = 204;
		lblTLTrungbinh.TextAlign = (ContentAlignment)32;
		((Control)label15).AutoSize = true;
		((Control)label15).Font = new Font("Microsoft Sans Serif", 14.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label15).Location = new Point(658, 110);
		((Control)label15).Name = "label15";
		((Control)label15).Size = new Size(124, 24);
		((Control)label15).TabIndex = 203;
		((Control)label15).Text = "TL TB 1 con";
		dataGridTemp.AllowUserToAddRows = false;
		dataGridTemp.AllowUserToOrderColumns = true;
		dataGridTemp.ColumnHeadersHeightSizeMode = (DataGridViewColumnHeadersHeightSizeMode)2;
		dataGridTemp.Columns.AddRange((DataGridViewColumn[])(object)new DataGridViewColumn[16]
		{
			(DataGridViewColumn)macan1,
			(DataGridViewColumn)tl1,
			(DataGridViewColumn)solong1,
			(DataGridViewColumn)socon1,
			(DataGridViewColumn)macan2,
			(DataGridViewColumn)tl2,
			(DataGridViewColumn)solong2,
			(DataGridViewColumn)socon2,
			(DataGridViewColumn)macan3,
			(DataGridViewColumn)tl3,
			(DataGridViewColumn)solong3,
			(DataGridViewColumn)socon3,
			(DataGridViewColumn)macan4,
			(DataGridViewColumn)tl4,
			(DataGridViewColumn)solong4,
			(DataGridViewColumn)socon4
		});
		((Control)dataGridTemp).DataBindings.Add(new Binding("Font", (object)Settings.Default, "font", true, (DataSourceUpdateMode)1));
		((Control)dataGridTemp).Font = Settings.Default.font;
		((Control)dataGridTemp).Location = new Point(22, 469);
		((Control)dataGridTemp).Name = "dataGridTemp";
		dataGridTemp.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)dataGridTemp).Size = new Size(603, 141);
		((Control)dataGridTemp).TabIndex = 205;
		((Control)dataGridTemp).Visible = false;
		((DataGridViewColumn)macan1).HeaderText = "macan1";
		((DataGridViewColumn)macan1).Name = "macan1";
		((DataGridViewColumn)tl1).HeaderText = "tl1";
		((DataGridViewColumn)tl1).Name = "tl1";
		((DataGridViewColumn)solong1).HeaderText = "solong1";
		((DataGridViewColumn)solong1).Name = "solong1";
		((DataGridViewColumn)socon1).HeaderText = "socon1";
		((DataGridViewColumn)socon1).Name = "socon1";
		((DataGridViewColumn)macan2).HeaderText = "macan2";
		((DataGridViewColumn)macan2).Name = "macan2";
		((DataGridViewColumn)tl2).HeaderText = "tl2";
		((DataGridViewColumn)tl2).Name = "tl2";
		((DataGridViewColumn)solong2).HeaderText = "solong2";
		((DataGridViewColumn)solong2).Name = "solong2";
		((DataGridViewColumn)socon2).HeaderText = "socon2";
		((DataGridViewColumn)socon2).Name = "socon2";
		((DataGridViewColumn)macan3).HeaderText = "macan3";
		((DataGridViewColumn)macan3).Name = "macan3";
		((DataGridViewColumn)tl3).HeaderText = "tl3";
		((DataGridViewColumn)tl3).Name = "tl3";
		((DataGridViewColumn)solong3).HeaderText = "solong3";
		((DataGridViewColumn)solong3).Name = "solong3";
		((DataGridViewColumn)socon3).HeaderText = "socon3";
		((DataGridViewColumn)socon3).Name = "socon3";
		((DataGridViewColumn)macan4).HeaderText = "macan4";
		((DataGridViewColumn)macan4).Name = "macan4";
		((DataGridViewColumn)tl4).HeaderText = "tl4";
		((DataGridViewColumn)tl4).Name = "tl4";
		((DataGridViewColumn)solong4).HeaderText = "solong4";
		((DataGridViewColumn)solong4).Name = "solong4";
		((DataGridViewColumn)socon4).HeaderText = "socon4";
		((DataGridViewColumn)socon4).Name = "socon4";
		((Control)lblOndinhKytu).AutoSize = true;
		((Control)lblOndinhKytu).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblOndinhKytu).Location = new Point(169, 93);
		((Control)lblOndinhKytu).Name = "lblOndinhKytu";
		((Control)lblOndinhKytu).Size = new Size(98, 20);
		((Control)lblOndinhKytu).TabIndex = 206;
		((Control)lblOndinhKytu).Text = "ky tu on dinh";
		((Control)lblOndinhKytu).Visible = false;
		((Control)groupBox2).Controls.Add((Control)(object)label18);
		((Control)groupBox2).Controls.Add((Control)(object)cboCongCom1);
		((Control)groupBox2).Controls.Add((Control)(object)label13);
		((Control)groupBox2).Controls.Add((Control)(object)txtTLLuutu);
		((Control)groupBox2).Controls.Add((Control)(object)label17);
		((Control)groupBox2).Controls.Add((Control)(object)txtTGOndinh);
		((Control)groupBox2).Location = new Point(716, 230);
		((Control)groupBox2).Name = "groupBox2";
		((Control)groupBox2).Size = new Size(291, 130);
		((Control)groupBox2).TabIndex = 207;
		groupBox2.TabStop = false;
		((Control)groupBox2).Text = "Cân số 1 (xanh)";
		((Control)groupBox2).Visible = false;
		((Control)label18).AutoSize = true;
		((Control)label18).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label18).Location = new Point(14, 92);
		((Control)label18).Name = "label18";
		((Control)label18).Size = new Size(118, 25);
		((Control)label18).TabIndex = 176;
		((Control)label18).Text = "Cổng COM";
		cboCongCom1.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboCongCom1).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboCongCom1).FormattingEnabled = true;
		cboCongCom1.Items.AddRange(new object[10] { "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "COM10" });
		((Control)cboCongCom1).Location = new Point(160, 89);
		((Control)cboCongCom1).Name = "cboCongCom1";
		((Control)cboCongCom1).Size = new Size(106, 28);
		((Control)cboCongCom1).TabIndex = 175;
		((Control)label13).AutoSize = true;
		((Control)label13).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label13).Location = new Point(14, 55);
		((Control)label13).Name = "label13";
		((Control)label13).Size = new Size(139, 25);
		((Control)label13).TabIndex = 3;
		((Control)label13).Text = "TL lưu từ (kg)";
		((Control)txtTLLuutu).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtTLLuutu).Location = new Point(160, 52);
		((Control)txtTLLuutu).Name = "txtTLLuutu";
		((Control)txtTLLuutu).Size = new Size(106, 31);
		((Control)txtTLLuutu).TabIndex = 2;
		((Control)txtTLLuutu).Text = "0.3";
		((Control)label17).AutoSize = true;
		((Control)label17).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label17).Location = new Point(14, 18);
		((Control)label17).Name = "label17";
		((Control)label17).Size = new Size(118, 25);
		((Control)label17).TabIndex = 1;
		((Control)label17).Text = "TG ổn định";
		((Control)txtTGOndinh).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtTGOndinh).Location = new Point(160, 15);
		((Control)txtTGOndinh).Name = "txtTGOndinh";
		((Control)txtTGOndinh).Size = new Size(106, 31);
		((Control)txtTGOndinh).TabIndex = 0;
		((Control)txtTGOndinh).Text = "1.5";
		((Control)lblDaucan2).BackColor = Color.Black;
		((Control)lblDaucan2).Font = new Font("Arial", 48f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblDaucan2).ForeColor = Color.Red;
		((Control)lblDaucan2).Location = new Point(703, 169);
		((Control)lblDaucan2).Name = "lblDaucan2";
		((Control)lblDaucan2).Size = new Size(316, 72);
		((Control)lblDaucan2).TabIndex = 208;
		((Control)lblDaucan2).Text = "1";
		lblDaucan2.TextAlign = (ContentAlignment)32;
		((Control)lblDaucan2).Visible = false;
		((Control)lblDaucan2).DoubleClick += lblDaucan2_DoubleClick;
		((Control)txtSoluong2).Font = new Font("Arial", 20.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)txtSoluong2).Location = new Point(708, 202);
		((Control)txtSoluong2).Name = "txtSoluong2";
		((Control)txtSoluong2).Size = new Size(74, 39);
		((Control)txtSoluong2).TabIndex = 209;
		((Control)txtSoluong2).Text = "0";
		txtSoluong2.TextAlign = (HorizontalAlignment)2;
		((Control)txtSoluong2).Visible = false;
		((Control)txtSoluong2).TextChanged += txtSoluong2_TextChanged;
		((Control)groupBox3).Controls.Add((Control)(object)label19);
		((Control)groupBox3).Controls.Add((Control)(object)cboCongCom2);
		((Control)groupBox3).Controls.Add((Control)(object)label20);
		((Control)groupBox3).Controls.Add((Control)(object)txtTLLuutu2);
		((Control)groupBox3).Controls.Add((Control)(object)label21);
		((Control)groupBox3).Controls.Add((Control)(object)txtTGOndinh2);
		((Control)groupBox3).Location = new Point(716, 366);
		((Control)groupBox3).Name = "groupBox3";
		((Control)groupBox3).Size = new Size(291, 130);
		((Control)groupBox3).TabIndex = 208;
		groupBox3.TabStop = false;
		((Control)groupBox3).Text = "Cân số 2 - Cân bì (đỏ)";
		((Control)groupBox3).Visible = false;
		((Control)label19).AutoSize = true;
		((Control)label19).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label19).Location = new Point(14, 92);
		((Control)label19).Name = "label19";
		((Control)label19).Size = new Size(118, 25);
		((Control)label19).TabIndex = 176;
		((Control)label19).Text = "Cổng COM";
		cboCongCom2.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboCongCom2).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboCongCom2).FormattingEnabled = true;
		cboCongCom2.Items.AddRange(new object[10] { "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "COM10" });
		((Control)cboCongCom2).Location = new Point(160, 89);
		((Control)cboCongCom2).Name = "cboCongCom2";
		((Control)cboCongCom2).Size = new Size(106, 28);
		((Control)cboCongCom2).TabIndex = 175;
		((Control)label20).AutoSize = true;
		((Control)label20).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label20).Location = new Point(14, 55);
		((Control)label20).Name = "label20";
		((Control)label20).Size = new Size(139, 25);
		((Control)label20).TabIndex = 3;
		((Control)label20).Text = "TL lưu từ (kg)";
		((Control)txtTLLuutu2).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtTLLuutu2).Location = new Point(160, 52);
		((Control)txtTLLuutu2).Name = "txtTLLuutu2";
		((Control)txtTLLuutu2).Size = new Size(106, 31);
		((Control)txtTLLuutu2).TabIndex = 2;
		((Control)txtTLLuutu2).Text = "0.3";
		((Control)label21).AutoSize = true;
		((Control)label21).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label21).Location = new Point(14, 18);
		((Control)label21).Name = "label21";
		((Control)label21).Size = new Size(118, 25);
		((Control)label21).TabIndex = 1;
		((Control)label21).Text = "TG ổn định";
		((Control)txtTGOndinh2).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtTGOndinh2).Location = new Point(160, 15);
		((Control)txtTGOndinh2).Name = "txtTGOndinh2";
		((Control)txtTGOndinh2).Size = new Size(106, 31);
		((Control)txtTGOndinh2).TabIndex = 0;
		((Control)txtTGOndinh2).Text = "1.5";
		timerChekOndinh2.Interval = 500;
		timerChekOndinh2.Tick += timerChekOndinh2_Tick;
		timer2.Interval = 200;
		timer2.Tick += timer2_Tick;
		((Control)chkCantudong2).AutoSize = true;
		((Control)chkCantudong2).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkCantudong2).Location = new Point(929, 200);
		((Control)chkCantudong2).Name = "chkCantudong2";
		((Control)chkCantudong2).Size = new Size(228, 24);
		((Control)chkCantudong2).TabIndex = 209;
		((Control)chkCantudong2).Text = "Cân tự động 2 (Cân màu đỏ)";
		((ButtonBase)chkCantudong2).UseVisualStyleBackColor = true;
		((Control)chkCantudong2).Visible = false;
		chkCantudong2.CheckedChanged += chkCantudong2_CheckedChanged;
		((Control)txtSolongbi).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtSolongbi).Location = new Point(534, 218);
		((Control)txtSolongbi).Name = "txtSolongbi";
		((Control)txtSolongbi).Size = new Size(117, 26);
		((Control)txtSolongbi).TabIndex = 211;
		((Control)txtSolongbi).Text = "0";
		((Control)label22).AutoSize = true;
		((Control)label22).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label22).Location = new Point(419, 221);
		((Control)label22).Name = "label22";
		((Control)label22).Size = new Size(109, 20);
		((Control)label22).TabIndex = 212;
		((Control)label22).Text = "Số lồng cân bì";
		((Control)lblSolancanbi).BackColor = Color.FromArgb(224, 224, 224);
		((Control)lblSolancanbi).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblSolancanbi).ForeColor = Color.Red;
		((Control)lblSolancanbi).Location = new Point(1093, 36);
		((Control)lblSolancanbi).Name = "lblSolancanbi";
		((Control)lblSolancanbi).Size = new Size(97, 31);
		((Control)lblSolancanbi).TabIndex = 214;
		lblSolancanbi.TextAlign = (ContentAlignment)32;
		((Control)label24).AutoSize = true;
		((Control)label24).Font = new Font("Microsoft Sans Serif", 14.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label24).Location = new Point(925, 42);
		((Control)label24).Name = "label24";
		((Control)label24).Size = new Size(132, 24);
		((Control)label24).TabIndex = 213;
		((Control)label24).Text = "Số lần cân bì";
		((Control)lblSolongcanbi).BackColor = Color.FromArgb(224, 224, 224);
		((Control)lblSolongcanbi).Font = new Font("Microsoft Sans Serif", 15.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblSolongcanbi).ForeColor = Color.Red;
		((Control)lblSolongcanbi).Location = new Point(1094, 70);
		((Control)lblSolongcanbi).Name = "lblSolongcanbi";
		((Control)lblSolongcanbi).Size = new Size(97, 31);
		((Control)lblSolongcanbi).TabIndex = 216;
		lblSolongcanbi.TextAlign = (ContentAlignment)32;
		((Control)label25).AutoSize = true;
		((Control)label25).Font = new Font("Microsoft Sans Serif", 14.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label25).Location = new Point(926, 76);
		((Control)label25).Name = "label25";
		((Control)label25).Size = new Size(145, 24);
		((Control)label25).TabIndex = 215;
		((Control)label25).Text = "Số lồng cân bì";
		((Control)btnSave2).BackgroundImageLayout = (ImageLayout)0;
		((Control)btnSave2).Enabled = false;
		((ButtonBase)btnSave2).FlatAppearance.BorderSize = 0;
		((Control)btnSave2).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave2).Image = (Image)(object)Resources.save_icon_red;
		((ButtonBase)btnSave2).ImageAlign = (ContentAlignment)16;
		((Control)btnSave2).Location = new Point(988, 230);
		((Control)btnSave2).Name = "btnSave2";
		((Control)btnSave2).Size = new Size(197, 42);
		((Control)btnSave2).TabIndex = 210;
		((Control)btnSave2).Text = "       Lưu  DL cân 2 [F2]";
		((ButtonBase)btnSave2).UseVisualStyleBackColor = true;
		((Control)btnSave2).Visible = false;
		((Control)btnSave2).Click += btnSave2_Click;
		((Control)btnInphieu).BackgroundImageLayout = (ImageLayout)0;
		((Control)btnInphieu).Enabled = false;
		((ButtonBase)btnInphieu).FlatAppearance.BorderSize = 0;
		((Control)btnInphieu).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnInphieu).Image = (Image)(object)Resources.printer_icon;
		((ButtonBase)btnInphieu).ImageAlign = (ContentAlignment)16;
		((Control)btnInphieu).Location = new Point(230, 295);
		((Control)btnInphieu).Name = "btnInphieu";
		((Control)btnInphieu).Size = new Size(197, 42);
		((Control)btnInphieu).TabIndex = 10;
		((Control)btnInphieu).Text = "       In phiếu";
		((ButtonBase)btnInphieu).UseVisualStyleBackColor = true;
		((Control)btnInphieu).Click += btnInphieu_Click;
		((Control)btnTaophieu).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnTaophieu).FlatAppearance.BorderSize = 0;
		((Control)btnTaophieu).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnTaophieu).Image = (Image)(object)Resources.layer_create_icon;
		((ButtonBase)btnTaophieu).ImageAlign = (ContentAlignment)16;
		((Control)btnTaophieu).Location = new Point(27, 249);
		((Control)btnTaophieu).Name = "btnTaophieu";
		((Control)btnTaophieu).Size = new Size(197, 42);
		((Control)btnTaophieu).TabIndex = 8;
		((Control)btnTaophieu).Text = "       Tạo phiếu cân";
		((ButtonBase)btnTaophieu).UseVisualStyleBackColor = true;
		((Control)btnTaophieu).Click += btnTaophieu_Click;
		((Control)btnDelete).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnDelete).FlatAppearance.BorderSize = 0;
		((Control)btnDelete).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnDelete).Image = (Image)(object)Resources.delete;
		((ButtonBase)btnDelete).ImageAlign = (ContentAlignment)16;
		((Control)btnDelete).Location = new Point(1092, 104);
		((Control)btnDelete).Name = "btnDelete";
		((Control)btnDelete).Size = new Size(101, 35);
		((Control)btnDelete).TabIndex = 14;
		((Control)btnDelete).Text = "       Xóa";
		((ButtonBase)btnDelete).UseVisualStyleBackColor = true;
		((Control)btnDelete).Click += btnDelete_Click;
		((Control)btnKetthuc).BackgroundImageLayout = (ImageLayout)0;
		((Control)btnKetthuc).Enabled = false;
		((ButtonBase)btnKetthuc).FlatAppearance.BorderSize = 0;
		((Control)btnKetthuc).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnKetthuc).Image = (Image)(object)Resources.log_in;
		((ButtonBase)btnKetthuc).ImageAlign = (ContentAlignment)16;
		((Control)btnKetthuc).Location = new Point(27, 295);
		((Control)btnKetthuc).Name = "btnKetthuc";
		((Control)btnKetthuc).Size = new Size(197, 42);
		((Control)btnKetthuc).TabIndex = 9;
		((Control)btnKetthuc).Text = "       Xong phiếu";
		((ButtonBase)btnKetthuc).UseVisualStyleBackColor = true;
		((Control)btnKetthuc).Click += btnKetthuc_Click;
		((Control)btnSave).BackgroundImageLayout = (ImageLayout)0;
		((Control)btnSave).Enabled = false;
		((ButtonBase)btnSave).FlatAppearance.BorderSize = 0;
		((Control)btnSave).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave).Image = (Image)(object)Resources.Save1;
		((ButtonBase)btnSave).ImageAlign = (ContentAlignment)16;
		((Control)btnSave).Location = new Point(230, 249);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(197, 42);
		((Control)btnSave).TabIndex = 11;
		((Control)btnSave).Text = "       Lưu  DL cân [F1]";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(1195, 632);
		((Control)this).Controls.Add((Control)(object)lblSolongcanbi);
		((Control)this).Controls.Add((Control)(object)label25);
		((Control)this).Controls.Add((Control)(object)lblSolancanbi);
		((Control)this).Controls.Add((Control)(object)label24);
		((Control)this).Controls.Add((Control)(object)chkCantudong2);
		((Control)this).Controls.Add((Control)(object)chkCanbi);
		((Control)this).Controls.Add((Control)(object)txtSolongbi);
		((Control)this).Controls.Add((Control)(object)chkCantudong);
		((Control)this).Controls.Add((Control)(object)label22);
		((Control)this).Controls.Add((Control)(object)btnSave2);
		((Control)this).Controls.Add((Control)(object)groupBox3);
		((Control)this).Controls.Add((Control)(object)txtSoluong2);
		((Control)this).Controls.Add((Control)(object)lblOndinhKytu);
		((Control)this).Controls.Add((Control)(object)lblDaucan2);
		((Control)this).Controls.Add((Control)(object)groupBox2);
		((Control)this).Controls.Add((Control)(object)dataGridTemp);
		((Control)this).Controls.Add((Control)(object)lblTLTrungbinh);
		((Control)this).Controls.Add((Control)(object)label15);
		((Control)this).Controls.Add((Control)(object)lblTongsocon);
		((Control)this).Controls.Add((Control)(object)label16);
		((Control)this).Controls.Add((Control)(object)lblTongsolong);
		((Control)this).Controls.Add((Control)(object)label14);
		((Control)this).Controls.Add((Control)(object)txtSocon1long);
		((Control)this).Controls.Add((Control)(object)label12);
		((Control)this).Controls.Add((Control)(object)panel2);
		((Control)this).Controls.Add((Control)(object)txtSolong);
		((Control)this).Controls.Add((Control)(object)label11);
		((Control)this).Controls.Add((Control)(object)lblSophieu);
		((Control)this).Controls.Add((Control)(object)label8);
		((Control)this).Controls.Add((Control)(object)btnInphieu);
		((Control)this).Controls.Add((Control)(object)panel1);
		((Control)this).Controls.Add((Control)(object)groupBox1);
		((Control)this).Controls.Add((Control)(object)btnTaophieu);
		((Control)this).Controls.Add((Control)(object)dataGridPhieucan);
		((Control)this).Controls.Add((Control)(object)btnDelete);
		((Control)this).Controls.Add((Control)(object)lblSolancan);
		((Control)this).Controls.Add((Control)(object)lblTotal);
		((Control)this).Controls.Add((Control)(object)label7);
		((Control)this).Controls.Add((Control)(object)label6);
		((Control)this).Controls.Add((Control)(object)btnKetthuc);
		((Control)this).Controls.Add((Control)(object)cboXuatnhap);
		((Control)this).Controls.Add((Control)(object)label5);
		((Control)this).Controls.Add((Control)(object)cboSoxe);
		((Control)this).Controls.Add((Control)(object)label4);
		((Control)this).Controls.Add((Control)(object)txtNguoinhan);
		((Control)this).Controls.Add((Control)(object)label3);
		((Control)this).Controls.Add((Control)(object)cboTenKH);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)cboMaKH);
		((Control)this).Controls.Add((Control)(object)cboTenhang);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)cboMahang);
		((Control)this).Controls.Add((Control)(object)txtSoluong);
		((Control)this).Controls.Add((Control)(object)lblDaucan);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Controls.Add((Control)(object)dataGrid);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).KeyPreview = true;
		((Form)this).Margin = new Padding(4);
		((Control)this).Name = "frmCan";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Cân gà";
		((Form)this).WindowState = (FormWindowState)2;
		((Form)this).Load += frmCan_Load;
		((Control)this).DoubleClick += frmCan_DoubleClick;
		((Form)this).FormClosed += new FormClosedEventHandler(frmCan_FormClosed);
		((Control)this).Resize += frmCan_Resize;
		((Control)this).KeyDown += new KeyEventHandler(frmCan_KeyDown);
		((ISupportInitialize)dataGrid).EndInit();
		((ISupportInitialize)dataGridPhieucan).EndInit();
		((Control)groupBox1).ResumeLayout(false);
		((Control)groupBox1).PerformLayout();
		((ISupportInitialize)dataGridTemp).EndInit();
		((Control)groupBox2).ResumeLayout(false);
		((Control)groupBox2).PerformLayout();
		((Control)groupBox3).ResumeLayout(false);
		((Control)groupBox3).PerformLayout();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
