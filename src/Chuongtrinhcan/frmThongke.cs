using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Label = System.Windows.Forms.Label;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using Timer = System.Windows.Forms.Timer;
using Point = System.Drawing.Point;
using Font = System.Drawing.Font;
using Chuongtrinhcan.Properties;
using Microsoft.Office.Interop.Excel;

namespace Chuongtrinhcan;

public class frmThongke : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private DataFetcher dataFetcher2 = DataFetcher.Instance;

	private string excelName;

	private IContainer components = null;

	private Label label1;

	private Button btnThoat;

	private Button btnXem;

	private Button btnExcel;

	private SaveFileDialog saveFileDialog1;

	private DataGridView dataGrid;

	private CheckBox chkNhanvien;

	private ComboBox cboNhanvien;

	private CheckBox chkSanpham;

	private ComboBox cboSanpham;

	private DateTimePicker txtEndDate;

	private Label label10;

	private Label label9;

	private DateTimePicker txtStartDate;

	private CheckBox chkKhachhang;

	private ComboBox cboKhachhang;

	private CheckBox chkXuatnhap;

	private ComboBox cboXuatnhap;

	private CheckBox chkGhichu;

	private ComboBox cboGhichu;

	private CheckBox chkSophieu;

	private ComboBox cboSophieu;

	private Button btnXemTong;

	private Panel panel1;

	private ComboBox cboSoxe;

	private CheckBox chkSoxe;

	private ComboBox cboHinhthuccan;

	private CheckBox chkHinhthuccan;

	private DataGridViewTextBoxColumn id;

	private DataGridViewTextBoxColumn STT2;

	private DataGridViewTextBoxColumn sophieu;

	private DataGridViewTextBoxColumn soxe;

	private DataGridViewTextBoxColumn ma;

	private DataGridViewTextBoxColumn ten;

	private DataGridViewTextBoxColumn makh;

	private DataGridViewTextBoxColumn tenkh;

	private DataGridViewTextBoxColumn tlcan;

	private DataGridViewTextBoxColumn solong;

	private DataGridViewTextBoxColumn socon1long;

	private DataGridViewTextBoxColumn tongsocon;

	private DataGridViewTextBoxColumn tlbi;

	private DataGridViewTextBoxColumn solongbi;

	private DataGridViewTextBoxColumn ghichu;

	private DataGridViewTextBoxColumn ngaycan;

	private DataGridViewTextBoxColumn nhanvien;

	private DataGridViewTextBoxColumn loaican;

	public frmThongke()
	{
		InitializeComponent();
	}

	private void frmThongke_Load(object sender, EventArgs e)
	{
		DBUtil.LoadMaunen((Form)(object)this);
		txtStartDate.Value = DateTime.Now;
		txtStartDate.CustomFormat = "dd/MM/yyyy";
		txtEndDate.Value = DateTime.Now;
		txtEndDate.CustomFormat = "dd/MM/yyyy";
		string text = "";
		text = "select nv_id,nv_id & ' - ' & tennv as fullname from nhanvien order by nv_id";
		DBUtil.loadComboValue(cboNhanvien, text, "fullname", "nv_id");
		text = "select id,mahang & ' - ' & tenhang as loaihang from hanghoa order by mahang";
		DBUtil.loadComboValue(cboSanpham, text, "loaihang", "id");
		text = "select id,makh & ' - ' & tenkh as loaihang from khachhang order by makh";
		DBUtil.loadComboValue(cboKhachhang, text, "loaihang", "id");
		text = "select distinct ghichu from phieucan where ghichu <>'' order by ghichu";
		DBUtil.loadCombo(cboGhichu, text, "ghichu");
		text = "select distinct sophieu from phieucan order by sophieu";
		DBUtil.loadCombo(cboSophieu, text, "sophieu");
		text = "select distinct xuatnhap from phieucan where xuatnhap <>'' order by xuatnhap";
		DBUtil.loadCombo(cboXuatnhap, text, "xuatnhap");
		text = "select distinct soxe from phieucan where soxe <>'' order by soxe";
		DBUtil.loadCombo(cboSoxe, text, "soxe");
		((ListControl)cboNhanvien).SelectedIndex = -1;
		((ListControl)cboSanpham).SelectedIndex = -1;
		((ListControl)cboKhachhang).SelectedIndex = -1;
		((Control)cboXuatnhap).Text = "";
		((Control)cboGhichu).Text = "";
		((Control)cboSophieu).Text = "";
		((Control)cboSoxe).Text = "";
	}

	private void btnThoat_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void btnXem_Click(object sender, EventArgs e)
	{
		XemTKChitiet();
	}

	private void XemTKChitiet()
	{
		try
		{
			dataGrid.RowCount = 0;
			string text = "";
			string text2 = "";
			text = "select stt,sophieu,soxe,mahang,tenhang,makh,tenkh,xuatnhap,ghichu,ngaycanct, ";
			text += " iif(loaican=0,trongluong,0) as tlga,iif(loaican=1,trongluong,0) as tlbi,tennv,solong,socon1long,solong*socon1long as tongsocon, ";
			text += " iif(loaican=0,solong,0) as solongga,iif(loaican=1,solong,0) as solongbi,";
			text += " iif(loaican=0,'Cân gà','Cân bì') as hinhthuccan,loaican ";
			text += " from hanghoa h, khachhang k, phieucan p, chitietphieucan c, nhanvien n ";
			text += "where h.id=c.hang_id and n.nv_id=p.nv_id and k.id=p.kh_id and p.stt=c.phieu_id ";
			string text3 = text;
			text = text3 + " and ngaycanct >=#" + txtStartDate.Value.ToString("MM/dd/yyyy") + " 00:00:01# and ngaycanct <= #" + txtEndDate.Value.ToString("MM/dd/yyyy") + " 23:59:59# ";
			if (chkNhanvien.Checked)
			{
				object obj = text2;
				text2 = string.Concat(obj, " and n.nv_id='", ((ListControl)cboNhanvien).SelectedValue, "' ");
			}
			if (chkSanpham.Checked)
			{
				object obj = text2;
				text2 = string.Concat(obj, " and c.hang_id=", ((ListControl)cboSanpham).SelectedValue, " ");
			}
			if (chkKhachhang.Checked)
			{
				object obj = text2;
				text2 = string.Concat(obj, " and p.kh_id=", ((ListControl)cboKhachhang).SelectedValue, " ");
			}
			if (chkXuatnhap.Checked)
			{
				text2 = text2 + " and p.xuatnhap='" + ((Control)cboXuatnhap).Text + "' ";
			}
			if (chkGhichu.Checked)
			{
				text2 = text2 + " and p.ghichu='" + ((Control)cboGhichu).Text + "' ";
			}
			if (chkSophieu.Checked)
			{
				text2 = text2 + " and p.sophieu=" + ((Control)cboSophieu).Text;
			}
			if (chkSoxe.Checked)
			{
				text2 = text2 + " and soxe='" + ((Control)cboSoxe).Text + "'";
			}
			if (chkHinhthuccan.Checked)
			{
				text2 = text2 + " and loaican=" + ((ListControl)cboHinhthuccan).SelectedIndex;
			}
			text += text2;
			text += " order by sophieu,loaican,ngaycanct ";
			dataFetcher.setCommand(text, 1);
			DataSet dataSet = dataFetcher.fetchDataSet();
			int num = 0;
			double num2 = 0.0;
			double num3 = 0.0;
			double num4 = 0.0;
			double num5 = 0.0;
			double num6 = 0.0;
			foreach (DataRow row in dataSet.Tables[0].Rows)
			{
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = row["stt"];
				dataGrid.Rows[num].Cells[1].Value = (num + 1).ToString();
				dataGrid.Rows[num].Cells[2].Value = row["sophieu"];
				dataGrid.Rows[num].Cells[3].Value = row["soxe"];
				dataGrid.Rows[num].Cells[4].Value = row["mahang"];
				dataGrid.Rows[num].Cells[5].Value = row["tenhang"];
				dataGrid.Rows[num].Cells[6].Value = row["makh"];
				dataGrid.Rows[num].Cells[7].Value = row["tenkh"];
				dataGrid.Rows[num].Cells[8].Value = row["tlga"];
				dataGrid.Rows[num].Cells[9].Value = row["solongga"];
				dataGrid.Rows[num].Cells[10].Value = row["socon1long"];
				dataGrid.Rows[num].Cells[11].Value = row["tongsocon"];
				dataGrid.Rows[num].Cells[12].Value = row["tlbi"];
				dataGrid.Rows[num].Cells[13].Value = row["solongbi"];
				dataGrid.Rows[num].Cells[14].Value = row["ghichu"];
				dataGrid.Rows[num].Cells[15].Value = row["ngaycanct"];
				dataGrid.Rows[num].Cells[16].Value = row["tennv"];
				dataGrid.Rows[num].Cells[17].Value = row["hinhthuccan"];
				num2 += double.Parse(row["tlga"].ToString());
				num3 += double.Parse(row["solongga"].ToString());
				num4 += double.Parse(row["tongsocon"].ToString());
				num5 += double.Parse(row["tlbi"].ToString());
				num6 += double.Parse(row["solongbi"].ToString());
			}
			if (dataGrid.Rows.Count >= 2)
			{
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = "";
				dataGrid.Rows[num].Cells[1].Value = "";
				dataGrid.Rows[num].Cells[2].Value = "";
				dataGrid.Rows[num].Cells[3].Value = "";
				dataGrid.Rows[num].Cells[4].Value = "";
				dataGrid.Rows[num].Cells[5].Value = "";
				dataGrid.Rows[num].Cells[6].Value = "";
				dataGrid.Rows[num].Cells[7].Value = "TỔNG";
				dataGrid.Rows[num].Cells[8].Value = num2.ToString("#,##0.0");
				dataGrid.Rows[num].Cells[9].Value = num3.ToString();
				dataGrid.Rows[num].Cells[10].Value = "";
				dataGrid.Rows[num].Cells[11].Value = num4.ToString();
				dataGrid.Rows[num].Cells[12].Value = num5.ToString("#,##0.0");
				dataGrid.Rows[num].Cells[13].Value = num6.ToString();
				dataGrid.Rows[num].Cells[14].Value = "";
				dataGrid.Rows[num].Cells[15].Value = "";
				dataGrid.Rows[num].Cells[16].Value = "";
				dataGrid.Rows[num].Cells[17].Value = "";
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = "";
				dataGrid.Rows[num].Cells[1].Value = "";
				dataGrid.Rows[num].Cells[2].Value = "";
				dataGrid.Rows[num].Cells[3].Value = "";
				dataGrid.Rows[num].Cells[4].Value = "";
				dataGrid.Rows[num].Cells[5].Value = "";
				dataGrid.Rows[num].Cells[6].Value = "";
				dataGrid.Rows[num].Cells[7].Value = "TỔNG TL GÀ";
				dataGrid.Rows[num].Cells[8].Value = (num2 - num5).ToString("#,##0.0");
				dataGrid.Rows[num].Cells[9].Value = "TL TB 1 CON";
				dataGrid.Rows[num].Cells[10].Value = Math.Round((num2 - num5 / num6 * num3) / num4, 3).ToString("#,##0.000");
				dataGrid.Rows[num].Cells[11].Value = "";
				dataGrid.Rows[num].Cells[12].Value = "";
				dataGrid.Rows[num].Cells[13].Value = "";
				dataGrid.Rows[num].Cells[14].Value = "";
				dataGrid.Rows[num].Cells[15].Value = "";
				dataGrid.Rows[num].Cells[16].Value = "";
				dataGrid.Rows[num].Cells[17].Value = "";
			}
		}
		catch
		{
		}
		formatDataGrid();
		dataGrid.AutoResizeColumns();
	}

	private void btnXemTong_Click(object sender, EventArgs e)
	{
		try
		{
			dataGrid.RowCount = 0;
			string text = "";
			string text2 = "";
			text = "select stt,sophieu,soxe,mahang,tenhang,makh,tenkh,ghichu,xuatnhap,ngaycan,";
			text += " iif(loaican=0,sum(trongluong),0) as tongtlga,iif(loaican=1,sum(trongluong),0) as tongtlbi,";
			text += " iif(loaican=0,sum(solong),0) as tongsolongga,iif(loaican=1,sum(solong),0) as tongsolongbi,";
			text += " sum(solong*socon1long) as tongsocon,round(avg(socon1long),1) as soconTB,tennv, ";
			text += " iif(loaican=0,'Cân gà','Cân bì') as hinhthuccan ";
			text += " from hanghoa h, khachhang k, phieucan p, chitietphieucan c, nhanvien n ";
			text += "where h.id=c.hang_id and n.nv_id=p.nv_id and k.id=p.kh_id and p.stt=c.phieu_id ";
			string text3 = text;
			text = text3 + " and ngaycan >=#" + txtStartDate.Value.ToString("MM/dd/yyyy") + " 00:00:01# and ngaycan <= #" + txtEndDate.Value.ToString("MM/dd/yyyy") + " 23:59:59# ";
			if (chkNhanvien.Checked)
			{
				object obj = text2;
				text2 = string.Concat(obj, " and n.nv_id='", ((ListControl)cboNhanvien).SelectedValue, "' ");
			}
			if (chkSanpham.Checked)
			{
				object obj = text2;
				text2 = string.Concat(obj, " and c.hang_id=", ((ListControl)cboSanpham).SelectedValue, " ");
			}
			if (chkKhachhang.Checked)
			{
				object obj = text2;
				text2 = string.Concat(obj, " and p.kh_id=", ((ListControl)cboKhachhang).SelectedValue, " ");
			}
			if (chkXuatnhap.Checked)
			{
				text2 = text2 + " and p.xuatnhap='" + ((Control)cboXuatnhap).Text + "' ";
			}
			if (chkGhichu.Checked)
			{
				text2 = text2 + " and p.ghichu='" + ((Control)cboGhichu).Text + "' ";
			}
			if (chkSophieu.Checked)
			{
				text2 = text2 + " and p.sophieu=" + ((Control)cboSophieu).Text;
			}
			if (chkSoxe.Checked)
			{
				text2 = text2 + " and soxe='" + ((Control)cboSoxe).Text + "'";
			}
			if (chkHinhthuccan.Checked)
			{
				text2 = text2 + " and loaican=" + ((ListControl)cboHinhthuccan).SelectedIndex;
			}
			text += text2;
			text += " group by stt,soxe,sophieu,mahang,tenhang,makh,tenkh,ghichu,xuatnhap,ngaycan,tennv,loaican";
			text += " order by ngaycan,sophieu ";
			dataFetcher.setCommand(text, 1);
			DataSet dataSet = dataFetcher.fetchDataSet();
			int num = 0;
			double num2 = 0.0;
			double num3 = 0.0;
			double num4 = 0.0;
			double num5 = 0.0;
			double num6 = 0.0;
			foreach (DataRow row in dataSet.Tables[0].Rows)
			{
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = row["stt"];
				dataGrid.Rows[num].Cells[1].Value = (num + 1).ToString();
				dataGrid.Rows[num].Cells[2].Value = row["sophieu"];
				dataGrid.Rows[num].Cells[3].Value = row["soxe"];
				dataGrid.Rows[num].Cells[4].Value = row["mahang"];
				dataGrid.Rows[num].Cells[5].Value = row["tenhang"];
				dataGrid.Rows[num].Cells[6].Value = row["makh"];
				dataGrid.Rows[num].Cells[7].Value = row["tenkh"];
				dataGrid.Rows[num].Cells[8].Value = row["tongtlga"];
				dataGrid.Rows[num].Cells[9].Value = row["tongsolongga"];
				dataGrid.Rows[num].Cells[10].Value = row["socontb"];
				dataGrid.Rows[num].Cells[11].Value = row["tongsocon"];
				dataGrid.Rows[num].Cells[12].Value = row["tongtlbi"];
				dataGrid.Rows[num].Cells[13].Value = row["tongsolongbi"];
				dataGrid.Rows[num].Cells[14].Value = row["ghichu"];
				dataGrid.Rows[num].Cells[15].Value = row["ngaycan"];
				dataGrid.Rows[num].Cells[16].Value = row["tennv"];
				dataGrid.Rows[num].Cells[17].Value = row["hinhthuccan"];
				num2 += double.Parse(row["tongtlga"].ToString());
				num3 += double.Parse(row["tongsocon"].ToString());
				num4 += double.Parse(row["tongsolongga"].ToString());
				num5 += double.Parse(row["tongtlbi"].ToString());
				num6 += double.Parse(row["tongsolongbi"].ToString());
			}
			if (dataGrid.Rows.Count >= 2)
			{
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = "";
				dataGrid.Rows[num].Cells[1].Value = "";
				dataGrid.Rows[num].Cells[2].Value = "";
				dataGrid.Rows[num].Cells[3].Value = "";
				dataGrid.Rows[num].Cells[4].Value = "";
				dataGrid.Rows[num].Cells[5].Value = "";
				dataGrid.Rows[num].Cells[6].Value = "";
				dataGrid.Rows[num].Cells[7].Value = "TỔNG";
				dataGrid.Rows[num].Cells[8].Value = num2.ToString("#,##0.0");
				dataGrid.Rows[num].Cells[9].Value = num4.ToString();
				dataGrid.Rows[num].Cells[10].Value = "";
				dataGrid.Rows[num].Cells[11].Value = num3.ToString();
				dataGrid.Rows[num].Cells[12].Value = num5.ToString("#,##0.0");
				dataGrid.Rows[num].Cells[13].Value = num6.ToString();
				dataGrid.Rows[num].Cells[14].Value = "";
				dataGrid.Rows[num].Cells[15].Value = "";
				dataGrid.Rows[num].Cells[16].Value = "";
				dataGrid.Rows[num].Cells[17].Value = "";
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = "";
				dataGrid.Rows[num].Cells[1].Value = "";
				dataGrid.Rows[num].Cells[2].Value = "";
				dataGrid.Rows[num].Cells[3].Value = "";
				dataGrid.Rows[num].Cells[4].Value = "";
				dataGrid.Rows[num].Cells[5].Value = "";
				dataGrid.Rows[num].Cells[6].Value = "";
				dataGrid.Rows[num].Cells[7].Value = "TỔNG TL GÀ";
				dataGrid.Rows[num].Cells[8].Value = (num2 - num5).ToString("#,##0.0");
				dataGrid.Rows[num].Cells[9].Value = "TL TB 1 CON";
				dataGrid.Rows[num].Cells[10].Value = Math.Round((num2 - num5 / num6 * num4) / num3, 3).ToString("#,##0.000");
				dataGrid.Rows[num].Cells[11].Value = "";
				dataGrid.Rows[num].Cells[12].Value = "";
				dataGrid.Rows[num].Cells[13].Value = "";
				dataGrid.Rows[num].Cells[14].Value = "";
				dataGrid.Rows[num].Cells[15].Value = "";
				dataGrid.Rows[num].Cells[16].Value = "";
				dataGrid.Rows[num].Cells[17].Value = "";
			}
		}
		catch
		{
		}
		formatDataGrid();
		dataGrid.AutoResizeColumns();
	}

	private void formatDataGrid()
	{
		for (int i = 0; i < ((BaseCollection)dataGrid.Columns).Count; i++)
		{
			((DataGridViewCell)dataGrid.Columns[i].HeaderCell).Style.Alignment = (DataGridViewContentAlignment)32;
		}
	}

	private void btnExcel_Click(object sender, EventArgs e)
	{
		//IL_0275: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_004c: Expected O, but got Unknown
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Expected O, but got Unknown
		//IL_0250: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			((CommonDialog)saveFileDialog1).ShowDialog();
			((FileDialog)saveFileDialog1).DefaultExt = "xls";
			object value = Missing.Value;
			CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
			Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
			Application val = (Application)new ApplicationClass();
			Workbook val2 = ((_Application)val).Workbooks.Add(value);
			Worksheet val3 = (Worksheet)((_Workbook)val2).Worksheets[(object)1];
			for (int i = 1; i < ((BaseCollection)dataGrid.Columns).Count; i++)
			{
				((_Worksheet)val3).Cells[(object)1, (object)i] = dataGrid.Columns[i].HeaderText.ToUpper();
			}
			Range val4 = ((_Worksheet)val3)[(object)"O1", (object)"O30000"];
			val4.NumberFormat = "dd/MM/yyyy HH:mm:ss";
			for (short num = 0; num <= dataGrid.RowCount - 1; num++)
			{
				for (short num2 = 1; num2 <= dataGrid.ColumnCount - 1; num2++)
				{
					((_Worksheet)val3).Cells[(object)(num + 2), (object)num2] = dataGrid[(int)num2, (int)num].Value.ToString();
				}
			}
			val4 = ((_Worksheet)val3)[(object)"A1", (object)("Q" + (dataGrid.RowCount + 1))];
			val4.Borders.Weight = 2;
			val4 = ((_Worksheet)val3)[(object)"A1", (object)"Q1"];
			val4.Font.Bold = true;
			val4 = ((_Worksheet)val3)[(object)"A1", (object)"Q30000"];
			val4.Columns.AutoFit();
			((_Workbook)val2).SaveAs((object)excelName, (object)(XlFileFormat)(-4143), value, value, value, value, (XlSaveAsAccessMode)3, value, value, value, value, value);
			((_Workbook)val2).Close((object)true, value, value);
			((_Application)val).Quit();
			releaseObject(val3);
			releaseObject(val2);
			releaseObject(val);
			MessageBox.Show("Đã xuất file excel xong.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
		}
		catch (Exception ex)
		{
			MessageBox.Show("Đã có lỗi xảy ra. " + ex.Message, "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)16);
		}
	}

	private void releaseObject(object obj)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			Marshal.ReleaseComObject(obj);
			obj = null;
		}
		catch (Exception ex)
		{
			obj = null;
			MessageBox.Show("Không save được!" + ex.ToString());
		}
		finally
		{
			GC.Collect();
		}
	}

	private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
	{
		excelName = ((FileDialog)saveFileDialog1).FileName;
	}

	private void btnXoa_Click(object sender, EventArgs e)
	{
		//IL_0101: Unknown result type (might be due to invalid IL or missing references)
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Invalid comparison between Unknown and I4
		//IL_00d2: Unknown result type (might be due to invalid IL or missing references)
		if (dataGrid.Rows.Count <= 0 || (int)MessageBox.Show("Bạn có muốn xoá dữ liệu cân có STT [" + dataGrid.CurrentRow.Cells["STT2"].Value.ToString() + "] không?", "Thông báo", (MessageBoxButtons)4, (MessageBoxIcon)32) != 6)
		{
			return;
		}
		try
		{
			string text = "";
			text = "DELETE * FROM phieucan WHERE stt =" + dataGrid.CurrentRow.Cells["id"].Value.ToString();
			dataFetcher.setCommand(text, 1);
			int num = dataFetcher.execute();
			if (num != 1)
			{
				MessageBox.Show("Đã xóa.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
			}
			Thread.Sleep(1000);
			btnXem.PerformClick();
		}
		catch
		{
			MessageBox.Show("Không xóa được.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)64);
		}
	}

	private void frmThongke_Resize(object sender, EventArgs e)
	{
		((Control)panel1).Left = (((Control)this).Width - ((Control)panel1).Width) / 2;
		((Control)dataGrid).Width = ((Control)this).Width - 10;
		((Control)dataGrid).Height = ((Control)this).Height - 195;
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
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Expected O, but got Unknown
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Expected O, but got Unknown
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_003c: Expected O, but got Unknown
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected O, but got Unknown
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Expected O, but got Unknown
		//IL_005c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0066: Expected O, but got Unknown
		//IL_0067: Unknown result type (might be due to invalid IL or missing references)
		//IL_0071: Expected O, but got Unknown
		//IL_0072: Unknown result type (might be due to invalid IL or missing references)
		//IL_007c: Expected O, but got Unknown
		//IL_007d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Expected O, but got Unknown
		//IL_0088: Unknown result type (might be due to invalid IL or missing references)
		//IL_0092: Expected O, but got Unknown
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Expected O, but got Unknown
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a8: Expected O, but got Unknown
		//IL_00a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b3: Expected O, but got Unknown
		//IL_00b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00be: Expected O, but got Unknown
		//IL_00bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c9: Expected O, but got Unknown
		//IL_00ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d4: Expected O, but got Unknown
		//IL_00d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00df: Expected O, but got Unknown
		//IL_00e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ea: Expected O, but got Unknown
		//IL_00eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f5: Expected O, but got Unknown
		//IL_00f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0100: Expected O, but got Unknown
		//IL_0101: Unknown result type (might be due to invalid IL or missing references)
		//IL_010b: Expected O, but got Unknown
		//IL_010c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0116: Expected O, but got Unknown
		//IL_0117: Unknown result type (might be due to invalid IL or missing references)
		//IL_0121: Expected O, but got Unknown
		//IL_0122: Unknown result type (might be due to invalid IL or missing references)
		//IL_012c: Expected O, but got Unknown
		//IL_012d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0137: Expected O, but got Unknown
		//IL_0138: Unknown result type (might be due to invalid IL or missing references)
		//IL_0142: Expected O, but got Unknown
		//IL_0143: Unknown result type (might be due to invalid IL or missing references)
		//IL_014d: Expected O, but got Unknown
		//IL_014e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0158: Expected O, but got Unknown
		//IL_0159: Unknown result type (might be due to invalid IL or missing references)
		//IL_0163: Expected O, but got Unknown
		//IL_0164: Unknown result type (might be due to invalid IL or missing references)
		//IL_016e: Expected O, but got Unknown
		//IL_016f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0179: Expected O, but got Unknown
		//IL_017a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0184: Expected O, but got Unknown
		//IL_0185: Unknown result type (might be due to invalid IL or missing references)
		//IL_018f: Expected O, but got Unknown
		//IL_0190: Unknown result type (might be due to invalid IL or missing references)
		//IL_019a: Expected O, but got Unknown
		//IL_019b: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a5: Expected O, but got Unknown
		//IL_01a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b0: Expected O, but got Unknown
		//IL_01b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01bb: Expected O, but got Unknown
		//IL_01bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c6: Expected O, but got Unknown
		//IL_01c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d1: Expected O, but got Unknown
		//IL_01d2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01dc: Expected O, but got Unknown
		//IL_01dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e7: Expected O, but got Unknown
		//IL_01e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f2: Expected O, but got Unknown
		//IL_01f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fd: Expected O, but got Unknown
		//IL_01fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0208: Expected O, but got Unknown
		//IL_0209: Unknown result type (might be due to invalid IL or missing references)
		//IL_0213: Expected O, but got Unknown
		//IL_0214: Unknown result type (might be due to invalid IL or missing references)
		//IL_021e: Expected O, but got Unknown
		//IL_021f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0229: Expected O, but got Unknown
		//IL_022a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0234: Expected O, but got Unknown
		//IL_0235: Unknown result type (might be due to invalid IL or missing references)
		//IL_023f: Expected O, but got Unknown
		//IL_0240: Unknown result type (might be due to invalid IL or missing references)
		//IL_024a: Expected O, but got Unknown
		//IL_024b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0255: Expected O, but got Unknown
		//IL_0294: Unknown result type (might be due to invalid IL or missing references)
		//IL_029e: Expected O, but got Unknown
		//IL_0348: Unknown result type (might be due to invalid IL or missing references)
		//IL_0352: Expected O, but got Unknown
		//IL_0429: Unknown result type (might be due to invalid IL or missing references)
		//IL_0433: Expected O, but got Unknown
		//IL_050a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0514: Expected O, but got Unknown
		//IL_06d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_06df: Expected O, but got Unknown
		//IL_0787: Unknown result type (might be due to invalid IL or missing references)
		//IL_0791: Expected O, but got Unknown
		//IL_081f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0829: Expected O, but got Unknown
		//IL_08a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_08b3: Expected O, but got Unknown
		//IL_0941: Unknown result type (might be due to invalid IL or missing references)
		//IL_094b: Expected O, but got Unknown
		//IL_09be: Unknown result type (might be due to invalid IL or missing references)
		//IL_09c8: Expected O, but got Unknown
		//IL_0a42: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a4c: Expected O, but got Unknown
		//IL_0aca: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ad4: Expected O, but got Unknown
		//IL_0b45: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b4f: Expected O, but got Unknown
		//IL_0bc9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bd3: Expected O, but got Unknown
		//IL_0c61: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c6b: Expected O, but got Unknown
		//IL_0ceb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cf5: Expected O, but got Unknown
		//IL_0d9e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0da8: Expected O, but got Unknown
		//IL_0e31: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e3b: Expected O, but got Unknown
		//IL_0ee6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ef0: Expected O, but got Unknown
		//IL_0f79: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f83: Expected O, but got Unknown
		//IL_1022: Unknown result type (might be due to invalid IL or missing references)
		//IL_102c: Expected O, but got Unknown
		//IL_10be: Unknown result type (might be due to invalid IL or missing references)
		//IL_10c8: Expected O, but got Unknown
		//IL_143f: Unknown result type (might be due to invalid IL or missing references)
		//IL_1449: Expected O, but got Unknown
		//IL_14f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_14fd: Expected O, but got Unknown
		//IL_159c: Unknown result type (might be due to invalid IL or missing references)
		//IL_15a6: Expected O, but got Unknown
		//IL_1623: Unknown result type (might be due to invalid IL or missing references)
		//IL_162d: Expected O, but got Unknown
		//IL_1ba5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1baf: Expected O, but got Unknown
		//IL_1bc5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1bcf: Expected O, but got Unknown
		//IL_1bda: Unknown result type (might be due to invalid IL or missing references)
		DataGridViewCellStyle val = new DataGridViewCellStyle();
		DataGridViewCellStyle val2 = new DataGridViewCellStyle();
		DataGridViewCellStyle val3 = new DataGridViewCellStyle();
		DataGridViewCellStyle val4 = new DataGridViewCellStyle();
		DataGridViewCellStyle val5 = new DataGridViewCellStyle();
		DataGridViewCellStyle val6 = new DataGridViewCellStyle();
		DataGridViewCellStyle val7 = new DataGridViewCellStyle();
		DataGridViewCellStyle val8 = new DataGridViewCellStyle();
		DataGridViewCellStyle val9 = new DataGridViewCellStyle();
		DataGridViewCellStyle val10 = new DataGridViewCellStyle();
		DataGridViewCellStyle val11 = new DataGridViewCellStyle();
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmThongke));
		label1 = new Label();
		saveFileDialog1 = new SaveFileDialog();
		btnExcel = new Button();
		btnThoat = new Button();
		btnXem = new Button();
		dataGrid = new DataGridView();
		chkNhanvien = new CheckBox();
		cboNhanvien = new ComboBox();
		chkSanpham = new CheckBox();
		cboSanpham = new ComboBox();
		txtEndDate = new DateTimePicker();
		label10 = new Label();
		label9 = new Label();
		txtStartDate = new DateTimePicker();
		chkKhachhang = new CheckBox();
		cboKhachhang = new ComboBox();
		chkXuatnhap = new CheckBox();
		cboXuatnhap = new ComboBox();
		chkGhichu = new CheckBox();
		cboGhichu = new ComboBox();
		chkSophieu = new CheckBox();
		cboSophieu = new ComboBox();
		btnXemTong = new Button();
		panel1 = new Panel();
		cboHinhthuccan = new ComboBox();
		chkHinhthuccan = new CheckBox();
		cboSoxe = new ComboBox();
		chkSoxe = new CheckBox();
		id = new DataGridViewTextBoxColumn();
		STT2 = new DataGridViewTextBoxColumn();
		sophieu = new DataGridViewTextBoxColumn();
		soxe = new DataGridViewTextBoxColumn();
		ma = new DataGridViewTextBoxColumn();
		ten = new DataGridViewTextBoxColumn();
		makh = new DataGridViewTextBoxColumn();
		tenkh = new DataGridViewTextBoxColumn();
		tlcan = new DataGridViewTextBoxColumn();
		solong = new DataGridViewTextBoxColumn();
		socon1long = new DataGridViewTextBoxColumn();
		tongsocon = new DataGridViewTextBoxColumn();
		tlbi = new DataGridViewTextBoxColumn();
		solongbi = new DataGridViewTextBoxColumn();
		ghichu = new DataGridViewTextBoxColumn();
		ngaycan = new DataGridViewTextBoxColumn();
		nhanvien = new DataGridViewTextBoxColumn();
		loaican = new DataGridViewTextBoxColumn();
		((ISupportInitialize)dataGrid).BeginInit();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 20.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(343, 5);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(356, 31);
		((Control)label1).TabIndex = 117;
		((Control)label1).Text = "THỐNG KÊ DỮ LIỆU CÂN";
		((FileDialog)saveFileDialog1).FileOk += saveFileDialog1_FileOk;
		((Control)btnExcel).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnExcel).FlatAppearance.BorderSize = 0;
		((Control)btnExcel).Font = new Font("Microsoft Sans Serif", 11.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnExcel).Image = (Image)(object)Resources.Excel_icon;
		((ButtonBase)btnExcel).ImageAlign = (ContentAlignment)16;
		((Control)btnExcel).Location = new Point(594, 108);
		((Control)btnExcel).Name = "btnExcel";
		((Control)btnExcel).Size = new Size(174, 41);
		((Control)btnExcel).TabIndex = 14;
		((Control)btnExcel).Text = "        Xuất Excel";
		((ButtonBase)btnExcel).UseVisualStyleBackColor = true;
		((Control)btnExcel).Click += btnExcel_Click;
		((Control)btnThoat).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnThoat).FlatAppearance.BorderSize = 0;
		((Control)btnThoat).Font = new Font("Microsoft Sans Serif", 11.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnThoat).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnThoat).ImageAlign = (ContentAlignment)16;
		((Control)btnThoat).Location = new Point(774, 108);
		((Control)btnThoat).Name = "btnThoat";
		((Control)btnThoat).Size = new Size(174, 41);
		((Control)btnThoat).TabIndex = 15;
		((Control)btnThoat).Text = "        Thoát";
		((ButtonBase)btnThoat).UseVisualStyleBackColor = true;
		((Control)btnThoat).Click += btnThoat_Click;
		((Control)btnXem).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnXem).FlatAppearance.BorderSize = 0;
		((Control)btnXem).Font = new Font("Microsoft Sans Serif", 11.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnXem).Image = (Image)(object)Resources.report_icon;
		((ButtonBase)btnXem).ImageAlign = (ContentAlignment)16;
		((Control)btnXem).Location = new Point(401, 108);
		((Control)btnXem).Name = "btnXem";
		((Control)btnXem).Size = new Size(187, 41);
		((Control)btnXem).TabIndex = 13;
		((Control)btnXem).Text = "        Thống kê chi tiết";
		((ButtonBase)btnXem).UseVisualStyleBackColor = true;
		((Control)btnXem).Click += btnXem_Click;
		dataGrid.AllowUserToAddRows = false;
		dataGrid.AllowUserToOrderColumns = true;
		dataGrid.ColumnHeadersHeightSizeMode = (DataGridViewColumnHeadersHeightSizeMode)2;
		dataGrid.Columns.AddRange((DataGridViewColumn[])(object)new DataGridViewColumn[18]
		{
			(DataGridViewColumn)id,
			(DataGridViewColumn)STT2,
			(DataGridViewColumn)sophieu,
			(DataGridViewColumn)soxe,
			(DataGridViewColumn)ma,
			(DataGridViewColumn)ten,
			(DataGridViewColumn)makh,
			(DataGridViewColumn)tenkh,
			(DataGridViewColumn)tlcan,
			(DataGridViewColumn)solong,
			(DataGridViewColumn)socon1long,
			(DataGridViewColumn)tongsocon,
			(DataGridViewColumn)tlbi,
			(DataGridViewColumn)solongbi,
			(DataGridViewColumn)ghichu,
			(DataGridViewColumn)ngaycan,
			(DataGridViewColumn)nhanvien,
			(DataGridViewColumn)loaican
		});
		((Control)dataGrid).DataBindings.Add(new Binding("Font", (object)Settings.Default, "font12", true, (DataSourceUpdateMode)1));
		((Control)dataGrid).Font = Settings.Default.font12;
		((Control)dataGrid).Location = new Point(2, 166);
		((Control)dataGrid).Name = "dataGrid";
		dataGrid.RowTemplate.Height = 25;
		dataGrid.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)dataGrid).Size = new Size(1134, 494);
		((Control)dataGrid).TabIndex = 119;
		((Control)chkNhanvien).AutoSize = true;
		((Control)chkNhanvien).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkNhanvien).Location = new Point(560, 47);
		((Control)chkNhanvien).Name = "chkNhanvien";
		((Control)chkNhanvien).Size = new Size(98, 24);
		((Control)chkNhanvien).TabIndex = 158;
		((Control)chkNhanvien).Text = "Nhân viên";
		((ButtonBase)chkNhanvien).UseVisualStyleBackColor = true;
		cboNhanvien.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboNhanvien).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboNhanvien).FormattingEnabled = true;
		((Control)cboNhanvien).Location = new Point(660, 43);
		((Control)cboNhanvien).Name = "cboNhanvien";
		((Control)cboNhanvien).Size = new Size(210, 28);
		((Control)cboNhanvien).TabIndex = 159;
		((Control)chkSanpham).AutoSize = true;
		((Control)chkSanpham).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkSanpham).Location = new Point(221, 49);
		((Control)chkSanpham).Name = "chkSanpham";
		((Control)chkSanpham).Size = new Size(98, 24);
		((Control)chkSanpham).TabIndex = 156;
		((Control)chkSanpham).Text = "Loại hàng";
		((ButtonBase)chkSanpham).UseVisualStyleBackColor = true;
		cboSanpham.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboSanpham).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboSanpham).FormattingEnabled = true;
		((Control)cboSanpham).Location = new Point(337, 45);
		((Control)cboSanpham).Name = "cboSanpham";
		((Control)cboSanpham).Size = new Size(215, 28);
		((Control)cboSanpham).TabIndex = 157;
		((Control)txtEndDate).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		txtEndDate.Format = (DateTimePickerFormat)8;
		((Control)txtEndDate).Location = new Point(96, 77);
		((Control)txtEndDate).Name = "txtEndDate";
		((Control)txtEndDate).Size = new Size(113, 26);
		((Control)txtEndDate).TabIndex = 155;
		((Control)label10).AutoSize = true;
		((Control)label10).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label10).Location = new Point(12, 53);
		((Control)label10).Name = "label10";
		((Control)label10).Size = new Size(65, 20);
		((Control)label10).TabIndex = 160;
		((Control)label10).Text = "Từ ngày";
		((Control)label9).AutoSize = true;
		((Control)label9).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label9).Location = new Point(12, 83);
		((Control)label9).Name = "label9";
		((Control)label9).Size = new Size(77, 20);
		((Control)label9).TabIndex = 161;
		((Control)label9).Text = "Đến ngày";
		((Control)txtStartDate).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		txtStartDate.Format = (DateTimePickerFormat)8;
		((Control)txtStartDate).Location = new Point(96, 49);
		((Control)txtStartDate).Name = "txtStartDate";
		((Control)txtStartDate).Size = new Size(113, 26);
		((Control)txtStartDate).TabIndex = 154;
		((Control)chkKhachhang).AutoSize = true;
		((Control)chkKhachhang).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkKhachhang).Location = new Point(221, 77);
		((Control)chkKhachhang).Name = "chkKhachhang";
		((Control)chkKhachhang).Size = new Size(113, 24);
		((Control)chkKhachhang).TabIndex = 162;
		((Control)chkKhachhang).Text = "Khách hàng";
		((ButtonBase)chkKhachhang).UseVisualStyleBackColor = true;
		cboKhachhang.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboKhachhang).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboKhachhang).FormattingEnabled = true;
		((Control)cboKhachhang).Location = new Point(337, 75);
		((Control)cboKhachhang).Name = "cboKhachhang";
		((Control)cboKhachhang).Size = new Size(215, 28);
		((Control)cboKhachhang).TabIndex = 163;
		((Control)chkXuatnhap).AutoSize = true;
		((Control)chkXuatnhap).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkXuatnhap).Location = new Point(42, 11);
		((Control)chkXuatnhap).Name = "chkXuatnhap";
		((Control)chkXuatnhap).Size = new Size(102, 24);
		((Control)chkXuatnhap).TabIndex = 164;
		((Control)chkXuatnhap).Text = "Xuất nhập";
		((ButtonBase)chkXuatnhap).UseVisualStyleBackColor = true;
		((Control)chkXuatnhap).Visible = false;
		cboXuatnhap.AutoCompleteMode = (AutoCompleteMode)3;
		cboXuatnhap.AutoCompleteSource = (AutoCompleteSource)256;
		((Control)cboXuatnhap).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboXuatnhap).FormattingEnabled = true;
		((Control)cboXuatnhap).Location = new Point(150, 7);
		((Control)cboXuatnhap).Name = "cboXuatnhap";
		((Control)cboXuatnhap).Size = new Size(63, 28);
		((Control)cboXuatnhap).TabIndex = 165;
		((Control)cboXuatnhap).Visible = false;
		((Control)chkGhichu).AutoSize = true;
		((Control)chkGhichu).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkGhichu).Location = new Point(753, 4);
		((Control)chkGhichu).Name = "chkGhichu";
		((Control)chkGhichu).Size = new Size(83, 24);
		((Control)chkGhichu).TabIndex = 166;
		((Control)chkGhichu).Text = "Ghi chú";
		((ButtonBase)chkGhichu).UseVisualStyleBackColor = true;
		((Control)chkGhichu).Visible = false;
		cboGhichu.AutoCompleteMode = (AutoCompleteMode)3;
		cboGhichu.AutoCompleteSource = (AutoCompleteSource)256;
		((Control)cboGhichu).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboGhichu).FormattingEnabled = true;
		((Control)cboGhichu).Location = new Point(863, 4);
		((Control)cboGhichu).Name = "cboGhichu";
		((Control)cboGhichu).Size = new Size(126, 28);
		((Control)cboGhichu).TabIndex = 167;
		((Control)cboGhichu).Visible = false;
		((Control)chkSophieu).AutoSize = true;
		((Control)chkSophieu).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkSophieu).Location = new Point(560, 78);
		((Control)chkSophieu).Name = "chkSophieu";
		((Control)chkSophieu).Size = new Size(91, 24);
		((Control)chkSophieu).TabIndex = 168;
		((Control)chkSophieu).Text = "Số phiếu";
		((ButtonBase)chkSophieu).UseVisualStyleBackColor = true;
		cboSophieu.AutoCompleteMode = (AutoCompleteMode)3;
		cboSophieu.AutoCompleteSource = (AutoCompleteSource)256;
		((Control)cboSophieu).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboSophieu).FormattingEnabled = true;
		((Control)cboSophieu).Location = new Point(660, 74);
		((Control)cboSophieu).Name = "cboSophieu";
		((Control)cboSophieu).Size = new Size(210, 28);
		((Control)cboSophieu).TabIndex = 169;
		((Control)btnXemTong).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnXemTong).FlatAppearance.BorderSize = 0;
		((Control)btnXemTong).Font = new Font("Microsoft Sans Serif", 11.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnXemTong).Image = (Image)(object)Resources.report_icon;
		((ButtonBase)btnXemTong).ImageAlign = (ContentAlignment)16;
		((Control)btnXemTong).Location = new Point(221, 108);
		((Control)btnXemTong).Name = "btnXemTong";
		((Control)btnXemTong).Size = new Size(174, 41);
		((Control)btnXemTong).TabIndex = 170;
		((Control)btnXemTong).Text = "        Thống kê tổng";
		((ButtonBase)btnXemTong).UseVisualStyleBackColor = true;
		((Control)btnXemTong).Click += btnXemTong_Click;
		((Control)panel1).Controls.Add((Control)(object)cboHinhthuccan);
		((Control)panel1).Controls.Add((Control)(object)chkHinhthuccan);
		((Control)panel1).Controls.Add((Control)(object)cboSoxe);
		((Control)panel1).Controls.Add((Control)(object)chkSoxe);
		((Control)panel1).Controls.Add((Control)(object)btnXemTong);
		((Control)panel1).Controls.Add((Control)(object)cboGhichu);
		((Control)panel1).Controls.Add((Control)(object)chkSophieu);
		((Control)panel1).Controls.Add((Control)(object)cboSophieu);
		((Control)panel1).Controls.Add((Control)(object)chkGhichu);
		((Control)panel1).Controls.Add((Control)(object)chkXuatnhap);
		((Control)panel1).Controls.Add((Control)(object)cboXuatnhap);
		((Control)panel1).Controls.Add((Control)(object)chkKhachhang);
		((Control)panel1).Controls.Add((Control)(object)cboKhachhang);
		((Control)panel1).Controls.Add((Control)(object)chkNhanvien);
		((Control)panel1).Controls.Add((Control)(object)cboNhanvien);
		((Control)panel1).Controls.Add((Control)(object)chkSanpham);
		((Control)panel1).Controls.Add((Control)(object)cboSanpham);
		((Control)panel1).Controls.Add((Control)(object)txtEndDate);
		((Control)panel1).Controls.Add((Control)(object)label10);
		((Control)panel1).Controls.Add((Control)(object)label9);
		((Control)panel1).Controls.Add((Control)(object)txtStartDate);
		((Control)panel1).Controls.Add((Control)(object)btnExcel);
		((Control)panel1).Controls.Add((Control)(object)label1);
		((Control)panel1).Controls.Add((Control)(object)btnThoat);
		((Control)panel1).Controls.Add((Control)(object)btnXem);
		((Control)panel1).Location = new Point(12, 9);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(1124, 152);
		((Control)panel1).TabIndex = 171;
		cboHinhthuccan.AutoCompleteMode = (AutoCompleteMode)3;
		cboHinhthuccan.AutoCompleteSource = (AutoCompleteSource)256;
		cboHinhthuccan.DropDownStyle = (ComboBoxStyle)2;
		((Control)cboHinhthuccan).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboHinhthuccan).FormattingEnabled = true;
		cboHinhthuccan.Items.AddRange(new object[2] { "Cân gà", "Cân bì" });
		((Control)cboHinhthuccan).Location = new Point(1002, 76);
		((Control)cboHinhthuccan).Name = "cboHinhthuccan";
		((Control)cboHinhthuccan).Size = new Size(119, 28);
		((Control)cboHinhthuccan).TabIndex = 174;
		((Control)chkHinhthuccan).AutoSize = true;
		((Control)chkHinhthuccan).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkHinhthuccan).Location = new Point(882, 78);
		((Control)chkHinhthuccan).Name = "chkHinhthuccan";
		((Control)chkHinhthuccan).Size = new Size(126, 24);
		((Control)chkHinhthuccan).TabIndex = 173;
		((Control)chkHinhthuccan).Text = "Hình thức cân";
		((ButtonBase)chkHinhthuccan).UseVisualStyleBackColor = true;
		cboSoxe.AutoCompleteMode = (AutoCompleteMode)3;
		cboSoxe.AutoCompleteSource = (AutoCompleteSource)256;
		((Control)cboSoxe).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((ListControl)cboSoxe).FormattingEnabled = true;
		((Control)cboSoxe).Location = new Point(1002, 45);
		((Control)cboSoxe).Name = "cboSoxe";
		((Control)cboSoxe).Size = new Size(119, 28);
		((Control)cboSoxe).TabIndex = 172;
		((Control)chkSoxe).AutoSize = true;
		((Control)chkSoxe).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkSoxe).Location = new Point(882, 47);
		((Control)chkSoxe).Name = "chkSoxe";
		((Control)chkSoxe).Size = new Size(68, 24);
		((Control)chkSoxe).TabIndex = 171;
		((Control)chkSoxe).Text = "Số xe";
		((ButtonBase)chkSoxe).UseVisualStyleBackColor = true;
		((DataGridViewColumn)id).HeaderText = "ID";
		((DataGridViewColumn)id).Name = "id";
		((DataGridViewBand)id).Visible = false;
		val.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)STT2).DefaultCellStyle = val;
		((DataGridViewColumn)STT2).HeaderText = "STT";
		((DataGridViewColumn)STT2).Name = "STT2";
		((DataGridViewColumn)STT2).Width = 60;
		val2.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)sophieu).DefaultCellStyle = val2;
		((DataGridViewColumn)sophieu).HeaderText = "Số phiếu";
		((DataGridViewColumn)sophieu).Name = "sophieu";
		val3.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)soxe).DefaultCellStyle = val3;
		((DataGridViewColumn)soxe).HeaderText = "Số xe";
		((DataGridViewColumn)soxe).Name = "soxe";
		((DataGridViewColumn)soxe).Width = 120;
		((DataGridViewColumn)ma).HeaderText = "Mã hàng";
		((DataGridViewColumn)ma).Name = "ma";
		val4.Format = "##0.00";
		((DataGridViewBand)ten).DefaultCellStyle = val4;
		((DataGridViewColumn)ten).HeaderText = "Tên hàng";
		((DataGridViewColumn)ten).Name = "ten";
		((DataGridViewColumn)ten).Width = 150;
		((DataGridViewColumn)makh).HeaderText = "Mã KH";
		((DataGridViewColumn)makh).Name = "makh";
		((DataGridViewColumn)makh).Width = 180;
		((DataGridViewColumn)tenkh).HeaderText = "Tên KH";
		((DataGridViewColumn)tenkh).Name = "tenkh";
		((DataGridViewColumn)tenkh).Width = 150;
		val5.Alignment = (DataGridViewContentAlignment)64;
		val5.Format = "#,##0.0";
		((DataGridViewBand)tlcan).DefaultCellStyle = val5;
		((DataGridViewColumn)tlcan).HeaderText = "TL cân gà";
		((DataGridViewColumn)tlcan).Name = "tlcan";
		((DataGridViewColumn)tlcan).Width = 120;
		val6.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)solong).DefaultCellStyle = val6;
		((DataGridViewColumn)solong).HeaderText = "Số lồng gà";
		((DataGridViewColumn)solong).Name = "solong";
		((DataGridViewColumn)solong).Width = 120;
		val7.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)socon1long).DefaultCellStyle = val7;
		((DataGridViewColumn)socon1long).HeaderText = "Số con 1 lồng";
		((DataGridViewColumn)socon1long).Name = "socon1long";
		((DataGridViewColumn)socon1long).Width = 180;
		val8.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)tongsocon).DefaultCellStyle = val8;
		((DataGridViewColumn)tongsocon).HeaderText = "Tổng số con";
		((DataGridViewColumn)tongsocon).Name = "tongsocon";
		((DataGridViewColumn)tongsocon).Width = 150;
		val9.Alignment = (DataGridViewContentAlignment)64;
		val9.Format = "#,##0.0";
		((DataGridViewBand)tlbi).DefaultCellStyle = val9;
		((DataGridViewColumn)tlbi).HeaderText = "TL cân bì";
		((DataGridViewColumn)tlbi).Name = "tlbi";
		((DataGridViewColumn)tlbi).Width = 140;
		val10.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)solongbi).DefaultCellStyle = val10;
		((DataGridViewColumn)solongbi).HeaderText = "Số lồng bì";
		((DataGridViewColumn)solongbi).Name = "solongbi";
		((DataGridViewColumn)solongbi).Width = 140;
		((DataGridViewColumn)ghichu).HeaderText = "Ghi chú";
		((DataGridViewColumn)ghichu).Name = "ghichu";
		((DataGridViewColumn)ghichu).Width = 150;
		val11.Format = "dd/MM/yyyy HH:mm:ss";
		val11.NullValue = null;
		((DataGridViewBand)ngaycan).DefaultCellStyle = val11;
		((DataGridViewColumn)ngaycan).HeaderText = "Ngày cân";
		((DataGridViewColumn)ngaycan).Name = "ngaycan";
		((DataGridViewColumn)ngaycan).Width = 170;
		((DataGridViewColumn)nhanvien).HeaderText = "Nhân viên cân";
		((DataGridViewColumn)nhanvien).Name = "nhanvien";
		((DataGridViewColumn)nhanvien).Width = 150;
		((DataGridViewColumn)loaican).HeaderText = "Hình thức cân";
		((DataGridViewColumn)loaican).Name = "loaican";
		((DataGridViewColumn)loaican).Width = 150;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(1137, 662);
		((Control)this).Controls.Add((Control)(object)panel1);
		((Control)this).Controls.Add((Control)(object)dataGrid);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).KeyPreview = true;
		((Form)this).Margin = new Padding(4);
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmThongke";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Thống kê chi tiết xuất nhập";
		((Form)this).Load += frmThongke_Load;
		((Control)this).Resize += frmThongke_Resize;
		((ISupportInitialize)dataGrid).EndInit();
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
