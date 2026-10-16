using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmKhachhang : Form
{
	private IContainer components = null;

	private ErrorProvider nameError;

	private Label label4;

	private TextBox txtTim;

	private Label label6;

	private DataGridView dataGrid;

	private Label label5;

	private TextBox txtMa;

	private Label label1;

	private Button btnThoat;

	private Button btnXoa;

	private Button btnSave;

	private Label lblMessage;

	private TextBox txtName;

	private Label label2;

	private TextBox txtDiachi;

	private Label label3;

	private TextBox txtDienthoai;

	private DataGridViewTextBoxColumn id;

	private DataGridViewTextBoxColumn stt;

	private DataGridViewTextBoxColumn ma;

	private DataGridViewTextBoxColumn ten;

	private DataGridViewTextBoxColumn diachi;

	private DataGridViewTextBoxColumn dienthoai;

	private DataFetcher dataFetcher = DataFetcher.Instance;

	private int flagSave;

	private string gID;

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
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Expected O, but got Unknown
		//IL_003a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		//IL_0045: Unknown result type (might be due to invalid IL or missing references)
		//IL_004f: Expected O, but got Unknown
		//IL_0050: Unknown result type (might be due to invalid IL or missing references)
		//IL_005a: Expected O, but got Unknown
		//IL_005b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		//IL_0066: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Expected O, but got Unknown
		//IL_0071: Unknown result type (might be due to invalid IL or missing references)
		//IL_007b: Expected O, but got Unknown
		//IL_007c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0086: Expected O, but got Unknown
		//IL_0087: Unknown result type (might be due to invalid IL or missing references)
		//IL_0091: Expected O, but got Unknown
		//IL_0092: Unknown result type (might be due to invalid IL or missing references)
		//IL_009c: Expected O, but got Unknown
		//IL_009d: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a7: Expected O, but got Unknown
		//IL_00a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b2: Expected O, but got Unknown
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bd: Expected O, but got Unknown
		//IL_00be: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c8: Expected O, but got Unknown
		//IL_00c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d3: Expected O, but got Unknown
		//IL_00d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00de: Expected O, but got Unknown
		//IL_00df: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e9: Expected O, but got Unknown
		//IL_00ea: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f4: Expected O, but got Unknown
		//IL_00f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ff: Expected O, but got Unknown
		//IL_0100: Unknown result type (might be due to invalid IL or missing references)
		//IL_010a: Expected O, but got Unknown
		//IL_010b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0115: Expected O, but got Unknown
		//IL_0116: Unknown result type (might be due to invalid IL or missing references)
		//IL_0120: Expected O, but got Unknown
		//IL_0121: Unknown result type (might be due to invalid IL or missing references)
		//IL_012b: Expected O, but got Unknown
		//IL_0177: Unknown result type (might be due to invalid IL or missing references)
		//IL_0181: Expected O, but got Unknown
		//IL_01f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ff: Expected O, but got Unknown
		//IL_0287: Unknown result type (might be due to invalid IL or missing references)
		//IL_0291: Expected O, but got Unknown
		//IL_03b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_03bb: Expected O, but got Unknown
		//IL_060b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0615: Expected O, but got Unknown
		//IL_0686: Unknown result type (might be due to invalid IL or missing references)
		//IL_0690: Expected O, but got Unknown
		//IL_06fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0706: Expected O, but got Unknown
		//IL_0784: Unknown result type (might be due to invalid IL or missing references)
		//IL_078e: Expected O, but got Unknown
		//IL_0813: Unknown result type (might be due to invalid IL or missing references)
		//IL_081d: Expected O, but got Unknown
		//IL_089b: Unknown result type (might be due to invalid IL or missing references)
		//IL_08a5: Expected O, but got Unknown
		//IL_097e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0988: Expected O, but got Unknown
		//IL_0a61: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a6b: Expected O, but got Unknown
		//IL_0b2f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b39: Expected O, but got Unknown
		//IL_0bad: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bb7: Expected O, but got Unknown
		//IL_0c26: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c30: Expected O, but got Unknown
		//IL_0ca4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cae: Expected O, but got Unknown
		//IL_0e76: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e80: Expected O, but got Unknown
		//IL_0e95: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e9f: Expected O, but got Unknown
		//IL_0ea2: Unknown result type (might be due to invalid IL or missing references)
		components = new Container();
		DataGridViewCellStyle val = new DataGridViewCellStyle();
		DataGridViewCellStyle val2 = new DataGridViewCellStyle();
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmKhachhang));
		nameError = new ErrorProvider(components);
		label4 = new Label();
		txtTim = new TextBox();
		label6 = new Label();
		dataGrid = new DataGridView();
		id = new DataGridViewTextBoxColumn();
		stt = new DataGridViewTextBoxColumn();
		ma = new DataGridViewTextBoxColumn();
		ten = new DataGridViewTextBoxColumn();
		diachi = new DataGridViewTextBoxColumn();
		dienthoai = new DataGridViewTextBoxColumn();
		label5 = new Label();
		txtMa = new TextBox();
		label1 = new Label();
		lblMessage = new Label();
		txtName = new TextBox();
		btnThoat = new Button();
		btnXoa = new Button();
		btnSave = new Button();
		label2 = new Label();
		txtDiachi = new TextBox();
		label3 = new Label();
		txtDienthoai = new TextBox();
		((ISupportInitialize)nameError).BeginInit();
		((ISupportInitialize)dataGrid).BeginInit();
		((Control)this).SuspendLayout();
		nameError.ContainerControl = (ContainerControl)(object)this;
		((Control)label4).AutoSize = true;
		((Control)label4).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label4).Location = new Point(42, 182);
		((Control)label4).Name = "label4";
		((Control)label4).Size = new Size(60, 20);
		((Control)label4).TabIndex = 171;
		((Control)label4).Text = "Tìm KH";
		((Control)txtTim).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtTim).Location = new Point(122, 179);
		((Control)txtTim).Name = "txtTim";
		((Control)txtTim).Size = new Size(280, 26);
		((Control)txtTim).TabIndex = 9;
		((Control)txtTim).TextChanged += txtTim_TextChanged;
		((Control)label6).AutoSize = true;
		((Control)label6).Font = new Font("Microsoft Sans Serif", 21.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label6).ForeColor = Color.Black;
		((Control)label6).Location = new Point(167, 6);
		((Control)label6).Name = "label6";
		((Control)label6).Size = new Size(406, 33);
		((Control)label6).TabIndex = 168;
		((Control)label6).Text = "DANH SÁCH KHÁCH HÀNG";
		dataGrid.AllowUserToAddRows = false;
		dataGrid.AllowUserToOrderColumns = true;
		dataGrid.AllowUserToResizeRows = false;
		dataGrid.BorderStyle = (BorderStyle)2;
		dataGrid.ColumnHeadersHeightSizeMode = (DataGridViewColumnHeadersHeightSizeMode)2;
		dataGrid.Columns.AddRange((DataGridViewColumn[])(object)new DataGridViewColumn[6]
		{
			(DataGridViewColumn)id,
			(DataGridViewColumn)stt,
			(DataGridViewColumn)ma,
			(DataGridViewColumn)ten,
			(DataGridViewColumn)diachi,
			(DataGridViewColumn)dienthoai
		});
		((Control)dataGrid).DataBindings.Add(new Binding("Font", (object)Settings.Default, "font12", true, (DataSourceUpdateMode)1));
		dataGrid.EditMode = (DataGridViewEditMode)4;
		((Control)dataGrid).Font = Settings.Default.font12;
		((Control)dataGrid).Location = new Point(1, 209);
		dataGrid.MultiSelect = false;
		((Control)dataGrid).Name = "dataGrid";
		dataGrid.RowTemplate.Height = 25;
		dataGrid.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)dataGrid).Size = new Size(762, 402);
		((Control)dataGrid).TabIndex = 166;
		((Control)dataGrid).DoubleClick += dataGrid_DoubleClick;
		((DataGridViewColumn)id).HeaderText = "ID";
		((DataGridViewColumn)id).Name = "id";
		((DataGridViewBand)id).Visible = false;
		((DataGridViewColumn)id).Width = 115;
		val.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)stt).DefaultCellStyle = val;
		((DataGridViewColumn)stt).FillWeight = 60.9137f;
		((DataGridViewColumn)stt).HeaderText = "STT";
		((DataGridViewColumn)stt).Name = "stt";
		((DataGridViewColumn)stt).Width = 70;
		val2.Alignment = (DataGridViewContentAlignment)16;
		((DataGridViewBand)ma).DefaultCellStyle = val2;
		((DataGridViewColumn)ma).FillWeight = 107.8173f;
		((DataGridViewColumn)ma).HeaderText = "Mã";
		((DataGridViewColumn)ma).Name = "ma";
		((DataGridViewColumn)ten).HeaderText = "Tên";
		((DataGridViewColumn)ten).Name = "ten";
		((DataGridViewColumn)ten).Width = 200;
		((DataGridViewColumn)diachi).HeaderText = "Địa chỉ";
		((DataGridViewColumn)diachi).Name = "diachi";
		((DataGridViewColumn)diachi).Width = 200;
		((DataGridViewColumn)dienthoai).HeaderText = "Điện thoại";
		((DataGridViewColumn)dienthoai).Name = "dienthoai";
		((DataGridViewColumn)dienthoai).Width = 120;
		((Control)label5).AutoSize = true;
		((Control)label5).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label5).Location = new Point(51, 106);
		((Control)label5).Name = "label5";
		((Control)label5).Size = new Size(36, 20);
		((Control)label5).TabIndex = 167;
		((Control)label5).Text = "Tên";
		((Control)txtMa).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtMa).Location = new Point(122, 73);
		((Control)txtMa).Name = "txtMa";
		((Control)txtMa).Size = new Size(280, 26);
		((Control)txtMa).TabIndex = 1;
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(51, 76);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(31, 20);
		((Control)label1).TabIndex = 170;
		((Control)label1).Text = "Mã";
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblMessage).ForeColor = Color.Red;
		((Control)lblMessage).Location = new Point(170, 42);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(74, 20);
		((Control)lblMessage).TabIndex = 169;
		((Control)lblMessage).Text = "message";
		((Control)txtName).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtName).Location = new Point(122, 103);
		((Control)txtName).Name = "txtName";
		((Control)txtName).Size = new Size(280, 26);
		((Control)txtName).TabIndex = 2;
		((Control)btnThoat).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnThoat).FlatAppearance.BorderSize = 0;
		((Control)btnThoat).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnThoat).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnThoat).ImageAlign = (ContentAlignment)16;
		((Control)btnThoat).Location = new Point(408, 132);
		((Control)btnThoat).Name = "btnThoat";
		((Control)btnThoat).Size = new Size(137, 41);
		((Control)btnThoat).TabIndex = 8;
		((Control)btnThoat).Text = "       Thoát";
		((ButtonBase)btnThoat).UseVisualStyleBackColor = true;
		((Control)btnThoat).Click += btnThoat_Click;
		((Control)btnXoa).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnXoa).FlatAppearance.BorderSize = 0;
		((Control)btnXoa).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnXoa).Image = (Image)(object)Resources.delete;
		((ButtonBase)btnXoa).ImageAlign = (ContentAlignment)16;
		((Control)btnXoa).Location = new Point(265, 132);
		((Control)btnXoa).Name = "btnXoa";
		((Control)btnXoa).Size = new Size(137, 41);
		((Control)btnXoa).TabIndex = 6;
		((Control)btnXoa).Text = "       Xóa";
		((ButtonBase)btnXoa).UseVisualStyleBackColor = true;
		((Control)btnXoa).Click += btnXoa_Click;
		((Control)btnSave).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnSave).FlatAppearance.BorderSize = 0;
		((Control)btnSave).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave).Image = (Image)(object)Resources.Save1;
		((ButtonBase)btnSave).ImageAlign = (ContentAlignment)16;
		((Control)btnSave).Location = new Point(122, 132);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(137, 41);
		((Control)btnSave).TabIndex = 5;
		((Control)btnSave).Text = "       Lưu";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		((Control)label2).AutoSize = true;
		((Control)label2).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label2).Location = new Point(416, 106);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(81, 20);
		((Control)label2).TabIndex = 175;
		((Control)label2).Text = "Điện thoại";
		((Control)txtDiachi).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtDiachi).Location = new Point(498, 73);
		((Control)txtDiachi).Name = "txtDiachi";
		((Control)txtDiachi).Size = new Size(235, 26);
		((Control)txtDiachi).TabIndex = 3;
		((Control)label3).AutoSize = true;
		((Control)label3).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label3).Location = new Point(416, 76);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(57, 20);
		((Control)label3).TabIndex = 176;
		((Control)label3).Text = "Địa chỉ";
		((Control)txtDienthoai).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtDienthoai).Location = new Point(498, 103);
		((Control)txtDienthoai).Name = "txtDienthoai";
		((Control)txtDienthoai).Size = new Size(235, 26);
		((Control)txtDienthoai).TabIndex = 4;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(9f, 18f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(764, 616);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)txtDiachi);
		((Control)this).Controls.Add((Control)(object)label3);
		((Control)this).Controls.Add((Control)(object)txtDienthoai);
		((Control)this).Controls.Add((Control)(object)label4);
		((Control)this).Controls.Add((Control)(object)txtTim);
		((Control)this).Controls.Add((Control)(object)label6);
		((Control)this).Controls.Add((Control)(object)dataGrid);
		((Control)this).Controls.Add((Control)(object)label5);
		((Control)this).Controls.Add((Control)(object)txtMa);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)btnThoat);
		((Control)this).Controls.Add((Control)(object)btnXoa);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)txtName);
		((Control)this).Font = new Font("Microsoft Sans Serif", 11.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmKhachhang";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Danh sách Khách hàng";
		((Form)this).Load += frmCongnhan_Load;
		((ISupportInitialize)nameError).EndInit();
		((ISupportInitialize)dataGrid).EndInit();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}

	public frmKhachhang()
	{
		InitializeComponent();
	}

	private void frmCongnhan_Load(object sender, EventArgs e)
	{
		DBUtil.LoadMaunen((Form)(object)this);
		((Control)lblMessage).Text = "";
		showData("");
		flagSave = 0;
	}

	public void showData(string keyword)
	{
		try
		{
			dataGrid.RowCount = 0;
			string text = "";
			if (keyword == "")
			{
				text = "SELECT id,makh,tenkh,diachi,dienthoai from khachhang order by makh ";
			}
			else
			{
				text = "SELECT id,makh,tenkh,diachi,dienthoai from khachhang ";
				string text2 = text;
				text = text2 + " where makh like N'%" + keyword + "%' or tenkh like N'%" + keyword + "%' and  order by makh ";
			}
			dataFetcher.setCommand(text, 1);
			DataSet dataSet = dataFetcher.fetchDataSet();
			int num = 0;
			foreach (DataRow row in dataSet.Tables[0].Rows)
			{
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = row["id"];
				dataGrid.Rows[num].Cells[1].Value = (num + 1).ToString();
				dataGrid.Rows[num].Cells[2].Value = row["makh"];
				dataGrid.Rows[num].Cells[3].Value = row["tenkh"];
				dataGrid.Rows[num].Cells[4].Value = row["diachi"];
				dataGrid.Rows[num].Cells[5].Value = row["dienthoai"];
			}
		}
		catch
		{
		}
		formatDataGrid();
	}

	private void formatDataGrid()
	{
		for (int i = 0; i < ((BaseCollection)dataGrid.Columns).Count; i++)
		{
			((DataGridViewCell)dataGrid.Columns[i].HeaderCell).Style.Alignment = (DataGridViewContentAlignment)32;
		}
	}

	private void btnSave_Click(object sender, EventArgs e)
	{
		if (checkForm())
		{
			if (flagSave == 1)
			{
				updateData();
				Thread.Sleep(1000);
				showData(((Control)txtTim).Text);
			}
			else
			{
				insertData();
				Thread.Sleep(1000);
				showData(((Control)txtTim).Text);
			}
			((Control)txtMa).Enabled = true;
			((Control)txtMa).Focus();
		}
	}

	private void btnThoat_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void btnXoa_Click(object sender, EventArgs e)
	{
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Invalid comparison between Unknown and I4
		if (dataGrid.Rows.Count <= 0 || (int)MessageBox.Show("Bạn có muốn xóa Khách hàng có mã [" + dataGrid.CurrentRow.Cells["ma"].Value.ToString() + "] không?", "Thông báo", (MessageBoxButtons)4, (MessageBoxIcon)32) != 6)
		{
			return;
		}
		try
		{
			string text = "";
			text = "DELETE * FROM khachhang WHERE id =" + dataGrid.CurrentRow.Cells["id"].Value.ToString();
			dataFetcher.setCommand(text, 1);
			int num = dataFetcher.execute();
			if (num != 1)
			{
				((Control)lblMessage).Text = "Không xóa được do Khách hàng này đã có phiếu tính tiền.";
			}
			Thread.Sleep(1000);
			showData(((Control)txtTim).Text);
			((Control)txtMa).Enabled = true;
			((Control)txtMa).Text = "";
			((Control)txtName).Text = "";
			((Control)txtDiachi).Text = "";
			((Control)txtDienthoai).Text = "";
			((Control)lblMessage).Text = "Đã xóa.";
			flagSave = 0;
		}
		catch
		{
			((Control)lblMessage).Text = "Không xóa được do Khách hàng này đã có phiếu tính tiền.";
		}
	}

	private bool checkForm()
	{
		bool result = true;
		nameError.Clear();
		if (((Control)txtMa).Text == "")
		{
			nameError.SetError((Control)(object)txtMa, "Vui lòng nhập Mã!");
			result = false;
		}
		if (((Control)txtName).Text == "")
		{
			nameError.SetError((Control)(object)txtName, "Vui lòng nhập Tên!");
			result = false;
		}
		if ((flagSave == 0) & DBUtil.isDuplicated(((Control)txtMa).Text, "makh", "khachhang"))
		{
			nameError.SetError((Control)(object)txtMa, "Mã bị trùng!");
			result = false;
		}
		return result;
	}

	private void insertData()
	{
		string query = "INSERT INTO khachhang(makh,tenkh,diachi,dienthoai) VALUES(@makh,@tenkh,@diachi,@dienthoai)";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@makh", ((Control)txtMa).Text);
		dataFetcher.addParameterWithValue("@tenkh", ((Control)txtName).Text);
		dataFetcher.addParameterWithValue("@diachi", ((Control)txtDiachi).Text);
		dataFetcher.addParameterWithValue("@dienthoai", ((Control)txtDienthoai).Text);
		int num = dataFetcher.execute();
		if (num == -1)
		{
			((Control)lblMessage).Text = "Không lưu được.";
			return;
		}
		((Control)lblMessage).Text = "Đã lưu.";
		((Control)txtMa).Text = "";
		((Control)txtMa).Focus();
		((Control)txtName).Text = "";
		((Control)txtDiachi).Text = "";
		((Control)txtDienthoai).Text = "";
	}

	private void updateData()
	{
		string query = "UPDATE khachhang SET tenkh=@tenkh,diachi=@diachi,dienthoai=@dienthoai WHERE id = @id";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@tenkh", ((Control)txtName).Text);
		dataFetcher.addParameterWithValue("@diachi", ((Control)txtDiachi).Text);
		dataFetcher.addParameterWithValue("@dienthoai", ((Control)txtDienthoai).Text);
		dataFetcher.addParameterWithValue("@id", gID);
		int num = dataFetcher.execute();
		if (num == -1)
		{
			((Control)lblMessage).Text = "Không lưu được.";
			return;
		}
		((Control)lblMessage).Text = "Đã lưu.";
		((Control)txtMa).Text = "";
		((Control)txtMa).Focus();
		((Control)txtName).Text = "";
		((Control)txtMa).Enabled = true;
		((Control)txtDienthoai).Text = "";
		((Control)txtDiachi).Text = "";
		flagSave = 0;
	}

	private void dataGrid_DoubleClick(object sender, EventArgs e)
	{
		try
		{
			gID = dataGrid.CurrentRow.Cells["id"].Value.ToString();
			((Control)txtMa).Text = dataGrid.CurrentRow.Cells["ma"].Value.ToString();
			((Control)txtName).Text = dataGrid.CurrentRow.Cells["ten"].Value.ToString();
			((Control)txtDiachi).Text = dataGrid.CurrentRow.Cells["diachi"].Value.ToString();
			((Control)txtDienthoai).Text = dataGrid.CurrentRow.Cells["dienthoai"].Value.ToString();
			((Control)txtMa).Enabled = false;
			flagSave = 1;
		}
		catch
		{
			gID = "0";
		}
	}

	private void txtTim_TextChanged(object sender, EventArgs e)
	{
		showData(((Control)txtTim).Text);
	}
}
