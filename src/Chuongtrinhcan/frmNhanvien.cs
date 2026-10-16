using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmNhanvien : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private int flagSave;

	private string gID;

	private IContainer components = null;

	private Label label4;

	private DataGridView dataGrid;

	private TextBox txtTim;

	private TextBox txtMa;

	private Label label1;

	private Button btnThoat;

	private Button btnXoa;

	private Button btnSave;

	private Label lblMessage;

	private TextBox txtName;

	private ErrorProvider nameError;

	private Label label5;

	private Label label6;

	private TextBox txtPassword;

	private Label label2;

	private Label label3;

	private CheckBox chkTrangthai;

	private DataGridViewTextBoxColumn id;

	private DataGridViewTextBoxColumn stt;

	private DataGridViewTextBoxColumn ma;

	private DataGridViewTextBoxColumn ten;

	private DataGridViewTextBoxColumn password;

	private DataGridViewTextBoxColumn trangthai;

	public frmNhanvien()
	{
		InitializeComponent();
	}

	private void frmNhanvien_Load(object sender, EventArgs e)
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
				text = "SELECT nv_id,tennv,matkhau,trangthai from nhanvien order by tennv ";
			}
			else
			{
				text = "SELECT nv_id,tennv,matkhau,trangthai from nhanvien";
				string text2 = text;
				text = text2 + " where tennv like '%" + keyword + "%' or nv_id like '%" + keyword + "%' order by tennv ";
			}
			dataFetcher.setCommand(text, 1);
			DataSet dataSet = dataFetcher.fetchDataSet();
			int num = 0;
			foreach (DataRow row in dataSet.Tables[0].Rows)
			{
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = row["nv_id"];
				dataGrid.Rows[num].Cells[1].Value = (num + 1).ToString();
				dataGrid.Rows[num].Cells[2].Value = row["nv_id"];
				dataGrid.Rows[num].Cells[3].Value = row["tennv"];
				dataGrid.Rows[num].Cells[4].Value = row["matkhau"];
				if (double.Parse(row["trangthai"].ToString()) == 1.0)
				{
					dataGrid.Rows[num].Cells[5].Value = "Khóa";
				}
				else
				{
					dataGrid.Rows[num].Cells[5].Value = "Mở";
				}
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
		//IL_00ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b3: Invalid comparison between Unknown and I4
		//IL_0067: Unknown result type (might be due to invalid IL or missing references)
		if (dataGrid.Rows.Count <= 0)
		{
			return;
		}
		if (dataGrid.CurrentRow.Cells["ma"].Value.ToString().ToLower() == "admin")
		{
			MessageBox.Show("Không xóa user admin được.", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)16);
		}
		else
		{
			if ((int)MessageBox.Show("Bạn có muốn xóa nhân viên có mã [" + dataGrid.CurrentRow.Cells["ma"].Value.ToString() + "] không?", "Thông báo", (MessageBoxButtons)4, (MessageBoxIcon)32) != 6)
			{
				return;
			}
			try
			{
				string text = "";
				text = "DELETE * FROM nhanvien WHERE nv_id ='" + gID + "'";
				dataFetcher.setCommand(text, 1);
				int num = dataFetcher.execute();
				if (num != 1)
				{
					((Control)lblMessage).Text = "Không xóa được do nhân viên này đã có thao tác với chương trình.";
				}
				Thread.Sleep(1000);
				showData(((Control)txtTim).Text);
				((Control)txtMa).Enabled = true;
				((Control)txtMa).Text = "";
				((Control)txtName).Text = "";
				((Control)txtPassword).Text = "";
				chkTrangthai.Checked = false;
				flagSave = 0;
			}
			catch
			{
				((Control)lblMessage).Text = "Không xóa được do nhân viên này đã có thao tác với chương trình.";
			}
		}
	}

	private bool checkForm()
	{
		bool result = true;
		nameError.Clear();
		if (((Control)txtMa).Text == "")
		{
			nameError.SetError((Control)(object)txtMa, "Vui lòng nhập Mã nhân viên!");
			result = false;
		}
		if (((Control)txtName).Text == "")
		{
			nameError.SetError((Control)(object)txtName, "Vui lòng nhập Tên nhân viên!");
			result = false;
		}
		if (((Control)txtPassword).Text == "")
		{
			nameError.SetError((Control)(object)txtPassword, "Vui lòng nhập Mật khẩu!");
			result = false;
		}
		if ((flagSave == 0) & DBUtil.isDuplicated(((Control)txtMa).Text, "nv_id", "nhanvien"))
		{
			nameError.SetError((Control)(object)txtMa, "Mã nhân viên bị trùng!");
			result = false;
		}
		return result;
	}

	private void insertData()
	{
		string query = "INSERT INTO nhanvien(nv_id,tennv,matkhau,trangthai) VALUES(@nv_id,@tennv,@matkhau,@trangthai)";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@nv_id", ((Control)txtMa).Text);
		dataFetcher.addParameterWithValue("@tennv", ((Control)txtName).Text);
		dataFetcher.addParameterWithValue("@matkhau", ((Control)txtPassword).Text);
		if (chkTrangthai.Checked)
		{
			dataFetcher.addParameterWithValue("@trangthai", "1");
		}
		else
		{
			dataFetcher.addParameterWithValue("@trangthai", "0");
		}
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
		((Control)txtPassword).Text = "";
		chkTrangthai.Checked = false;
	}

	private void updateData()
	{
		string query = "UPDATE nhanvien SET tennv=@tennv, matkhau = @matkhau,trangthai=@trangthai WHERE nv_id = @id";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@tennv", ((Control)txtName).Text);
		dataFetcher.addParameterWithValue("@matkhau", ((Control)txtPassword).Text);
		if (chkTrangthai.Checked)
		{
			dataFetcher.addParameterWithValue("@trangthai", "1");
		}
		else
		{
			dataFetcher.addParameterWithValue("@trangthai", "0");
		}
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
		((Control)txtPassword).Text = "";
		chkTrangthai.Checked = false;
		flagSave = 0;
	}

	private void dataGrid_DoubleClick(object sender, EventArgs e)
	{
		try
		{
			gID = dataGrid.CurrentRow.Cells["id"].Value.ToString();
			((Control)txtMa).Text = dataGrid.CurrentRow.Cells["ma"].Value.ToString();
			((Control)txtName).Text = dataGrid.CurrentRow.Cells["ten"].Value.ToString();
			((Control)txtPassword).Text = dataGrid.CurrentRow.Cells["password"].Value.ToString();
			if (dataGrid.CurrentRow.Cells["trangthai"].Value.ToString() == "Khóa")
			{
				chkTrangthai.Checked = true;
			}
			else
			{
				chkTrangthai.Checked = false;
			}
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
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_0033: Expected O, but got Unknown
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003e: Expected O, but got Unknown
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Expected O, but got Unknown
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		//IL_0060: Unknown result type (might be due to invalid IL or missing references)
		//IL_006a: Expected O, but got Unknown
		//IL_006b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Expected O, but got Unknown
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Expected O, but got Unknown
		//IL_0081: Unknown result type (might be due to invalid IL or missing references)
		//IL_008b: Expected O, but got Unknown
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0096: Expected O, but got Unknown
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
		//IL_016a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0174: Expected O, but got Unknown
		//IL_0281: Unknown result type (might be due to invalid IL or missing references)
		//IL_028b: Expected O, but got Unknown
		//IL_0357: Unknown result type (might be due to invalid IL or missing references)
		//IL_0361: Expected O, but got Unknown
		//IL_03df: Unknown result type (might be due to invalid IL or missing references)
		//IL_03e9: Expected O, but got Unknown
		//IL_0458: Unknown result type (might be due to invalid IL or missing references)
		//IL_0462: Expected O, but got Unknown
		//IL_04f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fc: Expected O, but got Unknown
		//IL_05d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e0: Expected O, but got Unknown
		//IL_06b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c3: Expected O, but got Unknown
		//IL_078a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0794: Expected O, but got Unknown
		//IL_0819: Unknown result type (might be due to invalid IL or missing references)
		//IL_0823: Expected O, but got Unknown
		//IL_089f: Unknown result type (might be due to invalid IL or missing references)
		//IL_08a9: Expected O, but got Unknown
		//IL_0927: Unknown result type (might be due to invalid IL or missing references)
		//IL_0931: Expected O, but got Unknown
		//IL_09b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_09c3: Expected O, but got Unknown
		//IL_0a2f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a39: Expected O, but got Unknown
		//IL_0aba: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ac4: Expected O, but got Unknown
		//IL_0b3d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b47: Expected O, but got Unknown
		//IL_0ea4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0eae: Expected O, but got Unknown
		//IL_0ec3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ecd: Expected O, but got Unknown
		//IL_0ed0: Unknown result type (might be due to invalid IL or missing references)
		components = new Container();
		DataGridViewCellStyle val = new DataGridViewCellStyle();
		DataGridViewCellStyle val2 = new DataGridViewCellStyle();
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmNhanvien));
		label4 = new Label();
		dataGrid = new DataGridView();
		txtTim = new TextBox();
		txtMa = new TextBox();
		label1 = new Label();
		btnThoat = new Button();
		btnXoa = new Button();
		btnSave = new Button();
		lblMessage = new Label();
		txtName = new TextBox();
		nameError = new ErrorProvider(components);
		label5 = new Label();
		label6 = new Label();
		txtPassword = new TextBox();
		label2 = new Label();
		chkTrangthai = new CheckBox();
		label3 = new Label();
		id = new DataGridViewTextBoxColumn();
		stt = new DataGridViewTextBoxColumn();
		ma = new DataGridViewTextBoxColumn();
		ten = new DataGridViewTextBoxColumn();
		password = new DataGridViewTextBoxColumn();
		trangthai = new DataGridViewTextBoxColumn();
		((ISupportInitialize)dataGrid).BeginInit();
		((ISupportInitialize)nameError).BeginInit();
		((Control)this).SuspendLayout();
		((Control)label4).AutoSize = true;
		((Control)label4).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label4).Location = new Point(24, 184);
		((Control)label4).Name = "label4";
		((Control)label4).Size = new Size(106, 20);
		((Control)label4).TabIndex = 152;
		((Control)label4).Text = "Tìm nhân viên";
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
			(DataGridViewColumn)password,
			(DataGridViewColumn)trangthai
		});
		((Control)dataGrid).DataBindings.Add(new Binding("Font", (object)Settings.Default, "font12", true, (DataSourceUpdateMode)1));
		dataGrid.EditMode = (DataGridViewEditMode)4;
		((Control)dataGrid).Font = Settings.Default.font12;
		((Control)dataGrid).Location = new Point(3, 216);
		dataGrid.MultiSelect = false;
		((Control)dataGrid).Name = "dataGrid";
		dataGrid.RowTemplate.Height = 25;
		dataGrid.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)dataGrid).Size = new Size(656, 447);
		((Control)dataGrid).TabIndex = 8;
		((Control)dataGrid).DoubleClick += dataGrid_DoubleClick;
		((Control)txtTim).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtTim).Location = new Point(155, 184);
		((Control)txtTim).Name = "txtTim";
		((Control)txtTim).Size = new Size(237, 26);
		((Control)txtTim).TabIndex = 10;
		((Control)txtTim).TextChanged += txtTim_TextChanged;
		((Control)txtMa).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtMa).Location = new Point(155, 75);
		((Control)txtMa).Name = "txtMa";
		((Control)txtMa).Size = new Size(237, 26);
		((Control)txtMa).TabIndex = 1;
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(24, 79);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(103, 20);
		((Control)label1).TabIndex = 149;
		((Control)label1).Text = "Mã nhân viên";
		((Control)btnThoat).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnThoat).FlatAppearance.BorderSize = 0;
		((Control)btnThoat).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnThoat).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnThoat).ImageAlign = (ContentAlignment)16;
		((Control)btnThoat).Location = new Point(441, 138);
		((Control)btnThoat).Name = "btnThoat";
		((Control)btnThoat).Size = new Size(137, 41);
		((Control)btnThoat).TabIndex = 9;
		((Control)btnThoat).Text = "       Thoát";
		((ButtonBase)btnThoat).UseVisualStyleBackColor = true;
		((Control)btnThoat).Click += btnThoat_Click;
		((Control)btnXoa).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnXoa).FlatAppearance.BorderSize = 0;
		((Control)btnXoa).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnXoa).Image = (Image)(object)Resources.delete;
		((ButtonBase)btnXoa).ImageAlign = (ContentAlignment)16;
		((Control)btnXoa).Location = new Point(298, 138);
		((Control)btnXoa).Name = "btnXoa";
		((Control)btnXoa).Size = new Size(137, 41);
		((Control)btnXoa).TabIndex = 8;
		((Control)btnXoa).Text = "       Xóa";
		((ButtonBase)btnXoa).UseVisualStyleBackColor = true;
		((Control)btnXoa).Click += btnXoa_Click;
		((Control)btnSave).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnSave).FlatAppearance.BorderSize = 0;
		((Control)btnSave).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnSave).Image = (Image)(object)Resources.Save1;
		((ButtonBase)btnSave).ImageAlign = (ContentAlignment)16;
		((Control)btnSave).Location = new Point(155, 138);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(137, 41);
		((Control)btnSave).TabIndex = 7;
		((Control)btnSave).Text = "       Lưu";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)lblMessage).ForeColor = Color.Red;
		((Control)lblMessage).Location = new Point(158, 45);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(74, 20);
		((Control)lblMessage).TabIndex = 148;
		((Control)lblMessage).Text = "message";
		((Control)txtName).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtName).Location = new Point(155, 105);
		((Control)txtName).Name = "txtName";
		((Control)txtName).Size = new Size(237, 26);
		((Control)txtName).TabIndex = 3;
		nameError.ContainerControl = (ContainerControl)(object)this;
		((Control)label5).AutoSize = true;
		((Control)label5).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label5).Location = new Point(24, 109);
		((Control)label5).Name = "label5";
		((Control)label5).Size = new Size(108, 20);
		((Control)label5).TabIndex = 146;
		((Control)label5).Text = "Tên nhân viên";
		((Control)label6).AutoSize = true;
		((Control)label6).Font = new Font("Microsoft Sans Serif", 21.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label6).ForeColor = Color.Black;
		((Control)label6).Location = new Point(149, 12);
		((Control)label6).Name = "label6";
		((Control)label6).Size = new Size(369, 33);
		((Control)label6).TabIndex = 147;
		((Control)label6).Text = "DANH SÁCH NHÂN VIÊN";
		((Control)txtPassword).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtPassword).Location = new Point(488, 73);
		((Control)txtPassword).Name = "txtPassword";
		((Control)txtPassword).Size = new Size(119, 26);
		((Control)txtPassword).TabIndex = 2;
		((Control)label2).AutoSize = true;
		((Control)label2).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label2).Location = new Point(405, 79);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(75, 20);
		((Control)label2).TabIndex = 154;
		((Control)label2).Text = "Mật khẩu";
		((Control)chkTrangthai).AutoSize = true;
		((Control)chkTrangthai).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)chkTrangthai).Location = new Point(488, 111);
		((Control)chkTrangthai).Name = "chkTrangthai";
		((Control)chkTrangthai).Size = new Size(15, 14);
		((Control)chkTrangthai).TabIndex = 4;
		((ButtonBase)chkTrangthai).UseVisualStyleBackColor = true;
		((Control)label3).AutoSize = true;
		((Control)label3).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label3).Location = new Point(405, 109);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(70, 20);
		((Control)label3).TabIndex = 156;
		((Control)label3).Text = "Khóa/Ẩn";
		((DataGridViewColumn)id).HeaderText = "ID";
		((DataGridViewColumn)id).Name = "id";
		((DataGridViewBand)id).Visible = false;
		((DataGridViewColumn)id).Width = 115;
		val.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)stt).DefaultCellStyle = val;
		((DataGridViewColumn)stt).FillWeight = 60.9137f;
		((DataGridViewColumn)stt).HeaderText = "STT";
		((DataGridViewColumn)stt).Name = "stt";
		((DataGridViewColumn)stt).Width = 50;
		val2.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)ma).DefaultCellStyle = val2;
		((DataGridViewColumn)ma).FillWeight = 107.8173f;
		((DataGridViewColumn)ma).HeaderText = "Mã";
		((DataGridViewColumn)ma).Name = "ma";
		((DataGridViewColumn)ma).Width = 140;
		((DataGridViewColumn)ten).HeaderText = "Tên";
		((DataGridViewColumn)ten).Name = "ten";
		((DataGridViewColumn)ten).Width = 200;
		((DataGridViewColumn)password).HeaderText = "Mật khẩu";
		((DataGridViewColumn)password).Name = "password";
		((DataGridViewBand)password).Visible = false;
		((DataGridViewColumn)trangthai).HeaderText = "Trạng thái";
		((DataGridViewColumn)trangthai).Name = "trangthai";
		((DataGridViewColumn)trangthai).Width = 130;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(9f, 18f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(663, 666);
		((Control)this).Controls.Add((Control)(object)label3);
		((Control)this).Controls.Add((Control)(object)chkTrangthai);
		((Control)this).Controls.Add((Control)(object)txtPassword);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)label4);
		((Control)this).Controls.Add((Control)(object)dataGrid);
		((Control)this).Controls.Add((Control)(object)txtTim);
		((Control)this).Controls.Add((Control)(object)txtMa);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)btnThoat);
		((Control)this).Controls.Add((Control)(object)btnXoa);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)txtName);
		((Control)this).Controls.Add((Control)(object)label5);
		((Control)this).Controls.Add((Control)(object)label6);
		((Control)this).Font = new Font("Microsoft Sans Serif", 11.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmNhanvien";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Danh sách nhân viên";
		((Form)this).Load += frmNhanvien_Load;
		((ISupportInitialize)dataGrid).EndInit();
		((ISupportInitialize)nameError).EndInit();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
