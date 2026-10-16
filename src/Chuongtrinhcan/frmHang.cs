using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Chuongtrinhcan.Properties;

namespace Chuongtrinhcan;

public class frmHang : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private int flagSave;

	private string gID;

	private IContainer components = null;

	private Label lblMessage;

	private DataGridView dataGrid;

	private ErrorProvider nameError;

	private Label label6;

	private TextBox txtName;

	private Label label5;

	private Button btnThoat;

	private Button btnXoa;

	private Button btnSave;

	private TextBox txtMa;

	private Label label1;

	private TextBox txtTim;

	private Label label4;

	private DataGridViewTextBoxColumn id;

	private DataGridViewTextBoxColumn stt;

	private DataGridViewTextBoxColumn mahang;

	private DataGridViewTextBoxColumn tenhang;

	public frmHang()
	{
		InitializeComponent();
	}

	private void frmHang_Load(object sender, EventArgs e)
	{
		DBUtil.LoadMaunen((Form)(object)this);
		if (Globals.gMaNV != "admin")
		{
			((Control)btnSave).Enabled = false;
			((Control)btnXoa).Enabled = false;
		}
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
			text = ((!(keyword == "")) ? ("SELECT id,mahang,tenhang FROM hanghoa where mahang like '%" + keyword + "%' or tenhang like '%" + keyword + "%' order by mahang ") : "SELECT id,mahang,tenhang FROM hanghoa order by mahang ");
			dataFetcher.setCommand(text, 1);
			DataSet dataSet = dataFetcher.fetchDataSet();
			int num = 0;
			foreach (DataRow row in dataSet.Tables[0].Rows)
			{
				num = dataGrid.Rows.Add();
				dataGrid.Rows[num].Cells[0].Value = row["id"];
				dataGrid.Rows[num].Cells[1].Value = (num + 1).ToString();
				dataGrid.Rows[num].Cells[2].Value = row["mahang"];
				dataGrid.Rows[num].Cells[3].Value = row["tenhang"];
			}
			formatDataGrid();
		}
		catch
		{
		}
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
				Thread.Sleep(700);
				showData(((Control)txtTim).Text);
				((Control)txtMa).Enabled = true;
			}
			else
			{
				insertData();
				Thread.Sleep(700);
				showData(((Control)txtTim).Text);
				((Control)txtMa).Enabled = true;
			}
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
		if (dataGrid.Rows.Count <= 0 || (int)MessageBox.Show("Bạn có muốn xoá mặt hàng có mã [" + dataGrid.CurrentRow.Cells["mahang"].Value.ToString() + "] không?", "Thông báo", (MessageBoxButtons)4, (MessageBoxIcon)32) != 6)
		{
			return;
		}
		try
		{
			string text = "";
			text = "DELETE * FROM HANGHOA WHERE id =" + dataGrid.CurrentRow.Cells["id"].Value.ToString();
			dataFetcher.setCommand(text, 1);
			int num = dataFetcher.execute();
			if (num != 1)
			{
				((Control)lblMessage).Text = "Đã xóa.";
			}
			Thread.Sleep(1000);
			showData(((Control)txtTim).Text);
			((Control)txtMa).Enabled = true;
			((Control)txtMa).Text = "";
			((Control)txtName).Text = "";
			flagSave = 0;
		}
		catch
		{
			((Control)lblMessage).Text = "Không xóa được do mặt hàng này đã có lượt cân.";
		}
	}

	private bool checkForm()
	{
		bool result = true;
		nameError.Clear();
		if (((Control)txtMa).Text == "")
		{
			nameError.SetError((Control)(object)txtMa, "Vui lòng nhập Mã hàng!");
			result = false;
		}
		if (((Control)txtName).Text == "")
		{
			nameError.SetError((Control)(object)txtName, "Vui lòng nhập Tên hàng!");
			result = false;
		}
		if ((flagSave == 0) & DBUtil.isDuplicated(((Control)txtMa).Text, "mahang", "hanghoa"))
		{
			nameError.SetError((Control)(object)txtMa, "Mã bị trùng!");
			result = false;
		}
		return result;
	}

	private void insertData()
	{
		string query = "INSERT INTO HANGHOA(MAHANG,TENHANG) VALUES(@ma,@name)";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@ma", ((Control)txtMa).Text);
		dataFetcher.addParameterWithValue("@name", ((Control)txtName).Text);
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
	}

	private void updateData()
	{
		string query = "UPDATE HANGHOA SET mahang=@ma,TENHANG = @name WHERE id = @id";
		dataFetcher.setCommand(query, 1);
		dataFetcher.addParameterWithValue("@ma", ((Control)txtMa).Text);
		dataFetcher.addParameterWithValue("@name", ((Control)txtName).Text);
		dataFetcher.addParameterWithValue("@id", gID.ToString());
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
		flagSave = 0;
	}

	private void dataGrid_DoubleClick(object sender, EventArgs e)
	{
		try
		{
			flagSave = 1;
			gID = dataGrid.CurrentRow.Cells["id"].Value.ToString();
			((Control)txtMa).Text = dataGrid.CurrentRow.Cells["mahang"].Value.ToString();
			((Control)txtName).Text = dataGrid.CurrentRow.Cells["tenhang"].Value.ToString();
			if (((Control)txtMa).Text == "pallet")
			{
				((Control)txtMa).Enabled = false;
			}
			else
			{
				((Control)txtMa).Enabled = true;
			}
		}
		catch
		{
			gID = "0";
		}
	}

	private void txtName_KeyPress(object sender, KeyPressEventArgs e)
	{
		if (e.KeyChar == '\r' && checkForm())
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
		}
	}

	private void txtTim_TextChanged(object sender, EventArgs e)
	{
		try
		{
			showData(((Control)txtTim).Text);
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
		((Form)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Expected O, but got Unknown
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Expected O, but got Unknown
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Expected O, but got Unknown
		//IL_0061: Unknown result type (might be due to invalid IL or missing references)
		//IL_006b: Expected O, but got Unknown
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Expected O, but got Unknown
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_0081: Expected O, but got Unknown
		//IL_0082: Unknown result type (might be due to invalid IL or missing references)
		//IL_008c: Expected O, but got Unknown
		//IL_008d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0097: Expected O, but got Unknown
		//IL_0098: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a2: Expected O, but got Unknown
		//IL_00a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ad: Expected O, but got Unknown
		//IL_00ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b8: Expected O, but got Unknown
		//IL_00b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c3: Expected O, but got Unknown
		//IL_00c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ce: Expected O, but got Unknown
		//IL_00cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d9: Expected O, but got Unknown
		//IL_00da: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e4: Expected O, but got Unknown
		//IL_00e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ef: Expected O, but got Unknown
		//IL_01b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c0: Expected O, but got Unknown
		//IL_0245: Unknown result type (might be due to invalid IL or missing references)
		//IL_024f: Expected O, but got Unknown
		//IL_02ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b5: Expected O, but got Unknown
		//IL_02d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e0: Expected O, but got Unknown
		//IL_0351: Unknown result type (might be due to invalid IL or missing references)
		//IL_035b: Expected O, but got Unknown
		//IL_03ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d4: Expected O, but got Unknown
		//IL_0445: Unknown result type (might be due to invalid IL or missing references)
		//IL_044f: Expected O, but got Unknown
		//IL_04da: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e4: Expected O, but got Unknown
		//IL_0577: Unknown result type (might be due to invalid IL or missing references)
		//IL_0581: Expected O, but got Unknown
		//IL_065a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0664: Expected O, but got Unknown
		//IL_073d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0747: Expected O, but got Unknown
		//IL_088e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0898: Expected O, but got Unknown
		//IL_0bd8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0be2: Expected O, but got Unknown
		//IL_0bf7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c01: Expected O, but got Unknown
		//IL_0c04: Unknown result type (might be due to invalid IL or missing references)
		components = new Container();
		DataGridViewCellStyle val = new DataGridViewCellStyle();
		DataGridViewCellStyle val2 = new DataGridViewCellStyle();
		DataGridViewCellStyle val3 = new DataGridViewCellStyle();
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmHang));
		lblMessage = new Label();
		nameError = new ErrorProvider(components);
		label6 = new Label();
		txtName = new TextBox();
		label5 = new Label();
		txtMa = new TextBox();
		label1 = new Label();
		txtTim = new TextBox();
		label4 = new Label();
		btnThoat = new Button();
		btnXoa = new Button();
		btnSave = new Button();
		dataGrid = new DataGridView();
		id = new DataGridViewTextBoxColumn();
		stt = new DataGridViewTextBoxColumn();
		mahang = new DataGridViewTextBoxColumn();
		tenhang = new DataGridViewTextBoxColumn();
		((ISupportInitialize)nameError).BeginInit();
		((ISupportInitialize)dataGrid).BeginInit();
		((Control)this).SuspendLayout();
		((Control)lblMessage).AutoSize = true;
		((Control)lblMessage).ForeColor = Color.Red;
		((Control)lblMessage).Location = new Point(232, 42);
		((Control)lblMessage).Name = "lblMessage";
		((Control)lblMessage).Size = new Size(69, 18);
		((Control)lblMessage).TabIndex = 71;
		((Control)lblMessage).Text = "message";
		nameError.ContainerControl = (ContainerControl)(object)this;
		((Control)label6).AutoSize = true;
		((Control)label6).Font = new Font("Microsoft Sans Serif", 21.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label6).ForeColor = Color.Black;
		((Control)label6).Location = new Point(229, 9);
		((Control)label6).Name = "label6";
		((Control)label6).Size = new Size(361, 33);
		((Control)label6).TabIndex = 66;
		((Control)label6).Text = "DANH SÁCH SẢN PHẨM";
		((Control)txtName).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtName).Location = new Point(235, 99);
		((Control)txtName).Name = "txtName";
		((Control)txtName).Size = new Size(422, 26);
		((Control)txtName).TabIndex = 2;
		((Control)txtName).KeyPress += new KeyPressEventHandler(txtName_KeyPress);
		((Control)label5).AutoSize = true;
		((Control)label5).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label5).Location = new Point(162, 102);
		((Control)label5).Name = "label5";
		((Control)label5).Size = new Size(76, 20);
		((Control)label5).TabIndex = 64;
		((Control)label5).Text = "Tên hàng";
		((Control)txtMa).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtMa).Location = new Point(235, 69);
		((Control)txtMa).Name = "txtMa";
		((Control)txtMa).Size = new Size(422, 26);
		((Control)txtMa).TabIndex = 1;
		((Control)label1).AutoSize = true;
		((Control)label1).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label1).Location = new Point(162, 72);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(71, 20);
		((Control)label1).TabIndex = 75;
		((Control)label1).Text = "Mã hàng";
		((Control)txtTim).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)txtTim).Location = new Point(235, 178);
		((Control)txtTim).Name = "txtTim";
		((Control)txtTim).Size = new Size(422, 26);
		((Control)txtTim).TabIndex = 9;
		((Control)txtTim).TextChanged += txtTim_TextChanged;
		((Control)label4).AutoSize = true;
		((Control)label4).Font = new Font("Microsoft Sans Serif", 12f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Control)label4).Location = new Point(71, 181);
		((Control)label4).Name = "label4";
		((Control)label4).Size = new Size(162, 20);
		((Control)label4).TabIndex = 80;
		((Control)label4).Text = "Tìm theo mã hoặc tên";
		((Control)btnThoat).BackgroundImageLayout = (ImageLayout)0;
		((ButtonBase)btnThoat).FlatAppearance.BorderSize = 0;
		((Control)btnThoat).Font = new Font("Arial", 12f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ButtonBase)btnThoat).Image = (Image)(object)Resources.exit;
		((ButtonBase)btnThoat).ImageAlign = (ContentAlignment)16;
		((Control)btnThoat).Location = new Point(520, 131);
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
		((Control)btnXoa).Location = new Point(377, 131);
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
		((Control)btnSave).Location = new Point(234, 131);
		((Control)btnSave).Name = "btnSave";
		((Control)btnSave).Size = new Size(137, 41);
		((Control)btnSave).TabIndex = 5;
		((Control)btnSave).Text = "       Lưu";
		((ButtonBase)btnSave).UseVisualStyleBackColor = true;
		((Control)btnSave).Click += btnSave_Click;
		dataGrid.AllowUserToAddRows = false;
		dataGrid.AllowUserToOrderColumns = true;
		dataGrid.AllowUserToResizeRows = false;
		dataGrid.BorderStyle = (BorderStyle)2;
		dataGrid.ColumnHeadersHeightSizeMode = (DataGridViewColumnHeadersHeightSizeMode)2;
		dataGrid.Columns.AddRange((DataGridViewColumn[])(object)new DataGridViewColumn[4]
		{
			(DataGridViewColumn)id,
			(DataGridViewColumn)stt,
			(DataGridViewColumn)mahang,
			(DataGridViewColumn)tenhang
		});
		((Control)dataGrid).DataBindings.Add(new Binding("Font", (object)Settings.Default, "font12", true, (DataSourceUpdateMode)1));
		dataGrid.EditMode = (DataGridViewEditMode)4;
		((Control)dataGrid).Font = Settings.Default.font12;
		((Control)dataGrid).Location = new Point(3, 210);
		dataGrid.MultiSelect = false;
		((Control)dataGrid).Name = "dataGrid";
		dataGrid.RowTemplate.Height = 25;
		dataGrid.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)dataGrid).Size = new Size(808, 488);
		((Control)dataGrid).TabIndex = 10;
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
		((DataGridViewBand)mahang).DefaultCellStyle = val2;
		((DataGridViewColumn)mahang).FillWeight = 107.8173f;
		((DataGridViewColumn)mahang).HeaderText = "Mã hàng";
		((DataGridViewColumn)mahang).Name = "mahang";
		((DataGridViewColumn)mahang).Width = 177;
		((DataGridViewColumn)tenhang).AutoSizeMode = (DataGridViewAutoSizeColumnMode)1;
		val3.Alignment = (DataGridViewContentAlignment)16;
		((DataGridViewBand)tenhang).DefaultCellStyle = val3;
		((DataGridViewColumn)tenhang).FillWeight = 107.8173f;
		((DataGridViewColumn)tenhang).HeaderText = "Tên hàng";
		((DataGridViewColumn)tenhang).Name = "tenhang";
		((DataGridViewColumn)tenhang).Width = 400;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(9f, 18f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(815, 702);
		((Control)this).Controls.Add((Control)(object)txtTim);
		((Control)this).Controls.Add((Control)(object)label4);
		((Control)this).Controls.Add((Control)(object)txtMa);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)btnThoat);
		((Control)this).Controls.Add((Control)(object)btnXoa);
		((Control)this).Controls.Add((Control)(object)btnSave);
		((Control)this).Controls.Add((Control)(object)lblMessage);
		((Control)this).Controls.Add((Control)(object)dataGrid);
		((Control)this).Controls.Add((Control)(object)label6);
		((Control)this).Controls.Add((Control)(object)txtName);
		((Control)this).Controls.Add((Control)(object)label5);
		((Control)this).Font = new Font("Microsoft Sans Serif", 11.25f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmHang";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Danh sách sản phẩm";
		((Form)this).Load += frmHang_Load;
		((ISupportInitialize)nameError).EndInit();
		((ISupportInitialize)dataGrid).EndInit();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
