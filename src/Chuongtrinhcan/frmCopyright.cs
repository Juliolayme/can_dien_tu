using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Chuongtrinhcan;

public class frmCopyright : Form
{
	private DataFetcher dataFetcher = DataFetcher.Instance;

	private IContainer components = null;

	private Label label1;

	private Label lblThongso;

	private Label label2;

	private TextBox textBox1;

	private Button button1;

	public frmCopyright()
	{
		InitializeComponent();
	}

	private void button1_Click(object sender, EventArgs e)
	{
		//IL_0167: Unknown result type (might be due to invalid IL or missing references)
		//IL_0159: Unknown result type (might be due to invalid IL or missing references)
		string text = "";
		string text2 = "";
		DateTime now = DateTime.Now;
		int year = now.Year;
		int month = now.Month;
		int day = now.Day;
		text2 = ((Control)lblThongso).Text;
		text2 = text2.Substring(text2.Length - 6, 2);
		text = (year + month).ToString() + day + day % 2 + day % 5 + day % 7 + day % 8;
		double num = 0.0;
		string cPUID = DBUtil.getCPUID();
		num = DBUtil.ConvertToTotalNumber(cPUID);
		string text3 = num.ToString() + num % 2.0 + num % 5.0 + num % 7.0;
		if (((Control)textBox1).Text == "vietduc123")
		{
			Registry.SetValue("HKEY_CURRENT_USER", "cprVietDucAnhChi", "vietduc123");
			MessageBox.Show("Key OK! Vui lòng chạy lại chương trình.");
		}
		else
		{
			MessageBox.Show("Sai key.");
		}
	}

	private void frmCopyright_Load(object sender, EventArgs e)
	{
		((Control)lblThongso).Text = DateTime.Now.ToString() + " - " + DBUtil.getCPUID();
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
		//IL_0314: Unknown result type (might be due to invalid IL or missing references)
		//IL_031e: Expected O, but got Unknown
		//IL_0333: Unknown result type (might be due to invalid IL or missing references)
		//IL_033d: Expected O, but got Unknown
		//IL_0340: Unknown result type (might be due to invalid IL or missing references)
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmCopyright));
		label1 = new Label();
		lblThongso = new Label();
		label2 = new Label();
		textBox1 = new TextBox();
		button1 = new Button();
		((Control)this).SuspendLayout();
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(13, 26);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(119, 16);
		((Control)label1).TabIndex = 1;
		((Control)label1).Text = "Thông số hệ thống";
		((Control)lblThongso).AutoSize = true;
		((Control)lblThongso).Location = new Point(138, 26);
		((Control)lblThongso).Name = "lblThongso";
		((Control)lblThongso).Size = new Size(76, 16);
		((Control)lblThongso).TabIndex = 2;
		((Control)lblThongso).Text = "lblThongso";
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(13, 55);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(67, 16);
		((Control)label2).TabIndex = 3;
		((Control)label2).Text = "Nhập Key";
		((Control)textBox1).Location = new Point(141, 55);
		((Control)textBox1).Name = "textBox1";
		textBox1.PasswordChar = '*';
		((Control)textBox1).Size = new Size(385, 22);
		((Control)textBox1).TabIndex = 4;
		((Control)button1).Location = new Point(141, 83);
		((Control)button1).Name = "button1";
		((Control)button1).Size = new Size(101, 38);
		((Control)button1).TabIndex = 5;
		((Control)button1).Text = "OK";
		((ButtonBase)button1).UseVisualStyleBackColor = true;
		((Control)button1).Click += button1_Click;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(0, 192, 192);
		((Form)this).ClientSize = new Size(537, 141);
		((Control)this).Controls.Add((Control)(object)button1);
		((Control)this).Controls.Add((Control)(object)textBox1);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)lblThongso);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Font = new Font("Microsoft Sans Serif", 9.75f, (FontStyle)0, (GraphicsUnit)3, (byte)0);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Control)this).Name = "frmCopyright";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Thông số hệ thống";
		((Form)this).Load += frmCopyright_Load;
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
