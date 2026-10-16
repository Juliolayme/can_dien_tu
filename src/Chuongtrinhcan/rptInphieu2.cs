using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using DDCssLib;
using DataDynamics.ActiveReports;
using DataDynamics.ActiveReports.DataSources;

namespace Chuongtrinhcan;

public class rptInphieu2 : ActiveReport3
{
	private Detail detail;

	private ReportHeader reportHeader1;

	private TextBox txtCty;

	private TextBox txtDiachi;

	private TextBox txtDienthoai;

	private TextBox textBox4;

	private Label label1;

	private Label txtSophieu;

	private ReportFooter reportFooter1;

	private PageHeader pageHeader1;

	private PageFooter pageFooter1;

	private Label label3;

	private Label label4;

	private Label label5;

	private Label label11;

	private Label label12;

	private Label label13;

	private TextBox txtSoxe;

	private TextBox txtKhachhang;

	private TextBox txtGhichu;

	private Label label26;

	private TextBox txtNgaycan;

	private TextBox txtKyten;

	private Label label2;

	private Label label6;

	private Label label7;

	private Label label8;

	private Label label9;

	private Label label15;

	private Label label17;

	private Label label18;

	private Label label36;

	private Label label37;

	private Label label38;

	private Label label39;

	private Label label40;

	private Label label41;

	private Label label42;

	private Label label43;

	private TextBox txtTl1;

	private TextBox txtSolong1;

	private TextBox txtSocon1;

	private TextBox txtSTT1;

	private TextBox txtTl2;

	private TextBox txtSolong2;

	private TextBox txtSocon2;

	private TextBox txtSTT2;

	private TextBox txtTL3;

	private TextBox txtSolong3;

	private TextBox txtSocon3;

	private TextBox txtSTT3;

	private TextBox txtTL4;

	private TextBox txtSolong4;

	private TextBox txtSocon4;

	private TextBox txtSTT4;

	private TextBox txtTotalSomacan;

	private TextBox txtTotalTLgalong;

	private Label label14;

	private TextBox txtTotalSolong;

	private Label label45;

	private TextBox txtTotalTLcabi;

	private Label label46;

	private TextBox txtTotalSLgalong;

	private TextBox txtTlBinhquan;

	private Label label47;

	private Label label48;

	private Label label49;

	private TextBox txtTLBi;

	private Label label50;

	private Label label51;

	private Label label52;

	private Label label53;

	private Label label54;

	private Label label55;

	private Label label57;

	private Label label58;

	private GroupHeader groupHeader1;

	private GroupFooter groupFooter1;

	private TextBox txtLoaican;

	private TextBox txtTotalSomacanBi;

	private Label label10;

	private Label label16;

	private DataFetcher dataFetcher = DataFetcher.Instance;

	private OleDBDataSource myDatasource;

	public string rptSophieu;

	private double totalSL;

	private double totalSolong;

	private double totalSocon;

	private double totalSoMacan;

	private double tlLong = 0.0;

	private double tlBinhquanLong;

	protected override void Dispose(bool disposing)
	{
		if (disposing)
		{
		}
		((ActiveReport3)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Expected O, but got Unknown
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_0022: Expected O, but got Unknown
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_002d: Expected O, but got Unknown
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected O, but got Unknown
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Expected O, but got Unknown
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Expected O, but got Unknown
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0085: Expected O, but got Unknown
		//IL_0086: Unknown result type (might be due to invalid IL or missing references)
		//IL_0090: Expected O, but got Unknown
		//IL_0091: Unknown result type (might be due to invalid IL or missing references)
		//IL_009b: Expected O, but got Unknown
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a6: Expected O, but got Unknown
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Expected O, but got Unknown
		//IL_00b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bc: Expected O, but got Unknown
		//IL_00bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c7: Expected O, but got Unknown
		//IL_00c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d2: Expected O, but got Unknown
		//IL_00d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00dd: Expected O, but got Unknown
		//IL_00de: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e8: Expected O, but got Unknown
		//IL_00e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f3: Expected O, but got Unknown
		//IL_00f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fe: Expected O, but got Unknown
		//IL_00ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_0109: Expected O, but got Unknown
		//IL_010a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0114: Expected O, but got Unknown
		//IL_0115: Unknown result type (might be due to invalid IL or missing references)
		//IL_011f: Expected O, but got Unknown
		//IL_0120: Unknown result type (might be due to invalid IL or missing references)
		//IL_012a: Expected O, but got Unknown
		//IL_012b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0135: Expected O, but got Unknown
		//IL_0136: Unknown result type (might be due to invalid IL or missing references)
		//IL_0140: Expected O, but got Unknown
		//IL_0141: Unknown result type (might be due to invalid IL or missing references)
		//IL_014b: Expected O, but got Unknown
		//IL_014c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0156: Expected O, but got Unknown
		//IL_0157: Unknown result type (might be due to invalid IL or missing references)
		//IL_0161: Expected O, but got Unknown
		//IL_0162: Unknown result type (might be due to invalid IL or missing references)
		//IL_016c: Expected O, but got Unknown
		//IL_016d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0177: Expected O, but got Unknown
		//IL_0178: Unknown result type (might be due to invalid IL or missing references)
		//IL_0182: Expected O, but got Unknown
		//IL_0183: Unknown result type (might be due to invalid IL or missing references)
		//IL_018d: Expected O, but got Unknown
		//IL_018e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0198: Expected O, but got Unknown
		//IL_0199: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a3: Expected O, but got Unknown
		//IL_01a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ae: Expected O, but got Unknown
		//IL_01af: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b9: Expected O, but got Unknown
		//IL_01ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c4: Expected O, but got Unknown
		//IL_01c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cf: Expected O, but got Unknown
		//IL_01d0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01da: Expected O, but got Unknown
		//IL_01db: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e5: Expected O, but got Unknown
		//IL_01e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f0: Expected O, but got Unknown
		//IL_01f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fb: Expected O, but got Unknown
		//IL_01fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0206: Expected O, but got Unknown
		//IL_0207: Unknown result type (might be due to invalid IL or missing references)
		//IL_0211: Expected O, but got Unknown
		//IL_0212: Unknown result type (might be due to invalid IL or missing references)
		//IL_021c: Expected O, but got Unknown
		//IL_021d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0227: Expected O, but got Unknown
		//IL_0228: Unknown result type (might be due to invalid IL or missing references)
		//IL_0232: Expected O, but got Unknown
		//IL_0233: Unknown result type (might be due to invalid IL or missing references)
		//IL_023d: Expected O, but got Unknown
		//IL_023e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0248: Expected O, but got Unknown
		//IL_0249: Unknown result type (might be due to invalid IL or missing references)
		//IL_0253: Expected O, but got Unknown
		//IL_0254: Unknown result type (might be due to invalid IL or missing references)
		//IL_025e: Expected O, but got Unknown
		//IL_025f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0269: Expected O, but got Unknown
		//IL_026a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0274: Expected O, but got Unknown
		//IL_0275: Unknown result type (might be due to invalid IL or missing references)
		//IL_027f: Expected O, but got Unknown
		//IL_0280: Unknown result type (might be due to invalid IL or missing references)
		//IL_028a: Expected O, but got Unknown
		//IL_028b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0295: Expected O, but got Unknown
		//IL_0296: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a0: Expected O, but got Unknown
		//IL_02a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ab: Expected O, but got Unknown
		//IL_02ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b6: Expected O, but got Unknown
		//IL_02b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c1: Expected O, but got Unknown
		//IL_02c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_02cc: Expected O, but got Unknown
		//IL_02cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d7: Expected O, but got Unknown
		//IL_02d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e2: Expected O, but got Unknown
		//IL_02e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ed: Expected O, but got Unknown
		//IL_02ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f8: Expected O, but got Unknown
		//IL_02f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0303: Expected O, but got Unknown
		//IL_0304: Unknown result type (might be due to invalid IL or missing references)
		//IL_030e: Expected O, but got Unknown
		//IL_030f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0319: Expected O, but got Unknown
		//IL_031a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0324: Expected O, but got Unknown
		//IL_0325: Unknown result type (might be due to invalid IL or missing references)
		//IL_032f: Expected O, but got Unknown
		//IL_0330: Unknown result type (might be due to invalid IL or missing references)
		//IL_033a: Expected O, but got Unknown
		//IL_033b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0345: Expected O, but got Unknown
		//IL_0346: Unknown result type (might be due to invalid IL or missing references)
		//IL_0350: Expected O, but got Unknown
		//IL_0351: Unknown result type (might be due to invalid IL or missing references)
		//IL_035b: Expected O, but got Unknown
		//IL_035c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0366: Expected O, but got Unknown
		//IL_0367: Unknown result type (might be due to invalid IL or missing references)
		//IL_0371: Expected O, but got Unknown
		//IL_0372: Unknown result type (might be due to invalid IL or missing references)
		//IL_037c: Expected O, but got Unknown
		//IL_037d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0387: Expected O, but got Unknown
		//IL_0388: Unknown result type (might be due to invalid IL or missing references)
		//IL_0392: Expected O, but got Unknown
		//IL_0393: Unknown result type (might be due to invalid IL or missing references)
		//IL_039d: Expected O, but got Unknown
		//IL_5e36: Unknown result type (might be due to invalid IL or missing references)
		//IL_5e40: Expected O, but got Unknown
		//IL_5e56: Unknown result type (might be due to invalid IL or missing references)
		//IL_5e60: Expected O, but got Unknown
		//IL_5e76: Unknown result type (might be due to invalid IL or missing references)
		//IL_5e80: Expected O, but got Unknown
		//IL_5e96: Unknown result type (might be due to invalid IL or missing references)
		//IL_5ea0: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(rptInphieu2));
		OleDBDataSource val = new OleDBDataSource();
		txtCty = new TextBox();
		txtDiachi = new TextBox();
		txtDienthoai = new TextBox();
		textBox4 = new TextBox();
		label1 = new Label();
		txtSophieu = new Label();
		label3 = new Label();
		label4 = new Label();
		label5 = new Label();
		label11 = new Label();
		label12 = new Label();
		label13 = new Label();
		txtSoxe = new TextBox();
		txtKhachhang = new TextBox();
		txtGhichu = new TextBox();
		label26 = new Label();
		txtNgaycan = new TextBox();
		txtKyten = new TextBox();
		label2 = new Label();
		label6 = new Label();
		label7 = new Label();
		label8 = new Label();
		label9 = new Label();
		label15 = new Label();
		label17 = new Label();
		label18 = new Label();
		label36 = new Label();
		label37 = new Label();
		label38 = new Label();
		label39 = new Label();
		label40 = new Label();
		label41 = new Label();
		label42 = new Label();
		label43 = new Label();
		txtTl1 = new TextBox();
		txtSolong1 = new TextBox();
		txtSocon1 = new TextBox();
		txtSTT1 = new TextBox();
		txtTl2 = new TextBox();
		txtSolong2 = new TextBox();
		txtSocon2 = new TextBox();
		txtSTT2 = new TextBox();
		txtTL3 = new TextBox();
		txtSolong3 = new TextBox();
		txtSocon3 = new TextBox();
		txtSTT3 = new TextBox();
		txtTL4 = new TextBox();
		txtSolong4 = new TextBox();
		txtSocon4 = new TextBox();
		txtSTT4 = new TextBox();
		txtTotalSomacan = new TextBox();
		txtTotalTLgalong = new TextBox();
		label14 = new Label();
		txtTotalSolong = new TextBox();
		label45 = new Label();
		txtTotalTLcabi = new TextBox();
		label46 = new Label();
		txtTotalSLgalong = new TextBox();
		txtTlBinhquan = new TextBox();
		label47 = new Label();
		label48 = new Label();
		label49 = new Label();
		txtTLBi = new TextBox();
		label50 = new Label();
		label51 = new Label();
		label52 = new Label();
		label53 = new Label();
		label54 = new Label();
		label55 = new Label();
		label57 = new Label();
		label58 = new Label();
		txtLoaican = new TextBox();
		txtTotalSomacanBi = new TextBox();
		label10 = new Label();
		label16 = new Label();
		detail = new Detail();
		reportHeader1 = new ReportHeader();
		reportFooter1 = new ReportFooter();
		pageHeader1 = new PageHeader();
		pageFooter1 = new PageFooter();
		groupHeader1 = new GroupHeader();
		groupFooter1 = new GroupFooter();
		((ISupportInitialize)txtCty).BeginInit();
		((ISupportInitialize)txtDiachi).BeginInit();
		((ISupportInitialize)txtDienthoai).BeginInit();
		((ISupportInitialize)textBox4).BeginInit();
		((ISupportInitialize)label1).BeginInit();
		((ISupportInitialize)txtSophieu).BeginInit();
		((ISupportInitialize)label3).BeginInit();
		((ISupportInitialize)label4).BeginInit();
		((ISupportInitialize)label5).BeginInit();
		((ISupportInitialize)label11).BeginInit();
		((ISupportInitialize)label12).BeginInit();
		((ISupportInitialize)label13).BeginInit();
		((ISupportInitialize)txtSoxe).BeginInit();
		((ISupportInitialize)txtKhachhang).BeginInit();
		((ISupportInitialize)txtGhichu).BeginInit();
		((ISupportInitialize)label26).BeginInit();
		((ISupportInitialize)txtNgaycan).BeginInit();
		((ISupportInitialize)txtKyten).BeginInit();
		((ISupportInitialize)label2).BeginInit();
		((ISupportInitialize)label6).BeginInit();
		((ISupportInitialize)label7).BeginInit();
		((ISupportInitialize)label8).BeginInit();
		((ISupportInitialize)label9).BeginInit();
		((ISupportInitialize)label15).BeginInit();
		((ISupportInitialize)label17).BeginInit();
		((ISupportInitialize)label18).BeginInit();
		((ISupportInitialize)label36).BeginInit();
		((ISupportInitialize)label37).BeginInit();
		((ISupportInitialize)label38).BeginInit();
		((ISupportInitialize)label39).BeginInit();
		((ISupportInitialize)label40).BeginInit();
		((ISupportInitialize)label41).BeginInit();
		((ISupportInitialize)label42).BeginInit();
		((ISupportInitialize)label43).BeginInit();
		((ISupportInitialize)txtTl1).BeginInit();
		((ISupportInitialize)txtSolong1).BeginInit();
		((ISupportInitialize)txtSocon1).BeginInit();
		((ISupportInitialize)txtSTT1).BeginInit();
		((ISupportInitialize)txtTl2).BeginInit();
		((ISupportInitialize)txtSolong2).BeginInit();
		((ISupportInitialize)txtSocon2).BeginInit();
		((ISupportInitialize)txtSTT2).BeginInit();
		((ISupportInitialize)txtTL3).BeginInit();
		((ISupportInitialize)txtSolong3).BeginInit();
		((ISupportInitialize)txtSocon3).BeginInit();
		((ISupportInitialize)txtSTT3).BeginInit();
		((ISupportInitialize)txtTL4).BeginInit();
		((ISupportInitialize)txtSolong4).BeginInit();
		((ISupportInitialize)txtSocon4).BeginInit();
		((ISupportInitialize)txtSTT4).BeginInit();
		((ISupportInitialize)txtTotalSomacan).BeginInit();
		((ISupportInitialize)txtTotalTLgalong).BeginInit();
		((ISupportInitialize)label14).BeginInit();
		((ISupportInitialize)txtTotalSolong).BeginInit();
		((ISupportInitialize)label45).BeginInit();
		((ISupportInitialize)txtTotalTLcabi).BeginInit();
		((ISupportInitialize)label46).BeginInit();
		((ISupportInitialize)txtTotalSLgalong).BeginInit();
		((ISupportInitialize)txtTlBinhquan).BeginInit();
		((ISupportInitialize)label47).BeginInit();
		((ISupportInitialize)label48).BeginInit();
		((ISupportInitialize)label49).BeginInit();
		((ISupportInitialize)txtTLBi).BeginInit();
		((ISupportInitialize)label50).BeginInit();
		((ISupportInitialize)label51).BeginInit();
		((ISupportInitialize)label52).BeginInit();
		((ISupportInitialize)label53).BeginInit();
		((ISupportInitialize)label54).BeginInit();
		((ISupportInitialize)label55).BeginInit();
		((ISupportInitialize)label57).BeginInit();
		((ISupportInitialize)label58).BeginInit();
		((ISupportInitialize)txtLoaican).BeginInit();
		((ISupportInitialize)txtTotalSomacanBi).BeginInit();
		((ISupportInitialize)label10).BeginInit();
		((ISupportInitialize)label16).BeginInit();
		((ISupportInitialize)this).BeginInit();
		((ARControl)txtCty).Border.BottomColor = Color.Black;
		((ARControl)txtCty).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtCty).Border.LeftColor = Color.Black;
		((ARControl)txtCty).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtCty).Border.RightColor = Color.Black;
		((ARControl)txtCty).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtCty).Border.TopColor = Color.Black;
		((ARControl)txtCty).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtCty).Height = 0.25f;
		((ARControl)txtCty).Left = 0.0625f;
		((ARControl)txtCty).Name = "txtCty";
		componentResourceManager.ApplyResources(txtCty, "txtCty");
		((ARControl)txtCty).Top = 0.0625f;
		((ARControl)txtCty).Width = 6.875f;
		((ARControl)txtDiachi).Border.BottomColor = Color.Black;
		((ARControl)txtDiachi).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtDiachi).Border.LeftColor = Color.Black;
		((ARControl)txtDiachi).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtDiachi).Border.RightColor = Color.Black;
		((ARControl)txtDiachi).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtDiachi).Border.TopColor = Color.Black;
		((ARControl)txtDiachi).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtDiachi).Height = 0.25f;
		((ARControl)txtDiachi).Left = 0.0625f;
		((ARControl)txtDiachi).Name = "txtDiachi";
		componentResourceManager.ApplyResources(txtDiachi, "txtDiachi");
		((ARControl)txtDiachi).Top = 0.3125f;
		((ARControl)txtDiachi).Width = 6.875f;
		((ARControl)txtDienthoai).Border.BottomColor = Color.Black;
		((ARControl)txtDienthoai).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtDienthoai).Border.LeftColor = Color.Black;
		((ARControl)txtDienthoai).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtDienthoai).Border.RightColor = Color.Black;
		((ARControl)txtDienthoai).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtDienthoai).Border.TopColor = Color.Black;
		((ARControl)txtDienthoai).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtDienthoai).Height = 0.25f;
		((ARControl)txtDienthoai).Left = 0.0625f;
		((ARControl)txtDienthoai).Name = "txtDienthoai";
		componentResourceManager.ApplyResources(txtDienthoai, "txtDienthoai");
		txtDienthoai.Text = null;
		((ARControl)txtDienthoai).Top = 0.5625f;
		((ARControl)txtDienthoai).Width = 6.875f;
		((ARControl)textBox4).Border.BottomColor = Color.Black;
		((ARControl)textBox4).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)textBox4).Border.LeftColor = Color.Black;
		((ARControl)textBox4).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)textBox4).Border.RightColor = Color.Black;
		((ARControl)textBox4).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)textBox4).Border.TopColor = Color.Black;
		((ARControl)textBox4).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)textBox4).Height = 0.375f;
		((ARControl)textBox4).Left = 2.4375f;
		((ARControl)textBox4).Name = "textBox4";
		componentResourceManager.ApplyResources(textBox4, "textBox4");
		((ARControl)textBox4).Top = 0.8125f;
		((ARControl)textBox4).Width = 2.5f;
		((ARControl)label1).Border.BottomColor = Color.Black;
		((ARControl)label1).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label1).Border.LeftColor = Color.Black;
		((ARControl)label1).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label1).Border.RightColor = Color.Black;
		((ARControl)label1).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label1).Border.TopColor = Color.Black;
		((ARControl)label1).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label1).Height = 0.25f;
		label1.HyperLink = null;
		((ARControl)label1).Left = 5.1875f;
		((ARControl)label1).Name = "label1";
		componentResourceManager.ApplyResources(label1, "label1");
		((ARControl)label1).Top = 1.1875f;
		((ARControl)label1).Width = 0.6875f;
		((ARControl)txtSophieu).Border.BottomColor = Color.Black;
		((ARControl)txtSophieu).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtSophieu).Border.LeftColor = Color.Black;
		((ARControl)txtSophieu).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtSophieu).Border.RightColor = Color.Black;
		((ARControl)txtSophieu).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSophieu).Border.TopColor = Color.Black;
		((ARControl)txtSophieu).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSophieu).DataField = "sophieu";
		((ARControl)txtSophieu).Height = 0.25f;
		txtSophieu.HyperLink = null;
		((ARControl)txtSophieu).Left = 6f;
		((ARControl)txtSophieu).Name = "txtSophieu";
		componentResourceManager.ApplyResources(txtSophieu, "txtSophieu");
		((ARControl)txtSophieu).Top = 1.1875f;
		((ARControl)txtSophieu).Width = 1.4375f;
		((ARControl)label3).Border.BottomColor = Color.Black;
		((ARControl)label3).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label3).Border.LeftColor = Color.Black;
		((ARControl)label3).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label3).Border.RightColor = Color.Black;
		((ARControl)label3).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label3).Border.TopColor = Color.Black;
		((ARControl)label3).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label3).Height = 0.25f;
		label3.HyperLink = null;
		((ARControl)label3).Left = 5.1875f;
		((ARControl)label3).Name = "label3";
		componentResourceManager.ApplyResources(label3, "label3");
		((ARControl)label3).Top = 1.4375f;
		((ARControl)label3).Width = 0.6875f;
		((ARControl)label4).Border.BottomColor = Color.Black;
		((ARControl)label4).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label4).Border.LeftColor = Color.Black;
		((ARControl)label4).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label4).Border.RightColor = Color.Black;
		((ARControl)label4).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label4).Border.TopColor = Color.Black;
		((ARControl)label4).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label4).Height = 0.25f;
		label4.HyperLink = null;
		((ARControl)label4).Left = 0.0625f;
		((ARControl)label4).Name = "label4";
		componentResourceManager.ApplyResources(label4, "label4");
		((ARControl)label4).Top = 1.1875f;
		((ARControl)label4).Width = 1.1875f;
		((ARControl)label5).Border.BottomColor = Color.Black;
		((ARControl)label5).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label5).Border.LeftColor = Color.Black;
		((ARControl)label5).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label5).Border.RightColor = Color.Black;
		((ARControl)label5).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label5).Border.TopColor = Color.Black;
		((ARControl)label5).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label5).Height = 0.25f;
		label5.HyperLink = null;
		((ARControl)label5).Left = 0.0625f;
		((ARControl)label5).Name = "label5";
		componentResourceManager.ApplyResources(label5, "label5");
		((ARControl)label5).Top = 1.4375f;
		((ARControl)label5).Width = 1.1875f;
		((ARControl)label11).Border.BottomColor = Color.Black;
		((ARControl)label11).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label11).Border.LeftColor = Color.Black;
		((ARControl)label11).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label11).Border.RightColor = Color.Black;
		((ARControl)label11).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label11).Border.TopColor = Color.Black;
		((ARControl)label11).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label11).Height = 0.25f;
		label11.HyperLink = null;
		((ARControl)label11).Left = 5.875f;
		((ARControl)label11).Name = "label11";
		componentResourceManager.ApplyResources(label11, "label11");
		((ARControl)label11).Top = 1.4375f;
		((ARControl)label11).Width = 0.125f;
		((ARControl)label12).Border.BottomColor = Color.Black;
		((ARControl)label12).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label12).Border.LeftColor = Color.Black;
		((ARControl)label12).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label12).Border.RightColor = Color.Black;
		((ARControl)label12).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label12).Border.TopColor = Color.Black;
		((ARControl)label12).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label12).Height = 0.25f;
		label12.HyperLink = null;
		((ARControl)label12).Left = 1.25f;
		((ARControl)label12).Name = "label12";
		componentResourceManager.ApplyResources(label12, "label12");
		((ARControl)label12).Top = 1.1875f;
		((ARControl)label12).Width = 0.125f;
		((ARControl)label13).Border.BottomColor = Color.Black;
		((ARControl)label13).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label13).Border.LeftColor = Color.Black;
		((ARControl)label13).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label13).Border.RightColor = Color.Black;
		((ARControl)label13).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label13).Border.TopColor = Color.Black;
		((ARControl)label13).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label13).Height = 0.25f;
		label13.HyperLink = null;
		((ARControl)label13).Left = 1.25f;
		((ARControl)label13).Name = "label13";
		componentResourceManager.ApplyResources(label13, "label13");
		((ARControl)label13).Top = 1.4375f;
		((ARControl)label13).Width = 0.125f;
		((ARControl)txtSoxe).Border.BottomColor = Color.Black;
		((ARControl)txtSoxe).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtSoxe).Border.LeftColor = Color.Black;
		((ARControl)txtSoxe).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtSoxe).Border.RightColor = Color.Black;
		((ARControl)txtSoxe).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSoxe).Border.TopColor = Color.Black;
		((ARControl)txtSoxe).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSoxe).Height = 0.25f;
		((ARControl)txtSoxe).Left = 6f;
		((ARControl)txtSoxe).Name = "txtSoxe";
		componentResourceManager.ApplyResources(txtSoxe, "txtSoxe");
		txtSoxe.Text = null;
		((ARControl)txtSoxe).Top = 1.4375f;
		((ARControl)txtSoxe).Width = 1.4375f;
		((ARControl)txtKhachhang).Border.BottomColor = Color.Black;
		((ARControl)txtKhachhang).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtKhachhang).Border.LeftColor = Color.Black;
		((ARControl)txtKhachhang).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtKhachhang).Border.RightColor = Color.Black;
		((ARControl)txtKhachhang).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtKhachhang).Border.TopColor = Color.Black;
		((ARControl)txtKhachhang).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtKhachhang).Height = 0.25f;
		((ARControl)txtKhachhang).Left = 1.375f;
		((ARControl)txtKhachhang).Name = "txtKhachhang";
		componentResourceManager.ApplyResources(txtKhachhang, "txtKhachhang");
		txtKhachhang.Text = null;
		((ARControl)txtKhachhang).Top = 1.1875f;
		((ARControl)txtKhachhang).Width = 3.75f;
		((ARControl)txtGhichu).Border.BottomColor = Color.Black;
		((ARControl)txtGhichu).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtGhichu).Border.LeftColor = Color.Black;
		((ARControl)txtGhichu).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtGhichu).Border.RightColor = Color.Black;
		((ARControl)txtGhichu).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtGhichu).Border.TopColor = Color.Black;
		((ARControl)txtGhichu).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtGhichu).Height = 0.25f;
		((ARControl)txtGhichu).Left = 1.375f;
		((ARControl)txtGhichu).Name = "txtGhichu";
		componentResourceManager.ApplyResources(txtGhichu, "txtGhichu");
		txtGhichu.Text = null;
		((ARControl)txtGhichu).Top = 1.4375f;
		((ARControl)txtGhichu).Width = 3.75f;
		((ARControl)label26).Border.BottomColor = Color.Black;
		((ARControl)label26).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label26).Border.LeftColor = Color.Black;
		((ARControl)label26).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label26).Border.RightColor = Color.Black;
		((ARControl)label26).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label26).Border.TopColor = Color.Black;
		((ARControl)label26).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label26).Height = 0.25f;
		label26.HyperLink = null;
		((ARControl)label26).Left = 5.875f;
		((ARControl)label26).Name = "label26";
		componentResourceManager.ApplyResources(label26, "label26");
		((ARControl)label26).Top = 1.1875f;
		((ARControl)label26).Width = 0.125f;
		((ARControl)txtNgaycan).Border.BottomColor = Color.Black;
		((ARControl)txtNgaycan).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtNgaycan).Border.LeftColor = Color.Black;
		((ARControl)txtNgaycan).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtNgaycan).Border.RightColor = Color.Black;
		((ARControl)txtNgaycan).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtNgaycan).Border.TopColor = Color.Black;
		((ARControl)txtNgaycan).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtNgaycan).Height = 0.25f;
		((ARControl)txtNgaycan).Left = 4.25f;
		((ARControl)txtNgaycan).Name = "txtNgaycan";
		componentResourceManager.ApplyResources(txtNgaycan, "txtNgaycan");
		txtNgaycan.Text = null;
		((ARControl)txtNgaycan).Top = 1.0625f;
		((ARControl)txtNgaycan).Width = 3.1875f;
		((ARControl)txtKyten).Border.BottomColor = Color.Black;
		((ARControl)txtKyten).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtKyten).Border.LeftColor = Color.Black;
		((ARControl)txtKyten).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtKyten).Border.RightColor = Color.Black;
		((ARControl)txtKyten).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtKyten).Border.TopColor = Color.Black;
		((ARControl)txtKyten).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtKyten).Height = 0.25f;
		((ARControl)txtKyten).Left = 0.0625f;
		((ARControl)txtKyten).Name = "txtKyten";
		componentResourceManager.ApplyResources(txtKyten, "txtKyten");
		((ARControl)txtKyten).Top = 1.3125f;
		((ARControl)txtKyten).Width = 7.375f;
		((ARControl)label2).Border.BottomColor = Color.Black;
		((ARControl)label2).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label2).Border.LeftColor = Color.Black;
		((ARControl)label2).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label2).Border.RightColor = Color.Black;
		((ARControl)label2).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label2).Border.TopColor = Color.Black;
		((ARControl)label2).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label2).Height = 0.375f;
		label2.HyperLink = null;
		((ARControl)label2).Left = 0.0625f;
		((ARControl)label2).Name = "label2";
		componentResourceManager.ApplyResources(label2, "label2");
		((ARControl)label2).Top = 0f;
		((ARControl)label2).Width = 0.4375f;
		((ARControl)label6).Border.BottomColor = Color.Black;
		((ARControl)label6).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label6).Border.LeftColor = Color.Black;
		((ARControl)label6).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label6).Border.RightColor = Color.Black;
		((ARControl)label6).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label6).Border.TopColor = Color.Black;
		((ARControl)label6).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label6).Height = 0.375f;
		label6.HyperLink = null;
		((ARControl)label6).Left = 1.0625f;
		((ARControl)label6).Name = "label6";
		componentResourceManager.ApplyResources(label6, "label6");
		((ARControl)label6).Top = 0f;
		((ARControl)label6).Width = 0.4375f;
		((ARControl)label7).Border.BottomColor = Color.Black;
		((ARControl)label7).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label7).Border.LeftColor = Color.Black;
		((ARControl)label7).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label7).Border.RightColor = Color.Black;
		((ARControl)label7).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label7).Border.TopColor = Color.Black;
		((ARControl)label7).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label7).Height = 0.375f;
		label7.HyperLink = null;
		((ARControl)label7).Left = 0.5f;
		((ARControl)label7).Name = "label7";
		componentResourceManager.ApplyResources(label7, "label7");
		((ARControl)label7).Top = 0f;
		((ARControl)label7).Width = 0.5625f;
		((ARControl)label8).Border.BottomColor = Color.Black;
		((ARControl)label8).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label8).Border.LeftColor = Color.Black;
		((ARControl)label8).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label8).Border.RightColor = Color.Black;
		((ARControl)label8).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label8).Border.TopColor = Color.Black;
		((ARControl)label8).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label8).Height = 0.375f;
		label8.HyperLink = null;
		((ARControl)label8).Left = 1.5f;
		((ARControl)label8).Name = "label8";
		componentResourceManager.ApplyResources(label8, "label8");
		((ARControl)label8).Top = 0f;
		((ARControl)label8).Width = 0.4375f;
		((ARControl)label9).Border.BottomColor = Color.Black;
		((ARControl)label9).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label9).Border.LeftColor = Color.Black;
		((ARControl)label9).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label9).Border.RightColor = Color.Black;
		((ARControl)label9).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label9).Border.TopColor = Color.Black;
		((ARControl)label9).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label9).Height = 0.375f;
		label9.HyperLink = null;
		((ARControl)label9).Left = 1.9375f;
		((ARControl)label9).Name = "label9";
		componentResourceManager.ApplyResources(label9, "label9");
		((ARControl)label9).Top = 0f;
		((ARControl)label9).Width = 0.4375f;
		((ARControl)label15).Border.BottomColor = Color.Black;
		((ARControl)label15).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label15).Border.LeftColor = Color.Black;
		((ARControl)label15).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label15).Border.RightColor = Color.Black;
		((ARControl)label15).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label15).Border.TopColor = Color.Black;
		((ARControl)label15).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label15).Height = 0.375f;
		label15.HyperLink = null;
		((ARControl)label15).Left = 2.9375f;
		((ARControl)label15).Name = "label15";
		componentResourceManager.ApplyResources(label15, "label15");
		((ARControl)label15).Top = 0f;
		((ARControl)label15).Width = 0.4375f;
		((ARControl)label17).Border.BottomColor = Color.Black;
		((ARControl)label17).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label17).Border.LeftColor = Color.Black;
		((ARControl)label17).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label17).Border.RightColor = Color.Black;
		((ARControl)label17).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label17).Border.TopColor = Color.Black;
		((ARControl)label17).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label17).Height = 0.375f;
		label17.HyperLink = null;
		((ARControl)label17).Left = 2.375f;
		((ARControl)label17).Name = "label17";
		componentResourceManager.ApplyResources(label17, "label17");
		((ARControl)label17).Top = 0f;
		((ARControl)label17).Width = 0.5625f;
		((ARControl)label18).Border.BottomColor = Color.Black;
		((ARControl)label18).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label18).Border.LeftColor = Color.Black;
		((ARControl)label18).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label18).Border.RightColor = Color.Black;
		((ARControl)label18).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label18).Border.TopColor = Color.Black;
		((ARControl)label18).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label18).Height = 0.375f;
		label18.HyperLink = null;
		((ARControl)label18).Left = 3.375f;
		((ARControl)label18).Name = "label18";
		componentResourceManager.ApplyResources(label18, "label18");
		((ARControl)label18).Top = 0f;
		((ARControl)label18).Width = 0.4375f;
		((ARControl)label36).Border.BottomColor = Color.Black;
		((ARControl)label36).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label36).Border.LeftColor = Color.Black;
		((ARControl)label36).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label36).Border.RightColor = Color.Black;
		((ARControl)label36).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label36).Border.TopColor = Color.Black;
		((ARControl)label36).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label36).Height = 0.375f;
		label36.HyperLink = null;
		((ARControl)label36).Left = 3.8125f;
		((ARControl)label36).Name = "label36";
		componentResourceManager.ApplyResources(label36, "label36");
		((ARControl)label36).Top = 0f;
		((ARControl)label36).Width = 0.4375f;
		((ARControl)label37).Border.BottomColor = Color.Black;
		((ARControl)label37).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label37).Border.LeftColor = Color.Black;
		((ARControl)label37).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label37).Border.RightColor = Color.Black;
		((ARControl)label37).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label37).Border.TopColor = Color.Black;
		((ARControl)label37).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label37).Height = 0.375f;
		label37.HyperLink = null;
		((ARControl)label37).Left = 4.8125f;
		((ARControl)label37).Name = "label37";
		componentResourceManager.ApplyResources(label37, "label37");
		((ARControl)label37).Top = 0f;
		((ARControl)label37).Width = 0.4375f;
		((ARControl)label38).Border.BottomColor = Color.Black;
		((ARControl)label38).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label38).Border.LeftColor = Color.Black;
		((ARControl)label38).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label38).Border.RightColor = Color.Black;
		((ARControl)label38).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label38).Border.TopColor = Color.Black;
		((ARControl)label38).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label38).Height = 0.375f;
		label38.HyperLink = null;
		((ARControl)label38).Left = 4.25f;
		((ARControl)label38).Name = "label38";
		componentResourceManager.ApplyResources(label38, "label38");
		((ARControl)label38).Top = 0f;
		((ARControl)label38).Width = 0.5625f;
		((ARControl)label39).Border.BottomColor = Color.Black;
		((ARControl)label39).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label39).Border.LeftColor = Color.Black;
		((ARControl)label39).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label39).Border.RightColor = Color.Black;
		((ARControl)label39).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label39).Border.TopColor = Color.Black;
		((ARControl)label39).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label39).Height = 0.375f;
		label39.HyperLink = null;
		((ARControl)label39).Left = 5.25f;
		((ARControl)label39).Name = "label39";
		componentResourceManager.ApplyResources(label39, "label39");
		((ARControl)label39).Top = 0f;
		((ARControl)label39).Width = 0.4375f;
		((ARControl)label40).Border.BottomColor = Color.Black;
		((ARControl)label40).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label40).Border.LeftColor = Color.Black;
		((ARControl)label40).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label40).Border.RightColor = Color.Black;
		((ARControl)label40).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label40).Border.TopColor = Color.Black;
		((ARControl)label40).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label40).Height = 0.375f;
		label40.HyperLink = null;
		((ARControl)label40).Left = 5.6875f;
		((ARControl)label40).Name = "label40";
		componentResourceManager.ApplyResources(label40, "label40");
		((ARControl)label40).Top = 0f;
		((ARControl)label40).Width = 0.4375f;
		((ARControl)label41).Border.BottomColor = Color.Black;
		((ARControl)label41).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label41).Border.LeftColor = Color.Black;
		((ARControl)label41).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label41).Border.RightColor = Color.Black;
		((ARControl)label41).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label41).Border.TopColor = Color.Black;
		((ARControl)label41).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label41).Height = 0.375f;
		label41.HyperLink = null;
		((ARControl)label41).Left = 6.6875f;
		((ARControl)label41).Name = "label41";
		componentResourceManager.ApplyResources(label41, "label41");
		((ARControl)label41).Top = 0f;
		((ARControl)label41).Width = 0.4375f;
		((ARControl)label42).Border.BottomColor = Color.Black;
		((ARControl)label42).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label42).Border.LeftColor = Color.Black;
		((ARControl)label42).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label42).Border.RightColor = Color.Black;
		((ARControl)label42).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label42).Border.TopColor = Color.Black;
		((ARControl)label42).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label42).Height = 0.375f;
		label42.HyperLink = null;
		((ARControl)label42).Left = 6.125f;
		((ARControl)label42).Name = "label42";
		componentResourceManager.ApplyResources(label42, "label42");
		((ARControl)label42).Top = 0f;
		((ARControl)label42).Width = 0.5625f;
		((ARControl)label43).Border.BottomColor = Color.Black;
		((ARControl)label43).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)label43).Border.LeftColor = Color.Black;
		((ARControl)label43).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)label43).Border.RightColor = Color.Black;
		((ARControl)label43).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)label43).Border.TopColor = Color.Black;
		((ARControl)label43).Border.TopStyle = (BorderLineStyle)1;
		((ARControl)label43).Height = 0.375f;
		label43.HyperLink = null;
		((ARControl)label43).Left = 7.125f;
		((ARControl)label43).Name = "label43";
		componentResourceManager.ApplyResources(label43, "label43");
		((ARControl)label43).Top = 0f;
		((ARControl)label43).Width = 0.4375f;
		((ARControl)txtTl1).Border.BottomColor = Color.Black;
		((ARControl)txtTl1).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtTl1).Border.LeftColor = Color.Black;
		((ARControl)txtTl1).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtTl1).Border.RightColor = Color.Black;
		((ARControl)txtTl1).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTl1).Border.TopColor = Color.Black;
		((ARControl)txtTl1).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTl1).DataField = "tl1";
		((ARControl)txtTl1).Height = 0.25f;
		((ARControl)txtTl1).Left = 0.5f;
		((ARControl)txtTl1).Name = "txtTl1";
		componentResourceManager.ApplyResources(txtTl1, "txtTl1");
		((ARControl)txtTl1).Top = 0f;
		((ARControl)txtTl1).Width = 0.5625f;
		((ARControl)txtSolong1).Border.BottomColor = Color.Black;
		((ARControl)txtSolong1).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSolong1).Border.LeftColor = Color.Black;
		((ARControl)txtSolong1).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSolong1).Border.RightColor = Color.Black;
		((ARControl)txtSolong1).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSolong1).Border.TopColor = Color.Black;
		((ARControl)txtSolong1).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSolong1).DataField = "solong1";
		((ARControl)txtSolong1).Height = 0.25f;
		((ARControl)txtSolong1).Left = 1.0625f;
		((ARControl)txtSolong1).Name = "txtSolong1";
		componentResourceManager.ApplyResources(txtSolong1, "txtSolong1");
		((ARControl)txtSolong1).Top = 0f;
		((ARControl)txtSolong1).Width = 0.4375f;
		((ARControl)txtSocon1).Border.BottomColor = Color.Black;
		((ARControl)txtSocon1).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSocon1).Border.LeftColor = Color.Black;
		((ARControl)txtSocon1).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSocon1).Border.RightColor = Color.Black;
		((ARControl)txtSocon1).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSocon1).Border.TopColor = Color.Black;
		((ARControl)txtSocon1).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSocon1).DataField = "socon1";
		((ARControl)txtSocon1).Height = 0.25f;
		((ARControl)txtSocon1).Left = 1.5f;
		((ARControl)txtSocon1).Name = "txtSocon1";
		componentResourceManager.ApplyResources(txtSocon1, "txtSocon1");
		((ARControl)txtSocon1).Top = 0f;
		((ARControl)txtSocon1).Width = 0.4375f;
		((ARControl)txtSTT1).Border.BottomColor = Color.Black;
		((ARControl)txtSTT1).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSTT1).Border.LeftColor = Color.Black;
		((ARControl)txtSTT1).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSTT1).Border.RightColor = Color.Black;
		((ARControl)txtSTT1).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSTT1).Border.TopColor = Color.Black;
		((ARControl)txtSTT1).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSTT1).DataField = "macan1";
		((ARControl)txtSTT1).Height = 0.25f;
		((ARControl)txtSTT1).Left = 0.0625f;
		((ARControl)txtSTT1).Name = "txtSTT1";
		componentResourceManager.ApplyResources(txtSTT1, "txtSTT1");
		((ARControl)txtSTT1).Top = 0f;
		((ARControl)txtSTT1).Width = 0.4375f;
		((ARControl)txtTl2).Border.BottomColor = Color.Black;
		((ARControl)txtTl2).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtTl2).Border.LeftColor = Color.Black;
		((ARControl)txtTl2).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtTl2).Border.RightColor = Color.Black;
		((ARControl)txtTl2).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTl2).Border.TopColor = Color.Black;
		((ARControl)txtTl2).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTl2).DataField = "tl2";
		((ARControl)txtTl2).Height = 0.25f;
		((ARControl)txtTl2).Left = 2.375f;
		((ARControl)txtTl2).Name = "txtTl2";
		componentResourceManager.ApplyResources(txtTl2, "txtTl2");
		((ARControl)txtTl2).Top = 0f;
		((ARControl)txtTl2).Width = 0.5625f;
		((ARControl)txtSolong2).Border.BottomColor = Color.Black;
		((ARControl)txtSolong2).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSolong2).Border.LeftColor = Color.Black;
		((ARControl)txtSolong2).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSolong2).Border.RightColor = Color.Black;
		((ARControl)txtSolong2).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSolong2).Border.TopColor = Color.Black;
		((ARControl)txtSolong2).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSolong2).DataField = "solong2";
		((ARControl)txtSolong2).Height = 0.25f;
		((ARControl)txtSolong2).Left = 2.9375f;
		((ARControl)txtSolong2).Name = "txtSolong2";
		componentResourceManager.ApplyResources(txtSolong2, "txtSolong2");
		((ARControl)txtSolong2).Top = 0f;
		((ARControl)txtSolong2).Width = 0.4375f;
		((ARControl)txtSocon2).Border.BottomColor = Color.Black;
		((ARControl)txtSocon2).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSocon2).Border.LeftColor = Color.Black;
		((ARControl)txtSocon2).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSocon2).Border.RightColor = Color.Black;
		((ARControl)txtSocon2).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSocon2).Border.TopColor = Color.Black;
		((ARControl)txtSocon2).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSocon2).DataField = "socon2";
		((ARControl)txtSocon2).Height = 0.25f;
		((ARControl)txtSocon2).Left = 3.375f;
		((ARControl)txtSocon2).Name = "txtSocon2";
		componentResourceManager.ApplyResources(txtSocon2, "txtSocon2");
		((ARControl)txtSocon2).Top = 0f;
		((ARControl)txtSocon2).Width = 0.4375f;
		((ARControl)txtSTT2).Border.BottomColor = Color.Black;
		((ARControl)txtSTT2).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSTT2).Border.LeftColor = Color.Black;
		((ARControl)txtSTT2).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSTT2).Border.RightColor = Color.Black;
		((ARControl)txtSTT2).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSTT2).Border.TopColor = Color.Black;
		((ARControl)txtSTT2).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSTT2).DataField = "macan2";
		((ARControl)txtSTT2).Height = 0.25f;
		((ARControl)txtSTT2).Left = 1.9375f;
		((ARControl)txtSTT2).Name = "txtSTT2";
		componentResourceManager.ApplyResources(txtSTT2, "txtSTT2");
		((ARControl)txtSTT2).Top = 0f;
		((ARControl)txtSTT2).Width = 0.4375f;
		((ARControl)txtTL3).Border.BottomColor = Color.Black;
		((ARControl)txtTL3).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtTL3).Border.LeftColor = Color.Black;
		((ARControl)txtTL3).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtTL3).Border.RightColor = Color.Black;
		((ARControl)txtTL3).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTL3).Border.TopColor = Color.Black;
		((ARControl)txtTL3).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTL3).DataField = "tl3";
		((ARControl)txtTL3).Height = 0.25f;
		((ARControl)txtTL3).Left = 4.25f;
		((ARControl)txtTL3).Name = "txtTL3";
		componentResourceManager.ApplyResources(txtTL3, "txtTL3");
		((ARControl)txtTL3).Top = 0f;
		((ARControl)txtTL3).Width = 0.5625f;
		((ARControl)txtSolong3).Border.BottomColor = Color.Black;
		((ARControl)txtSolong3).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSolong3).Border.LeftColor = Color.Black;
		((ARControl)txtSolong3).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSolong3).Border.RightColor = Color.Black;
		((ARControl)txtSolong3).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSolong3).Border.TopColor = Color.Black;
		((ARControl)txtSolong3).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSolong3).DataField = "solong3";
		((ARControl)txtSolong3).Height = 0.25f;
		((ARControl)txtSolong3).Left = 4.8125f;
		((ARControl)txtSolong3).Name = "txtSolong3";
		componentResourceManager.ApplyResources(txtSolong3, "txtSolong3");
		((ARControl)txtSolong3).Top = 0f;
		((ARControl)txtSolong3).Width = 0.4375f;
		((ARControl)txtSocon3).Border.BottomColor = Color.Black;
		((ARControl)txtSocon3).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSocon3).Border.LeftColor = Color.Black;
		((ARControl)txtSocon3).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSocon3).Border.RightColor = Color.Black;
		((ARControl)txtSocon3).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSocon3).Border.TopColor = Color.Black;
		((ARControl)txtSocon3).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSocon3).DataField = "socon3";
		((ARControl)txtSocon3).Height = 0.25f;
		((ARControl)txtSocon3).Left = 5.25f;
		((ARControl)txtSocon3).Name = "txtSocon3";
		componentResourceManager.ApplyResources(txtSocon3, "txtSocon3");
		((ARControl)txtSocon3).Top = 0f;
		((ARControl)txtSocon3).Width = 0.4375f;
		((ARControl)txtSTT3).Border.BottomColor = Color.Black;
		((ARControl)txtSTT3).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSTT3).Border.LeftColor = Color.Black;
		((ARControl)txtSTT3).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSTT3).Border.RightColor = Color.Black;
		((ARControl)txtSTT3).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSTT3).Border.TopColor = Color.Black;
		((ARControl)txtSTT3).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSTT3).DataField = "macan3";
		((ARControl)txtSTT3).Height = 0.25f;
		((ARControl)txtSTT3).Left = 3.8125f;
		((ARControl)txtSTT3).Name = "txtSTT3";
		componentResourceManager.ApplyResources(txtSTT3, "txtSTT3");
		((ARControl)txtSTT3).Top = 0f;
		((ARControl)txtSTT3).Width = 0.4375f;
		((ARControl)txtTL4).Border.BottomColor = Color.Black;
		((ARControl)txtTL4).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtTL4).Border.LeftColor = Color.Black;
		((ARControl)txtTL4).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtTL4).Border.RightColor = Color.Black;
		((ARControl)txtTL4).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTL4).Border.TopColor = Color.Black;
		((ARControl)txtTL4).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTL4).DataField = "tl4";
		((ARControl)txtTL4).Height = 0.25f;
		((ARControl)txtTL4).Left = 6.125f;
		((ARControl)txtTL4).Name = "txtTL4";
		componentResourceManager.ApplyResources(txtTL4, "txtTL4");
		((ARControl)txtTL4).Top = 0f;
		((ARControl)txtTL4).Width = 0.5625f;
		((ARControl)txtSolong4).Border.BottomColor = Color.Black;
		((ARControl)txtSolong4).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSolong4).Border.LeftColor = Color.Black;
		((ARControl)txtSolong4).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSolong4).Border.RightColor = Color.Black;
		((ARControl)txtSolong4).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSolong4).Border.TopColor = Color.Black;
		((ARControl)txtSolong4).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSolong4).DataField = "solong4";
		((ARControl)txtSolong4).Height = 0.25f;
		((ARControl)txtSolong4).Left = 6.6875f;
		((ARControl)txtSolong4).Name = "txtSolong4";
		componentResourceManager.ApplyResources(txtSolong4, "txtSolong4");
		((ARControl)txtSolong4).Top = 0f;
		((ARControl)txtSolong4).Width = 0.4375f;
		((ARControl)txtSocon4).Border.BottomColor = Color.Black;
		((ARControl)txtSocon4).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSocon4).Border.LeftColor = Color.Black;
		((ARControl)txtSocon4).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSocon4).Border.RightColor = Color.Black;
		((ARControl)txtSocon4).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)txtSocon4).Border.TopColor = Color.Black;
		((ARControl)txtSocon4).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSocon4).DataField = "socon4";
		((ARControl)txtSocon4).Height = 0.25f;
		((ARControl)txtSocon4).Left = 7.125f;
		((ARControl)txtSocon4).Name = "txtSocon4";
		componentResourceManager.ApplyResources(txtSocon4, "txtSocon4");
		((ARControl)txtSocon4).Top = 0f;
		((ARControl)txtSocon4).Width = 0.4375f;
		((ARControl)txtSTT4).Border.BottomColor = Color.Black;
		((ARControl)txtSTT4).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtSTT4).Border.LeftColor = Color.Black;
		((ARControl)txtSTT4).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtSTT4).Border.RightColor = Color.Black;
		((ARControl)txtSTT4).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtSTT4).Border.TopColor = Color.Black;
		((ARControl)txtSTT4).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtSTT4).DataField = "macan4";
		((ARControl)txtSTT4).Height = 0.25f;
		((ARControl)txtSTT4).Left = 5.6875f;
		((ARControl)txtSTT4).Name = "txtSTT4";
		componentResourceManager.ApplyResources(txtSTT4, "txtSTT4");
		((ARControl)txtSTT4).Top = 0f;
		((ARControl)txtSTT4).Width = 0.4375f;
		((ARControl)txtTotalSomacan).Border.BottomColor = Color.Black;
		((ARControl)txtTotalSomacan).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSomacan).Border.LeftColor = Color.Black;
		((ARControl)txtTotalSomacan).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSomacan).Border.RightColor = Color.Black;
		((ARControl)txtTotalSomacan).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSomacan).Border.TopColor = Color.Black;
		((ARControl)txtTotalSomacan).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSomacan).Height = 0.25f;
		((ARControl)txtTotalSomacan).Left = 6.1875f;
		((ARControl)txtTotalSomacan).Name = "txtTotalSomacan";
		componentResourceManager.ApplyResources(txtTotalSomacan, "txtTotalSomacan");
		txtTotalSomacan.Text = null;
		((ARControl)txtTotalSomacan).Top = 0f;
		((ARControl)txtTotalSomacan).Width = 1.3125f;
		((ARControl)txtTotalTLgalong).Border.BottomColor = Color.Black;
		((ARControl)txtTotalTLgalong).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtTotalTLgalong).Border.LeftColor = Color.Black;
		((ARControl)txtTotalTLgalong).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtTotalTLgalong).Border.RightColor = Color.Black;
		((ARControl)txtTotalTLgalong).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTotalTLgalong).Border.TopColor = Color.Black;
		((ARControl)txtTotalTLgalong).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTotalTLgalong).Height = 0.25f;
		((ARControl)txtTotalTLgalong).Left = 2.3125f;
		((ARControl)txtTotalTLgalong).Name = "txtTotalTLgalong";
		componentResourceManager.ApplyResources(txtTotalTLgalong, "txtTotalTLgalong");
		txtTotalTLgalong.Text = null;
		((ARControl)txtTotalTLgalong).Top = 0.5f;
		((ARControl)txtTotalTLgalong).Width = 1.4375f;
		((ARControl)label14).Border.BottomColor = Color.Black;
		((ARControl)label14).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label14).Border.LeftColor = Color.Black;
		((ARControl)label14).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label14).Border.RightColor = Color.Black;
		((ARControl)label14).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label14).Border.TopColor = Color.Black;
		((ARControl)label14).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label14).Height = 0.25f;
		label14.HyperLink = null;
		((ARControl)label14).Left = 3.8125f;
		((ARControl)label14).Name = "label14";
		componentResourceManager.ApplyResources(label14, "label14");
		((ARControl)label14).Top = 0f;
		((ARControl)label14).Width = 2.25f;
		((ARControl)txtTotalSolong).Border.BottomColor = Color.Black;
		((ARControl)txtTotalSolong).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSolong).Border.LeftColor = Color.Black;
		((ARControl)txtTotalSolong).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSolong).Border.RightColor = Color.Black;
		((ARControl)txtTotalSolong).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSolong).Border.TopColor = Color.Black;
		((ARControl)txtTotalSolong).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSolong).Height = 0.25f;
		((ARControl)txtTotalSolong).Left = 2.3125f;
		((ARControl)txtTotalSolong).Name = "txtTotalSolong";
		componentResourceManager.ApplyResources(txtTotalSolong, "txtTotalSolong");
		txtTotalSolong.Text = null;
		((ARControl)txtTotalSolong).Top = 0.75f;
		((ARControl)txtTotalSolong).Width = 1.4375f;
		((ARControl)label45).Border.BottomColor = Color.Black;
		((ARControl)label45).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label45).Border.LeftColor = Color.Black;
		((ARControl)label45).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label45).Border.RightColor = Color.Black;
		((ARControl)label45).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label45).Border.TopColor = Color.Black;
		((ARControl)label45).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label45).Height = 0.25f;
		label45.HyperLink = null;
		((ARControl)label45).Left = 0.0625f;
		((ARControl)label45).Name = "label45";
		componentResourceManager.ApplyResources(label45, "label45");
		((ARControl)label45).Top = 0.75f;
		((ARControl)label45).Width = 2.125f;
		((ARControl)txtTotalTLcabi).Border.BottomColor = Color.Black;
		((ARControl)txtTotalTLcabi).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtTotalTLcabi).Border.LeftColor = Color.Black;
		((ARControl)txtTotalTLcabi).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtTotalTLcabi).Border.RightColor = Color.Black;
		((ARControl)txtTotalTLcabi).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTotalTLcabi).Border.TopColor = Color.Black;
		((ARControl)txtTotalTLcabi).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTotalTLcabi).Height = 0.25f;
		((ARControl)txtTotalTLcabi).Left = 2.3125f;
		((ARControl)txtTotalTLcabi).Name = "txtTotalTLcabi";
		componentResourceManager.ApplyResources(txtTotalTLcabi, "txtTotalTLcabi");
		txtTotalTLcabi.Text = null;
		((ARControl)txtTotalTLcabi).Top = 0f;
		((ARControl)txtTotalTLcabi).Width = 1.4375f;
		((ARControl)label46).Border.BottomColor = Color.Black;
		((ARControl)label46).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label46).Border.LeftColor = Color.Black;
		((ARControl)label46).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label46).Border.RightColor = Color.Black;
		((ARControl)label46).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label46).Border.TopColor = Color.Black;
		((ARControl)label46).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label46).Height = 0.25f;
		label46.HyperLink = null;
		((ARControl)label46).Left = 0.0625f;
		((ARControl)label46).Name = "label46";
		componentResourceManager.ApplyResources(label46, "label46");
		((ARControl)label46).Top = 0f;
		((ARControl)label46).Width = 2.125f;
		((ARControl)txtTotalSLgalong).Border.BottomColor = Color.Black;
		((ARControl)txtTotalSLgalong).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSLgalong).Border.LeftColor = Color.Black;
		((ARControl)txtTotalSLgalong).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSLgalong).Border.RightColor = Color.Black;
		((ARControl)txtTotalSLgalong).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSLgalong).Border.TopColor = Color.Black;
		((ARControl)txtTotalSLgalong).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSLgalong).Height = 0.25f;
		((ARControl)txtTotalSLgalong).Left = 6.1875f;
		((ARControl)txtTotalSLgalong).Name = "txtTotalSLgalong";
		componentResourceManager.ApplyResources(txtTotalSLgalong, "txtTotalSLgalong");
		txtTotalSLgalong.Text = null;
		((ARControl)txtTotalSLgalong).Top = 0.5f;
		((ARControl)txtTotalSLgalong).Width = 1.3125f;
		((ARControl)txtTlBinhquan).Border.BottomColor = Color.Black;
		((ARControl)txtTlBinhquan).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtTlBinhquan).Border.LeftColor = Color.Black;
		((ARControl)txtTlBinhquan).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtTlBinhquan).Border.RightColor = Color.Black;
		((ARControl)txtTlBinhquan).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTlBinhquan).Border.TopColor = Color.Black;
		((ARControl)txtTlBinhquan).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTlBinhquan).Height = 0.25f;
		((ARControl)txtTlBinhquan).Left = 6.1875f;
		((ARControl)txtTlBinhquan).Name = "txtTlBinhquan";
		componentResourceManager.ApplyResources(txtTlBinhquan, "txtTlBinhquan");
		txtTlBinhquan.Text = null;
		((ARControl)txtTlBinhquan).Top = 0.75f;
		((ARControl)txtTlBinhquan).Width = 1.3125f;
		((ARControl)label47).Border.BottomColor = Color.Black;
		((ARControl)label47).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label47).Border.LeftColor = Color.Black;
		((ARControl)label47).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label47).Border.RightColor = Color.Black;
		((ARControl)label47).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label47).Border.TopColor = Color.Black;
		((ARControl)label47).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label47).Height = 0.25f;
		label47.HyperLink = null;
		((ARControl)label47).Left = 3.8125f;
		((ARControl)label47).Name = "label47";
		componentResourceManager.ApplyResources(label47, "label47");
		((ARControl)label47).Top = 0.5f;
		((ARControl)label47).Width = 2.25f;
		((ARControl)label48).Border.BottomColor = Color.Black;
		((ARControl)label48).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label48).Border.LeftColor = Color.Black;
		((ARControl)label48).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label48).Border.RightColor = Color.Black;
		((ARControl)label48).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label48).Border.TopColor = Color.Black;
		((ARControl)label48).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label48).Height = 0.25f;
		label48.HyperLink = null;
		((ARControl)label48).Left = 3.8125f;
		((ARControl)label48).Name = "label48";
		componentResourceManager.ApplyResources(label48, "label48");
		((ARControl)label48).Top = 0.75f;
		((ARControl)label48).Width = 2.25f;
		((ARControl)label49).Border.BottomColor = Color.Black;
		((ARControl)label49).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label49).Border.LeftColor = Color.Black;
		((ARControl)label49).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label49).Border.RightColor = Color.Black;
		((ARControl)label49).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label49).Border.TopColor = Color.Black;
		((ARControl)label49).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label49).Height = 0.25f;
		label49.HyperLink = null;
		((ARControl)label49).Left = 0.0625f;
		((ARControl)label49).Name = "label49";
		componentResourceManager.ApplyResources(label49, "label49");
		((ARControl)label49).Top = 0.5f;
		((ARControl)label49).Width = 2.125f;
		((ARControl)txtTLBi).Border.BottomColor = Color.Black;
		((ARControl)txtTLBi).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtTLBi).Border.LeftColor = Color.Black;
		((ARControl)txtTLBi).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtTLBi).Border.RightColor = Color.Black;
		((ARControl)txtTLBi).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTLBi).Border.TopColor = Color.Black;
		((ARControl)txtTLBi).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTLBi).Height = 0.25f;
		((ARControl)txtTLBi).Left = 2.3125f;
		((ARControl)txtTLBi).Name = "txtTLBi";
		componentResourceManager.ApplyResources(txtTLBi, "txtTLBi");
		txtTLBi.Text = null;
		((ARControl)txtTLBi).Top = 0.25f;
		((ARControl)txtTLBi).Width = 1.4375f;
		((ARControl)label50).Border.BottomColor = Color.Black;
		((ARControl)label50).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label50).Border.LeftColor = Color.Black;
		((ARControl)label50).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label50).Border.RightColor = Color.Black;
		((ARControl)label50).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label50).Border.TopColor = Color.Black;
		((ARControl)label50).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label50).Height = 0.25f;
		label50.HyperLink = null;
		((ARControl)label50).Left = 0.0625f;
		((ARControl)label50).Name = "label50";
		componentResourceManager.ApplyResources(label50, "label50");
		((ARControl)label50).Top = 0.25f;
		((ARControl)label50).Width = 2.125f;
		((ARControl)label51).Border.BottomColor = Color.Black;
		((ARControl)label51).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label51).Border.LeftColor = Color.Black;
		((ARControl)label51).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label51).Border.RightColor = Color.Black;
		((ARControl)label51).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label51).Border.TopColor = Color.Black;
		((ARControl)label51).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label51).Height = 0.25f;
		label51.HyperLink = null;
		((ARControl)label51).Left = 2.1875f;
		((ARControl)label51).Name = "label51";
		componentResourceManager.ApplyResources(label51, "label51");
		((ARControl)label51).Top = 0f;
		((ARControl)label51).Width = 0.125f;
		((ARControl)label52).Border.BottomColor = Color.Black;
		((ARControl)label52).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label52).Border.LeftColor = Color.Black;
		((ARControl)label52).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label52).Border.RightColor = Color.Black;
		((ARControl)label52).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label52).Border.TopColor = Color.Black;
		((ARControl)label52).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label52).Height = 0.25f;
		label52.HyperLink = null;
		((ARControl)label52).Left = 2.1875f;
		((ARControl)label52).Name = "label52";
		componentResourceManager.ApplyResources(label52, "label52");
		((ARControl)label52).Top = 0.75f;
		((ARControl)label52).Width = 0.125f;
		((ARControl)label53).Border.BottomColor = Color.Black;
		((ARControl)label53).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label53).Border.LeftColor = Color.Black;
		((ARControl)label53).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label53).Border.RightColor = Color.Black;
		((ARControl)label53).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label53).Border.TopColor = Color.Black;
		((ARControl)label53).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label53).Height = 0.25f;
		label53.HyperLink = null;
		((ARControl)label53).Left = 2.1875f;
		((ARControl)label53).Name = "label53";
		componentResourceManager.ApplyResources(label53, "label53");
		((ARControl)label53).Top = 0.25f;
		((ARControl)label53).Width = 0.125f;
		((ARControl)label54).Border.BottomColor = Color.Black;
		((ARControl)label54).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label54).Border.LeftColor = Color.Black;
		((ARControl)label54).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label54).Border.RightColor = Color.Black;
		((ARControl)label54).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label54).Border.TopColor = Color.Black;
		((ARControl)label54).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label54).Height = 0.25f;
		label54.HyperLink = null;
		((ARControl)label54).Left = 2.1875f;
		((ARControl)label54).Name = "label54";
		componentResourceManager.ApplyResources(label54, "label54");
		((ARControl)label54).Top = 0.5f;
		((ARControl)label54).Width = 0.125f;
		((ARControl)label55).Border.BottomColor = Color.Black;
		((ARControl)label55).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label55).Border.LeftColor = Color.Black;
		((ARControl)label55).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label55).Border.RightColor = Color.Black;
		((ARControl)label55).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label55).Border.TopColor = Color.Black;
		((ARControl)label55).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label55).Height = 0.25f;
		label55.HyperLink = null;
		((ARControl)label55).Left = 6.0625f;
		((ARControl)label55).Name = "label55";
		componentResourceManager.ApplyResources(label55, "label55");
		((ARControl)label55).Top = 0f;
		((ARControl)label55).Width = 0.125f;
		((ARControl)label57).Border.BottomColor = Color.Black;
		((ARControl)label57).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label57).Border.LeftColor = Color.Black;
		((ARControl)label57).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label57).Border.RightColor = Color.Black;
		((ARControl)label57).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label57).Border.TopColor = Color.Black;
		((ARControl)label57).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label57).Height = 0.25f;
		label57.HyperLink = null;
		((ARControl)label57).Left = 6.0625f;
		((ARControl)label57).Name = "label57";
		componentResourceManager.ApplyResources(label57, "label57");
		((ARControl)label57).Top = 0.5f;
		((ARControl)label57).Width = 0.125f;
		((ARControl)label58).Border.BottomColor = Color.Black;
		((ARControl)label58).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label58).Border.LeftColor = Color.Black;
		((ARControl)label58).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label58).Border.RightColor = Color.Black;
		((ARControl)label58).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label58).Border.TopColor = Color.Black;
		((ARControl)label58).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label58).Height = 0.25f;
		label58.HyperLink = null;
		((ARControl)label58).Left = 6.0625f;
		((ARControl)label58).Name = "label58";
		componentResourceManager.ApplyResources(label58, "label58");
		((ARControl)label58).Top = 0.75f;
		((ARControl)label58).Width = 0.125f;
		((ARControl)txtLoaican).Border.BottomColor = Color.Black;
		((ARControl)txtLoaican).Border.BottomStyle = (BorderLineStyle)1;
		((ARControl)txtLoaican).Border.LeftColor = Color.Black;
		((ARControl)txtLoaican).Border.LeftStyle = (BorderLineStyle)1;
		((ARControl)txtLoaican).Border.RightColor = Color.Black;
		((ARControl)txtLoaican).Border.RightStyle = (BorderLineStyle)1;
		((ARControl)txtLoaican).Border.TopColor = Color.Black;
		((ARControl)txtLoaican).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtLoaican).DataField = "loaican";
		((ARControl)txtLoaican).Height = 0.25f;
		((ARControl)txtLoaican).Left = 0.0625f;
		((ARControl)txtLoaican).Name = "txtLoaican";
		componentResourceManager.ApplyResources(txtLoaican, "txtLoaican");
		((ARControl)txtLoaican).Top = 0f;
		((ARControl)txtLoaican).Width = 7.5f;
		((ARControl)txtTotalSomacanBi).Border.BottomColor = Color.Black;
		((ARControl)txtTotalSomacanBi).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSomacanBi).Border.LeftColor = Color.Black;
		((ARControl)txtTotalSomacanBi).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSomacanBi).Border.RightColor = Color.Black;
		((ARControl)txtTotalSomacanBi).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSomacanBi).Border.TopColor = Color.Black;
		((ARControl)txtTotalSomacanBi).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)txtTotalSomacanBi).Height = 0.25f;
		((ARControl)txtTotalSomacanBi).Left = 6.1875f;
		((ARControl)txtTotalSomacanBi).Name = "txtTotalSomacanBi";
		componentResourceManager.ApplyResources(txtTotalSomacanBi, "txtTotalSomacanBi");
		txtTotalSomacanBi.Text = null;
		((ARControl)txtTotalSomacanBi).Top = 0.25f;
		((ARControl)txtTotalSomacanBi).Width = 1.3125f;
		((ARControl)label10).Border.BottomColor = Color.Black;
		((ARControl)label10).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label10).Border.LeftColor = Color.Black;
		((ARControl)label10).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label10).Border.RightColor = Color.Black;
		((ARControl)label10).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label10).Border.TopColor = Color.Black;
		((ARControl)label10).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label10).Height = 0.25f;
		label10.HyperLink = null;
		((ARControl)label10).Left = 3.8125f;
		((ARControl)label10).Name = "label10";
		componentResourceManager.ApplyResources(label10, "label10");
		((ARControl)label10).Top = 0.25f;
		((ARControl)label10).Width = 2.25f;
		((ARControl)label16).Border.BottomColor = Color.Black;
		((ARControl)label16).Border.BottomStyle = (BorderLineStyle)0;
		((ARControl)label16).Border.LeftColor = Color.Black;
		((ARControl)label16).Border.LeftStyle = (BorderLineStyle)0;
		((ARControl)label16).Border.RightColor = Color.Black;
		((ARControl)label16).Border.RightStyle = (BorderLineStyle)0;
		((ARControl)label16).Border.TopColor = Color.Black;
		((ARControl)label16).Border.TopStyle = (BorderLineStyle)0;
		((ARControl)label16).Height = 0.25f;
		label16.HyperLink = null;
		((ARControl)label16).Left = 6.0625f;
		((ARControl)label16).Name = "label16";
		componentResourceManager.ApplyResources(label16, "label16");
		((ARControl)label16).Top = 0.25f;
		((ARControl)label16).Width = 0.125f;
		detail.ColumnSpacing = 0f;
		((Section)detail).Controls.AddRange((ARControl[])(object)new ARControl[16]
		{
			(ARControl)txtTl1,
			(ARControl)txtSolong1,
			(ARControl)txtSocon1,
			(ARControl)txtSTT1,
			(ARControl)txtTl2,
			(ARControl)txtSolong2,
			(ARControl)txtSocon2,
			(ARControl)txtSTT2,
			(ARControl)txtTL3,
			(ARControl)txtSolong3,
			(ARControl)txtSocon3,
			(ARControl)txtSTT3,
			(ARControl)txtTL4,
			(ARControl)txtSolong4,
			(ARControl)txtSocon4,
			(ARControl)txtSTT4
		});
		((Section)detail).Height = 0.2395833f;
		((Section)detail).Name = "detail";
		((Section)detail).Format += detail_Format;
		((Section)reportHeader1).Controls.AddRange((ARControl[])(object)new ARControl[16]
		{
			(ARControl)txtCty,
			(ARControl)txtDiachi,
			(ARControl)txtDienthoai,
			(ARControl)textBox4,
			(ARControl)label1,
			(ARControl)txtSophieu,
			(ARControl)label3,
			(ARControl)label4,
			(ARControl)label5,
			(ARControl)label11,
			(ARControl)label12,
			(ARControl)label13,
			(ARControl)txtSoxe,
			(ARControl)txtKhachhang,
			(ARControl)txtGhichu,
			(ARControl)label26
		});
		((Section)reportHeader1).Height = 1.708333f;
		((Section)reportHeader1).Name = "reportHeader1";
		((Section)reportHeader1).Format += reportHeader1_Format;
		((Section)reportFooter1).Controls.AddRange((ARControl[])(object)new ARControl[26]
		{
			(ARControl)txtKyten,
			(ARControl)txtNgaycan,
			(ARControl)txtTotalSomacan,
			(ARControl)txtTotalTLgalong,
			(ARControl)label14,
			(ARControl)txtTotalSolong,
			(ARControl)label45,
			(ARControl)txtTotalTLcabi,
			(ARControl)label46,
			(ARControl)txtTotalSLgalong,
			(ARControl)txtTlBinhquan,
			(ARControl)label47,
			(ARControl)label48,
			(ARControl)label49,
			(ARControl)txtTLBi,
			(ARControl)label50,
			(ARControl)label51,
			(ARControl)label52,
			(ARControl)label53,
			(ARControl)label54,
			(ARControl)label55,
			(ARControl)label57,
			(ARControl)label58,
			(ARControl)txtTotalSomacanBi,
			(ARControl)label10,
			(ARControl)label16
		});
		((Section)reportFooter1).Height = 1.604167f;
		((Section)reportFooter1).Name = "reportFooter1";
		((Section)reportFooter1).Format += reportFooter1_Format;
		((Section)pageHeader1).Controls.AddRange((ARControl[])(object)new ARControl[16]
		{
			(ARControl)label2,
			(ARControl)label6,
			(ARControl)label7,
			(ARControl)label8,
			(ARControl)label9,
			(ARControl)label15,
			(ARControl)label17,
			(ARControl)label18,
			(ARControl)label36,
			(ARControl)label37,
			(ARControl)label38,
			(ARControl)label39,
			(ARControl)label40,
			(ARControl)label41,
			(ARControl)label42,
			(ARControl)label43
		});
		((Section)pageHeader1).Height = 0.3333333f;
		((Section)pageHeader1).Name = "pageHeader1";
		((Section)pageFooter1).Height = 0f;
		((Section)pageFooter1).Name = "pageFooter1";
		((Section)groupHeader1).Controls.AddRange((ARControl[])(object)new ARControl[1] { (ARControl)txtLoaican });
		groupHeader1.DataField = "loaican";
		((Section)groupHeader1).Height = 0.1770833f;
		((Section)groupHeader1).Name = "groupHeader1";
		((Section)groupHeader1).Format += groupHeader1_Format;
		((Section)groupFooter1).Height = 0.01041667f;
		((Section)groupFooter1).Name = "groupFooter1";
		((ActiveReport3)this).MasterReport = false;
		val.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=sys/data.mdb;Jet OLEDB:Database Password=!@#^%$Nguyen1Van2A3;";
		val.SQL = "select * from phieucan";
		((ActiveReport3)this).DataSource = val;
		((ActiveReport3)this).PageSettings.DefaultPaperSize = false;
		((ActiveReport3)this).PageSettings.Margins.Bottom = 0.2f;
		((ActiveReport3)this).PageSettings.Margins.Left = 0.3f;
		((ActiveReport3)this).PageSettings.Margins.Right = 0.2f;
		((ActiveReport3)this).PageSettings.Margins.Top = 0.2f;
		((ActiveReport3)this).PageSettings.PaperHeight = 11.69291f;
		((ActiveReport3)this).PageSettings.PaperKind = (PaperKind)9;
		((ActiveReport3)this).PageSettings.PaperWidth = 8.267716f;
		((ActiveReport3)this).PrintWidth = 7.65625f;
		((ActiveReport3)this).Sections.Add((Section)(object)reportHeader1);
		((ActiveReport3)this).Sections.Add((Section)(object)pageHeader1);
		((ActiveReport3)this).Sections.Add((Section)(object)groupHeader1);
		((ActiveReport3)this).Sections.Add((Section)(object)detail);
		((ActiveReport3)this).Sections.Add((Section)(object)groupFooter1);
		((ActiveReport3)this).Sections.Add((Section)(object)pageFooter1);
		((ActiveReport3)this).Sections.Add((Section)(object)reportFooter1);
		((ActiveReport3)this).StyleSheet.Add(new StyleSheetRule("font-family: Arial; font-style: normal; text-decoration: none; font-weight: normal; font-size: 10pt; color: Black; ", "Normal"));
		((ActiveReport3)this).StyleSheet.Add(new StyleSheetRule("font-size: 16pt; font-weight: bold; ", "Heading1", "Normal"));
		((ActiveReport3)this).StyleSheet.Add(new StyleSheetRule("font-family: Times New Roman; font-size: 14pt; font-weight: bold; font-style: italic; ", "Heading2", "Normal"));
		((ActiveReport3)this).StyleSheet.Add(new StyleSheetRule("font-size: 13pt; font-weight: bold; ", "Heading3", "Normal"));
		((ActiveReport3)this).ReportStart += rptInphieu2_ReportStart;
		((ISupportInitialize)txtCty).EndInit();
		((ISupportInitialize)txtDiachi).EndInit();
		((ISupportInitialize)txtDienthoai).EndInit();
		((ISupportInitialize)textBox4).EndInit();
		((ISupportInitialize)label1).EndInit();
		((ISupportInitialize)txtSophieu).EndInit();
		((ISupportInitialize)label3).EndInit();
		((ISupportInitialize)label4).EndInit();
		((ISupportInitialize)label5).EndInit();
		((ISupportInitialize)label11).EndInit();
		((ISupportInitialize)label12).EndInit();
		((ISupportInitialize)label13).EndInit();
		((ISupportInitialize)txtSoxe).EndInit();
		((ISupportInitialize)txtKhachhang).EndInit();
		((ISupportInitialize)txtGhichu).EndInit();
		((ISupportInitialize)label26).EndInit();
		((ISupportInitialize)txtNgaycan).EndInit();
		((ISupportInitialize)txtKyten).EndInit();
		((ISupportInitialize)label2).EndInit();
		((ISupportInitialize)label6).EndInit();
		((ISupportInitialize)label7).EndInit();
		((ISupportInitialize)label8).EndInit();
		((ISupportInitialize)label9).EndInit();
		((ISupportInitialize)label15).EndInit();
		((ISupportInitialize)label17).EndInit();
		((ISupportInitialize)label18).EndInit();
		((ISupportInitialize)label36).EndInit();
		((ISupportInitialize)label37).EndInit();
		((ISupportInitialize)label38).EndInit();
		((ISupportInitialize)label39).EndInit();
		((ISupportInitialize)label40).EndInit();
		((ISupportInitialize)label41).EndInit();
		((ISupportInitialize)label42).EndInit();
		((ISupportInitialize)label43).EndInit();
		((ISupportInitialize)txtTl1).EndInit();
		((ISupportInitialize)txtSolong1).EndInit();
		((ISupportInitialize)txtSocon1).EndInit();
		((ISupportInitialize)txtSTT1).EndInit();
		((ISupportInitialize)txtTl2).EndInit();
		((ISupportInitialize)txtSolong2).EndInit();
		((ISupportInitialize)txtSocon2).EndInit();
		((ISupportInitialize)txtSTT2).EndInit();
		((ISupportInitialize)txtTL3).EndInit();
		((ISupportInitialize)txtSolong3).EndInit();
		((ISupportInitialize)txtSocon3).EndInit();
		((ISupportInitialize)txtSTT3).EndInit();
		((ISupportInitialize)txtTL4).EndInit();
		((ISupportInitialize)txtSolong4).EndInit();
		((ISupportInitialize)txtSocon4).EndInit();
		((ISupportInitialize)txtSTT4).EndInit();
		((ISupportInitialize)txtTotalSomacan).EndInit();
		((ISupportInitialize)txtTotalTLgalong).EndInit();
		((ISupportInitialize)label14).EndInit();
		((ISupportInitialize)txtTotalSolong).EndInit();
		((ISupportInitialize)label45).EndInit();
		((ISupportInitialize)txtTotalTLcabi).EndInit();
		((ISupportInitialize)label46).EndInit();
		((ISupportInitialize)txtTotalSLgalong).EndInit();
		((ISupportInitialize)txtTlBinhquan).EndInit();
		((ISupportInitialize)label47).EndInit();
		((ISupportInitialize)label48).EndInit();
		((ISupportInitialize)label49).EndInit();
		((ISupportInitialize)txtTLBi).EndInit();
		((ISupportInitialize)label50).EndInit();
		((ISupportInitialize)label51).EndInit();
		((ISupportInitialize)label52).EndInit();
		((ISupportInitialize)label53).EndInit();
		((ISupportInitialize)label54).EndInit();
		((ISupportInitialize)label55).EndInit();
		((ISupportInitialize)label57).EndInit();
		((ISupportInitialize)label58).EndInit();
		((ISupportInitialize)txtLoaican).EndInit();
		((ISupportInitialize)txtTotalSomacanBi).EndInit();
		((ISupportInitialize)label10).EndInit();
		((ISupportInitialize)label16).EndInit();
		((ISupportInitialize)this).EndInit();
	}

	public rptInphieu2()
	{
		InitializeComponent();
		txtDiachi.Text = DBUtil.getValueNoParam("select diachi from giaotiep");
		txtDienthoai.Text = DBUtil.getValueNoParam("select dienthoai from giaotiep");
		txtKyten.Text = DBUtil.getValueNoParam("select nguoican from giaotiep");
	}

	private void detail_Format(object sender, EventArgs e)
	{
		try
		{
			if (double.Parse(txtTl1.Text) == 0.0)
			{
				txtSTT1.Text = "";
				txtTl1.Text = "";
				txtSolong1.Text = "";
				txtSocon1.Text = "";
			}
			if (double.Parse(txtTl2.Text) == 0.0)
			{
				txtSTT2.Text = "";
				txtTl2.Text = "";
				txtSolong2.Text = "";
				txtSocon2.Text = "";
			}
			if (double.Parse(txtTL3.Text) == 0.0)
			{
				txtSTT3.Text = "";
				txtTL3.Text = "";
				txtSolong3.Text = "";
				txtSocon3.Text = "";
			}
			if (double.Parse(txtTL4.Text) == 0.0)
			{
				txtSTT4.Text = "";
				txtTL4.Text = "";
				txtSolong4.Text = "";
				txtSocon4.Text = "";
			}
			string text = "";
			string text2 = "";
			text = ((!(txtSTT1.Text.ToLower() == "tổng")) ? "text-align: center; font-weight: nomal; font-size: 10pt;font-family: Times New Roman; vertical-align: middle;" : "text-align: center; font-weight: bold; font-size: 10pt;font-family: Times New Roman; vertical-align: middle;");
			text2 = "text-align: center; font-weight: bold; font-size: 10pt;font-family: Times New Roman; vertical-align: middle;";
			txtSTT1.Style = text2;
			txtTl1.Style = text;
			txtSocon1.Style = text;
			txtSolong1.Style = text;
			txtSTT2.Style = text2;
			txtTl2.Style = text;
			txtSocon2.Style = text;
			txtSolong2.Style = text;
			txtSTT3.Style = text2;
			txtTL3.Style = text;
			txtSocon3.Style = text;
			txtSolong3.Style = text;
			txtSTT4.Style = text2;
			txtTL4.Style = text;
			txtSocon4.Style = text;
			txtSolong4.Style = text;
		}
		catch
		{
		}
	}

	private void reportHeader1_Format(object sender, EventArgs e)
	{
		try
		{
			string sql = "select tenkh from khachhang k,phieucan p where k.id=p.kh_id and stt=" + rptSophieu;
			txtKhachhang.Text = DBUtil.getValueNoParam(sql);
			sql = "select sophieu from phieucan where stt=" + rptSophieu;
			txtSophieu.Text = DBUtil.getValueNoParam(sql);
			sql = "select soxe from phieucan where stt=" + rptSophieu;
			txtSoxe.Text = DBUtil.getValueNoParam(sql);
			sql = "select ghichu from phieucan where stt=" + rptSophieu;
			txtGhichu.Text = DBUtil.getValueNoParam(sql);
		}
		catch
		{
		}
	}

	private void reportFooter1_Format(object sender, EventArgs e)
	{
		string sql = "select 'Ngày ' & day(ngaycan) & ' tháng ' & month(ngaycan) & ' năm ' & year(ngaycan) as ngaycanhang  from phieucan where stt=" + rptSophieu;
		txtNgaycan.Text = DBUtil.getValueNoParam(sql);
		sql = "select count(id) from chitietphieucan where loaican=0 and phieu_id=" + rptSophieu;
		double num = 0.0;
		try
		{
			num = double.Parse(DBUtil.getValueNoParam(sql));
		}
		catch
		{
		}
		sql = "select count(id) from chitietphieucan where loaican=1 and phieu_id=" + rptSophieu;
		double num2 = 0.0;
		try
		{
			num2 = double.Parse(DBUtil.getValueNoParam(sql));
		}
		catch
		{
		}
		sql = "select sum(trongluong) from chitietphieucan where loaican=1 and phieu_id=" + rptSophieu;
		double num3 = 0.0;
		try
		{
			num3 = double.Parse(DBUtil.getValueNoParam(sql));
		}
		catch
		{
		}
		sql = "select sum(trongluong) as tongsl,sum(solong) as tongsolong,sum(solong*socon1long) as tongsocon ";
		sql = sql + " from chitietphieucan where loaican=0 and phieu_id=" + rptSophieu;
		try
		{
			dataFetcher.setCommand(sql, 1);
			DataSet dataSet = dataFetcher.fetchDataSet();
			foreach (DataRow row in dataSet.Tables[0].Rows)
			{
				totalSL = double.Parse(row["tongsl"].ToString());
				totalSolong = double.Parse(row["tongsolong"].ToString());
				totalSocon = double.Parse(row["tongsocon"].ToString());
			}
		}
		catch
		{
		}
		txtTotalSomacan.Text = num.ToString("#,##0");
		txtTotalSomacanBi.Text = num2.ToString("#,##0");
		txtTotalTLcabi.Text = totalSL.ToString("#,##0.0");
		txtTotalSolong.Text = totalSolong.ToString("#,##0");
		txtTotalSLgalong.Text = totalSocon.ToString("#,##0");
		txtTLBi.Text = num3.ToString("#,##0.0");
		txtTotalTLgalong.Text = (totalSL - num3).ToString("#,##0.0");
		txtTlBinhquan.Text = ((totalSL - num3) / totalSocon).ToString("#,##0.0");
	}

	private void groupHeader1_Format(object sender, EventArgs e)
	{
		try
		{
			if (double.Parse(txtLoaican.Text) == 0.0)
			{
				txtLoaican.Text = "     Cân gà";
			}
			else
			{
				txtLoaican.Text = "     Cân bì lồng";
			}
		}
		catch
		{
		}
	}

	private void rptInphieu2_ReportStart(object sender, EventArgs e)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		try
		{
			((ActiveReport3)this).Document.Name = "Phiếu in";
			myDatasource = new OleDBDataSource();
			myDatasource.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=sys/data.mdb;Jet OLEDB:Database Password=!@#^%$Nguyen1Van2A3;";
			string text = "";
			text = "select * from phieuin where macan1 <> '0' ";
			myDatasource.SQL = text;
			((ActiveReport3)this).DataSource = myDatasource;
		}
		catch
		{
		}
	}
}
