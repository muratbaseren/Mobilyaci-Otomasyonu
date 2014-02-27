#Region "imports"
Imports System
Imports System.Data
Imports System.Data.OleDb
#End Region

Public Class frmUrunler
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MainMnuUrunler As System.Windows.Forms.MainMenu
    Friend WithEvents grpUrunOzellikleri As System.Windows.Forms.GroupBox
    Friend WithEvents txtUrunKodu As System.Windows.Forms.TextBox
    Friend WithEvents txtUrunGenislik As System.Windows.Forms.TextBox
    Friend WithEvents txtUrunUzunluk As System.Windows.Forms.TextBox
    Friend WithEvents txtUrunBrmFiyat As System.Windows.Forms.TextBox
    Friend WithEvents lstUrunEkstralar As System.Windows.Forms.ListBox
    Friend WithEvents btnEkstraEkle As System.Windows.Forms.Button
    Friend WithEvents btnEkstraSil As System.Windows.Forms.Button
    Friend WithEvents picUrunResmi As System.Windows.Forms.PictureBox
    Friend WithEvents grpUrunListesi As System.Windows.Forms.GroupBox
    Friend WithEvents cboUrunTuru As System.Windows.Forms.ComboBox
    Friend WithEvents lstUrunler As System.Windows.Forms.ListBox
    Friend WithEvents btnUrunResim As System.Windows.Forms.Button
    Friend WithEvents btnUrunBrmFiyat As System.Windows.Forms.Button
    Friend WithEvents mnuUrunEkle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUrunDuzenle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUrunSil As System.Windows.Forms.MenuItem
    Friend WithEvents mnuKapat As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUrunIslemleri2 As System.Windows.Forms.MenuItem
    Friend WithEvents txtUrunAdi As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtToplam As System.Windows.Forms.TextBox
    Friend WithEvents btnGenHesapla As System.Windows.Forms.Button
    Friend WithEvents btnUzunHesapla As System.Windows.Forms.Button
    Friend WithEvents btnKaydet As System.Windows.Forms.Button
    Friend WithEvents btnIptal As System.Windows.Forms.Button
    Friend WithEvents lblYTL As System.Windows.Forms.Label
    Friend WithEvents lblToplam As System.Windows.Forms.Label
    Friend WithEvents lblEkstraUrun As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuUrunEkle As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuUrunDuzenle As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuUrunSil As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmUrunler))
        Me.grpUrunOzellikleri = New System.Windows.Forms.GroupBox
        Me.btnIptal = New System.Windows.Forms.Button
        Me.btnKaydet = New System.Windows.Forms.Button
        Me.btnUzunHesapla = New System.Windows.Forms.Button
        Me.btnGenHesapla = New System.Windows.Forms.Button
        Me.lblYTL = New System.Windows.Forms.Label
        Me.txtUrunAdi = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.btnUrunBrmFiyat = New System.Windows.Forms.Button
        Me.btnUrunResim = New System.Windows.Forms.Button
        Me.picUrunResmi = New System.Windows.Forms.PictureBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtToplam = New System.Windows.Forms.TextBox
        Me.lblToplam = New System.Windows.Forms.Label
        Me.btnEkstraSil = New System.Windows.Forms.Button
        Me.btnEkstraEkle = New System.Windows.Forms.Button
        Me.lstUrunEkstralar = New System.Windows.Forms.ListBox
        Me.lblEkstraUrun = New System.Windows.Forms.Label
        Me.txtUrunBrmFiyat = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtUrunUzunluk = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtUrunGenislik = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtUrunKodu = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.grpUrunListesi = New System.Windows.Forms.GroupBox
        Me.lstUrunler = New System.Windows.Forms.ListBox
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.cmnuUrunEkle = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.cmnuUrunDuzenle = New System.Windows.Forms.MenuItem
        Me.cmnuUrunSil = New System.Windows.Forms.MenuItem
        Me.cboUrunTuru = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.MainMnuUrunler = New System.Windows.Forms.MainMenu
        Me.mnuUrunIslemleri2 = New System.Windows.Forms.MenuItem
        Me.mnuUrunEkle = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.mnuUrunDuzenle = New System.Windows.Forms.MenuItem
        Me.mnuUrunSil = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuKapat = New System.Windows.Forms.MenuItem
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.grpUrunOzellikleri.SuspendLayout()
        Me.grpUrunListesi.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpUrunOzellikleri
        '
        Me.grpUrunOzellikleri.Controls.Add(Me.btnIptal)
        Me.grpUrunOzellikleri.Controls.Add(Me.btnKaydet)
        Me.grpUrunOzellikleri.Controls.Add(Me.btnUzunHesapla)
        Me.grpUrunOzellikleri.Controls.Add(Me.btnGenHesapla)
        Me.grpUrunOzellikleri.Controls.Add(Me.lblYTL)
        Me.grpUrunOzellikleri.Controls.Add(Me.txtUrunAdi)
        Me.grpUrunOzellikleri.Controls.Add(Me.Label7)
        Me.grpUrunOzellikleri.Controls.Add(Me.btnUrunBrmFiyat)
        Me.grpUrunOzellikleri.Controls.Add(Me.btnUrunResim)
        Me.grpUrunOzellikleri.Controls.Add(Me.picUrunResmi)
        Me.grpUrunOzellikleri.Controls.Add(Me.GroupBox3)
        Me.grpUrunOzellikleri.Controls.Add(Me.GroupBox2)
        Me.grpUrunOzellikleri.Controls.Add(Me.txtToplam)
        Me.grpUrunOzellikleri.Controls.Add(Me.lblToplam)
        Me.grpUrunOzellikleri.Controls.Add(Me.btnEkstraSil)
        Me.grpUrunOzellikleri.Controls.Add(Me.btnEkstraEkle)
        Me.grpUrunOzellikleri.Controls.Add(Me.lstUrunEkstralar)
        Me.grpUrunOzellikleri.Controls.Add(Me.lblEkstraUrun)
        Me.grpUrunOzellikleri.Controls.Add(Me.txtUrunBrmFiyat)
        Me.grpUrunOzellikleri.Controls.Add(Me.Label6)
        Me.grpUrunOzellikleri.Controls.Add(Me.txtUrunUzunluk)
        Me.grpUrunOzellikleri.Controls.Add(Me.Label5)
        Me.grpUrunOzellikleri.Controls.Add(Me.txtUrunGenislik)
        Me.grpUrunOzellikleri.Controls.Add(Me.Label4)
        Me.grpUrunOzellikleri.Controls.Add(Me.txtUrunKodu)
        Me.grpUrunOzellikleri.Controls.Add(Me.Label3)
        Me.grpUrunOzellikleri.Location = New System.Drawing.Point(240, 16)
        Me.grpUrunOzellikleri.Name = "grpUrunOzellikleri"
        Me.grpUrunOzellikleri.Size = New System.Drawing.Size(352, 472)
        Me.grpUrunOzellikleri.TabIndex = 4
        Me.grpUrunOzellikleri.TabStop = False
        '
        'btnIptal
        '
        Me.btnIptal.BackColor = System.Drawing.Color.DarkCyan
        Me.btnIptal.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnIptal.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnIptal.Location = New System.Drawing.Point(240, 432)
        Me.btnIptal.Name = "btnIptal"
        Me.btnIptal.Size = New System.Drawing.Size(96, 32)
        Me.btnIptal.TabIndex = 28
        Me.btnIptal.Text = "ÝPTAL"
        '
        'btnKaydet
        '
        Me.btnKaydet.BackColor = System.Drawing.Color.DarkCyan
        Me.btnKaydet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnKaydet.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnKaydet.Location = New System.Drawing.Point(136, 432)
        Me.btnKaydet.Name = "btnKaydet"
        Me.btnKaydet.Size = New System.Drawing.Size(96, 32)
        Me.btnKaydet.TabIndex = 27
        Me.btnKaydet.Text = "KAYDET"
        '
        'btnUzunHesapla
        '
        Me.btnUzunHesapla.BackColor = System.Drawing.Color.DarkCyan
        Me.btnUzunHesapla.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUzunHesapla.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnUzunHesapla.Location = New System.Drawing.Point(272, 216)
        Me.btnUzunHesapla.Name = "btnUzunHesapla"
        Me.btnUzunHesapla.Size = New System.Drawing.Size(64, 23)
        Me.btnUzunHesapla.TabIndex = 26
        Me.btnUzunHesapla.Text = "Hesapla"
        '
        'btnGenHesapla
        '
        Me.btnGenHesapla.BackColor = System.Drawing.Color.DarkCyan
        Me.btnGenHesapla.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnGenHesapla.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnGenHesapla.Location = New System.Drawing.Point(104, 216)
        Me.btnGenHesapla.Name = "btnGenHesapla"
        Me.btnGenHesapla.Size = New System.Drawing.Size(64, 23)
        Me.btnGenHesapla.TabIndex = 25
        Me.btnGenHesapla.Text = "Hesapla"
        '
        'lblYTL
        '
        Me.lblYTL.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lblYTL.Location = New System.Drawing.Point(272, 440)
        Me.lblYTL.Name = "lblYTL"
        Me.lblYTL.Size = New System.Drawing.Size(48, 24)
        Me.lblYTL.TabIndex = 24
        Me.lblYTL.Text = "YTL"
        '
        'txtUrunAdi
        '
        Me.txtUrunAdi.AcceptsTab = True
        Me.txtUrunAdi.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunAdi.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunAdi.ForeColor = System.Drawing.Color.Brown
        Me.txtUrunAdi.Location = New System.Drawing.Point(16, 96)
        Me.txtUrunAdi.Name = "txtUrunAdi"
        Me.txtUrunAdi.Size = New System.Drawing.Size(152, 24)
        Me.txtUrunAdi.TabIndex = 23
        Me.txtUrunAdi.Text = ""
        Me.txtUrunAdi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(16, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(152, 23)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Ürün Adý"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnUrunBrmFiyat
        '
        Me.btnUrunBrmFiyat.BackColor = System.Drawing.Color.DarkCyan
        Me.btnUrunBrmFiyat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUrunBrmFiyat.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnUrunBrmFiyat.Location = New System.Drawing.Point(72, 304)
        Me.btnUrunBrmFiyat.Name = "btnUrunBrmFiyat"
        Me.btnUrunBrmFiyat.Size = New System.Drawing.Size(96, 23)
        Me.btnUrunBrmFiyat.TabIndex = 21
        Me.btnUrunBrmFiyat.Text = "Fiyat Deðiþtir"
        '
        'btnUrunResim
        '
        Me.btnUrunResim.BackColor = System.Drawing.Color.DarkCyan
        Me.btnUrunResim.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUrunResim.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnUrunResim.Location = New System.Drawing.Point(240, 144)
        Me.btnUrunResim.Name = "btnUrunResim"
        Me.btnUrunResim.Size = New System.Drawing.Size(96, 24)
        Me.btnUrunResim.TabIndex = 20
        Me.btnUrunResim.Text = "Resim Deðiþtir"
        '
        'picUrunResmi
        '
        Me.picUrunResmi.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picUrunResmi.Location = New System.Drawing.Point(184, 16)
        Me.picUrunResmi.Name = "picUrunResmi"
        Me.picUrunResmi.Size = New System.Drawing.Size(152, 144)
        Me.picUrunResmi.TabIndex = 19
        Me.picUrunResmi.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Location = New System.Drawing.Point(16, 384)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(320, 16)
        Me.GroupBox3.TabIndex = 18
        Me.GroupBox3.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Location = New System.Drawing.Point(16, 168)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(320, 16)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        '
        'txtToplam
        '
        Me.txtToplam.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtToplam.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtToplam.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtToplam.Location = New System.Drawing.Point(80, 432)
        Me.txtToplam.Name = "txtToplam"
        Me.txtToplam.ReadOnly = True
        Me.txtToplam.Size = New System.Drawing.Size(192, 32)
        Me.txtToplam.TabIndex = 16
        Me.txtToplam.Text = ""
        Me.txtToplam.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblToplam
        '
        Me.lblToplam.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.lblToplam.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblToplam.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lblToplam.ForeColor = System.Drawing.Color.Navy
        Me.lblToplam.Location = New System.Drawing.Point(80, 408)
        Me.lblToplam.Name = "lblToplam"
        Me.lblToplam.Size = New System.Drawing.Size(192, 23)
        Me.lblToplam.TabIndex = 15
        Me.lblToplam.Text = "...::: TOPLAM :::..."
        Me.lblToplam.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnEkstraSil
        '
        Me.btnEkstraSil.BackColor = System.Drawing.Color.DarkCyan
        Me.btnEkstraSil.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEkstraSil.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnEkstraSil.Location = New System.Drawing.Point(264, 352)
        Me.btnEkstraSil.Name = "btnEkstraSil"
        Me.btnEkstraSil.TabIndex = 14
        Me.btnEkstraSil.Text = "Sil"
        '
        'btnEkstraEkle
        '
        Me.btnEkstraEkle.BackColor = System.Drawing.Color.DarkCyan
        Me.btnEkstraEkle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEkstraEkle.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnEkstraEkle.Location = New System.Drawing.Point(184, 352)
        Me.btnEkstraEkle.Name = "btnEkstraEkle"
        Me.btnEkstraEkle.TabIndex = 13
        Me.btnEkstraEkle.Text = "Ekle"
        '
        'lstUrunEkstralar
        '
        Me.lstUrunEkstralar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstUrunEkstralar.Location = New System.Drawing.Point(184, 280)
        Me.lstUrunEkstralar.Name = "lstUrunEkstralar"
        Me.lstUrunEkstralar.Size = New System.Drawing.Size(152, 67)
        Me.lstUrunEkstralar.TabIndex = 12
        '
        'lblEkstraUrun
        '
        Me.lblEkstraUrun.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.lblEkstraUrun.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEkstraUrun.ForeColor = System.Drawing.Color.Navy
        Me.lblEkstraUrun.Location = New System.Drawing.Point(184, 256)
        Me.lblEkstraUrun.Name = "lblEkstraUrun"
        Me.lblEkstraUrun.Size = New System.Drawing.Size(152, 23)
        Me.lblEkstraUrun.TabIndex = 11
        Me.lblEkstraUrun.Text = "Ürün Ekstralar"
        Me.lblEkstraUrun.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUrunBrmFiyat
        '
        Me.txtUrunBrmFiyat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunBrmFiyat.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunBrmFiyat.Location = New System.Drawing.Point(16, 280)
        Me.txtUrunBrmFiyat.Name = "txtUrunBrmFiyat"
        Me.txtUrunBrmFiyat.Size = New System.Drawing.Size(152, 24)
        Me.txtUrunBrmFiyat.TabIndex = 8
        Me.txtUrunBrmFiyat.Text = ""
        Me.txtUrunBrmFiyat.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(16, 256)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(152, 23)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Ürün Birim Fiyatý"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUrunUzunluk
        '
        Me.txtUrunUzunluk.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunUzunluk.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunUzunluk.Location = New System.Drawing.Point(184, 216)
        Me.txtUrunUzunluk.Name = "txtUrunUzunluk"
        Me.txtUrunUzunluk.Size = New System.Drawing.Size(88, 24)
        Me.txtUrunUzunluk.TabIndex = 6
        Me.txtUrunUzunluk.Text = ""
        Me.txtUrunUzunluk.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(184, 192)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(152, 23)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Ürün Uzunluk(Cm)"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUrunGenislik
        '
        Me.txtUrunGenislik.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunGenislik.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunGenislik.Location = New System.Drawing.Point(16, 216)
        Me.txtUrunGenislik.Name = "txtUrunGenislik"
        Me.txtUrunGenislik.Size = New System.Drawing.Size(88, 24)
        Me.txtUrunGenislik.TabIndex = 4
        Me.txtUrunGenislik.Text = ""
        Me.txtUrunGenislik.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(16, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(152, 23)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Ürün Geniþlik(Cm)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUrunKodu
        '
        Me.txtUrunKodu.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtUrunKodu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunKodu.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunKodu.ForeColor = System.Drawing.Color.Brown
        Me.txtUrunKodu.Location = New System.Drawing.Point(16, 40)
        Me.txtUrunKodu.Name = "txtUrunKodu"
        Me.txtUrunKodu.ReadOnly = True
        Me.txtUrunKodu.Size = New System.Drawing.Size(152, 24)
        Me.txtUrunKodu.TabIndex = 2
        Me.txtUrunKodu.Text = ""
        Me.txtUrunKodu.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(16, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(152, 23)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Ürün Kodu"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grpUrunListesi
        '
        Me.grpUrunListesi.Controls.Add(Me.lstUrunler)
        Me.grpUrunListesi.Controls.Add(Me.cboUrunTuru)
        Me.grpUrunListesi.Controls.Add(Me.Label2)
        Me.grpUrunListesi.Controls.Add(Me.Label1)
        Me.grpUrunListesi.Location = New System.Drawing.Point(8, 16)
        Me.grpUrunListesi.Name = "grpUrunListesi"
        Me.grpUrunListesi.Size = New System.Drawing.Size(216, 472)
        Me.grpUrunListesi.TabIndex = 5
        Me.grpUrunListesi.TabStop = False
        '
        'lstUrunler
        '
        Me.lstUrunler.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstUrunler.ContextMenu = Me.ContextMenu1
        Me.lstUrunler.Location = New System.Drawing.Point(8, 120)
        Me.lstUrunler.Name = "lstUrunler"
        Me.lstUrunler.Size = New System.Drawing.Size(200, 340)
        Me.lstUrunler.Sorted = True
        Me.lstUrunler.TabIndex = 5
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmnuUrunEkle, Me.MenuItem3, Me.cmnuUrunDuzenle, Me.cmnuUrunSil})
        '
        'cmnuUrunEkle
        '
        Me.cmnuUrunEkle.Index = 0
        Me.cmnuUrunEkle.Text = "Ürün Ekle"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "-"
        '
        'cmnuUrunDuzenle
        '
        Me.cmnuUrunDuzenle.Index = 2
        Me.cmnuUrunDuzenle.Text = "Ürün Düzenle"
        '
        'cmnuUrunSil
        '
        Me.cmnuUrunSil.Index = 3
        Me.cmnuUrunSil.Text = "Ürün Sil"
        '
        'cboUrunTuru
        '
        Me.cboUrunTuru.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboUrunTuru.Items.AddRange(New Object() {"Dolaplar", "Kapýlar", "Koltuklar", "Masalar", "Diðer"})
        Me.cboUrunTuru.Location = New System.Drawing.Point(8, 40)
        Me.cboUrunTuru.Name = "cboUrunTuru"
        Me.cboUrunTuru.Size = New System.Drawing.Size(200, 24)
        Me.cboUrunTuru.TabIndex = 4
        Me.cboUrunTuru.Text = "Ürün Türü Seçiniz..."
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(200, 23)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Ürünler"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(200, 23)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Ürün Kategorisi"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MainMnuUrunler
        '
        Me.MainMnuUrunler.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuUrunIslemleri2})
        '
        'mnuUrunIslemleri2
        '
        Me.mnuUrunIslemleri2.Index = 0
        Me.mnuUrunIslemleri2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuUrunEkle, Me.MenuItem15, Me.mnuUrunDuzenle, Me.mnuUrunSil, Me.MenuItem1, Me.mnuKapat})
        Me.mnuUrunIslemleri2.Text = "&Ürün Ýþlemleri"
        '
        'mnuUrunEkle
        '
        Me.mnuUrunEkle.Index = 0
        Me.mnuUrunEkle.Text = "Ürün Ekl&e"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 1
        Me.MenuItem15.Text = "-"
        '
        'mnuUrunDuzenle
        '
        Me.mnuUrunDuzenle.Index = 2
        Me.mnuUrunDuzenle.Text = "Ürün Dü&zenle"
        '
        'mnuUrunSil
        '
        Me.mnuUrunSil.Index = 3
        Me.mnuUrunSil.Text = "Ürün &Sil"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 4
        Me.MenuItem1.Text = "-"
        '
        'mnuKapat
        '
        Me.mnuKapat.Index = 5
        Me.mnuKapat.Text = "&Kapat"
        '
        'frmUrunler
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(608, 507)
        Me.Controls.Add(Me.grpUrunListesi)
        Me.Controls.Add(Me.grpUrunOzellikleri)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMnuUrunler
        Me.Name = "frmUrunler"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "...::: Ürün Listesi :::..."
        Me.grpUrunOzellikleri.ResumeLayout(False)
        Me.grpUrunListesi.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Yerel Degiskenler"
    Private myConn As OleDbConnection
    Private myCmd As OleDbCommand
    Private myDR As OleDbDataReader
    Private myConnStr As String = _
        "Provider=Microsoft.ace.oledb.12.0;Data Source=" & Application.StartupPath & "\database.mdb;"

    Private ListedeVar As Boolean = False
    Private Urun_Turu_Tablosu As String
    Private ResimKonumu As String
    Private ResimAdi As String
    Private EkstraFiyatSil As Integer
    Private Mode As String = ""
#End Region


#Region "Metotlar"
    'Ürün ekleme, duzenleme iþlemlerinde gizlenecek yada gösterilecek nesnelerin ayarlanmasý..
    Private Sub VisibleAyarlari()
        If Mode = "Add" Then
            Me.btnUrunResim.Text = "Resim Ekle"
            Me.btnUrunResim.Visible = True
            Me.btnKaydet.Visible = True
            Me.btnIptal.Visible = True
            Me.btnGenHesapla.Visible = False
            Me.btnUzunHesapla.Visible = False
            Me.btnUrunBrmFiyat.Visible = False
            Me.btnEkstraEkle.Enabled = False
            Me.btnEkstraSil.Enabled = False
            Me.lstUrunEkstralar.Enabled = False
            Me.lblToplam.Visible = False
            Me.lblYTL.Visible = False
            Me.txtToplam.Visible = False

        ElseIf Mode = "Update" Then
            Me.btnUrunResim.Text = "Resim Deðiþtir"
            Me.btnUrunResim.Visible = True
            Me.btnKaydet.Visible = True
            Me.btnIptal.Visible = True
            Me.btnGenHesapla.Visible = False
            Me.btnUzunHesapla.Visible = False
            Me.btnUrunBrmFiyat.Visible = False
            Me.btnEkstraEkle.Enabled = False
            Me.btnEkstraSil.Enabled = False
            Me.lstUrunEkstralar.Enabled = False
            Me.lblToplam.Visible = False
            Me.lblYTL.Visible = False
            Me.txtToplam.Visible = False

        Else
            Me.btnUrunResim.Visible = False
            Me.btnKaydet.Visible = False
            Me.btnIptal.Visible = False
            Me.btnGenHesapla.Visible = True
            Me.btnUzunHesapla.Visible = True
            Me.btnUrunBrmFiyat.Visible = True
            Me.btnEkstraEkle.Enabled = True
            Me.btnEkstraSil.Enabled = True
            Me.lstUrunEkstralar.Enabled = True
            Me.lblToplam.Visible = True
            Me.lblYTL.Visible = True
            Me.txtToplam.Visible = True
        End If
    End Sub

    'Urun turu combobox ýna el ile deðer girilirse varlýðýnýn kontrolu..
    Private Sub UrunTuruTextKontrol()
        Dim ListedeVar As Boolean = False

        'Eðer Urunturu combo box ýnýn text kýsmýna el ile içeriðinde olmayan bir Urun turu girilirse kontrol edecek, text i sýfýrlýyacak yada database e baðlanacak..
        For i As Integer = 0 To Me.cboUrunTuru.Items.Count - 1
            If Me.cboUrunTuru.Text = Me.cboUrunTuru.Items(i) Then
                ListedeVar = True
            End If
        Next
    End Sub

    'urun turu combobox ýnýn doldurulmasý..
    Private Sub UrunTuruDoldur()
        Try
            Me.cboUrunTuru.Items.Clear()

            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "SELECT Urunler FROM tblUrunler"

            Dim myDR As OleDbDataReader

            myConn.Open()
            myDR = myCmd.ExecuteReader

            Do While myDR.Read
                Me.cboUrunTuru.Items.Add(myDR.Item("Urunler"))
            Loop

        Catch exOle As OleDbException
            MessageBox.Show(exOle.Message, "OLEDB HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "GENEL HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            'myConn.Close()
            'myCmd.Dispose()
            'myConn.Dispose()
        End Try
    End Sub

    'Seçilen urun turune göre urunlerin listelenmesi..
    Private Sub UrunListesiDoldur()
        Dim myTable As String

        Me.lstUrunler.Items.Clear()

        Try
            Dim objListboxNesnesi As ListBoxNesnesi

            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "SELECT * FROM " & Urun_Turu_Tablosu

            Dim myDR As OleDbDataReader

            myConn.Open()
            myDR = myCmd.ExecuteReader

            Do While myDR.Read
                objListboxNesnesi = New ListBoxNesnesi(myDR.Item("MobAdi").ToString, CInt(myDR.Item("MobNo")))

                Me.lstUrunler.Items.Add(objListboxNesnesi)
            Loop

            If Not Me.lstUrunler.Items.Count = 0 Then
                Me.lstUrunler.SelectedIndex = 0
            End If
        Catch exOle As OleDbException
            MessageBox.Show(exOle.Message, "OLEDB HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "GENEL HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            'myConn.Close()
            'myCmd.Dispose()
            'myConn.Dispose()
        End Try
    End Sub

    'Seçilen ürün bilgilerinin textboxlara yazýlmasý..
    Private Sub UrunBilgileri()
        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            Dim objListBoxNesnesi As ListBoxNesnesi
            objListBoxNesnesi = CType(Me.lstUrunler.SelectedItem, ListBoxNesnesi)

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "SELECT * FROM " & Urun_Turu_Tablosu & " WHERE MobNo=@mobno"

            myCmd.Parameters.Add("@mobno", CInt(objListBoxNesnesi.No))

            Dim myDR As OleDbDataReader

            myConn.Open()
            myDR = myCmd.ExecuteReader

            'Urun bilgileri textbox lara dolar..
            Do While myDR.Read
                Me.txtUrunKodu.Text = myDR.Item("MobNo")
                Me.txtUrunAdi.Text = myDR.Item("MobAdi")
                Me.txtUrunGenislik.Text = myDR.Item("MobGen")
                Me.txtUrunUzunluk.Text = myDR.Item("MobUzun")
                Me.txtUrunBrmFiyat.Text = myDR.Item("MobBrmAlanUcret")
                Me.picUrunResmi.Image = Image.FromFile(myDR.Item("MobResim"))
                ResimKonumu = myDR.Item("MobResim")

                'Resim boyutlandýrýlýyor..
                ShrinkImage(Me.picUrunResmi, Me.picUrunResmi, True)
            Loop

            Me.lstUrunEkstralar.Items.Clear()

        Catch exOle As OleDbException
            MessageBox.Show(exOle.Message, "OLEDB HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "GENEL HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            'myConn.Close()
            'myCmd.Dispose()
            'myConn.Dispose()
        End Try
    End Sub

    'Urunun birim fiyatýna göre hesap yapýlmasý..
    Private Sub ToplamHesapla()
        Dim AraToplam As Integer

        AraToplam = CInt(Me.txtUrunBrmFiyat.Text * Me.txtUrunGenislik.Text * Me.txtUrunUzunluk.Text)

        If Not Me.lstUrunEkstralar.Items.Count = Nothing Then
            For i As Integer = 0 To Me.lstUrunEkstralar.Items.Count - 1
                Me.lstUrunEkstralar.SelectedIndex = i
                UrunEkstraFiyatAyirma()
                AraToplam += EkstraFiyatSil
            Next
        End If

        Me.txtToplam.Text = AraToplam
    End Sub

    'Urun ekstrasýna girilen özelliðin fiyatýnýn isminden ayrýlmasý..
    Private Sub UrunEkstraFiyatAyirma()
        Dim AyrilcakKelime As String = Me.lstUrunEkstralar.SelectedItem.ToString
        Dim Deger As String

        For i As Integer = Len(AyrilcakKelime) To 1 Step -1
            Deger = Mid(AyrilcakKelime, i, 1)
            If Deger = "-" Then
                Deger = Mid(AyrilcakKelime, i + 2, (Len(AyrilcakKelime) - i + 1))
                Exit For
            End If
        Next

        EkstraFiyatSil = CInt(Deger)
    End Sub

    'Yeni ürün ekleme..
    Private Sub UrunEkle()
        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "INSERT INTO " & Urun_Turu_Tablosu & _
                " (MobAdi, MobResim, MobBrmAlanUcret, MobGen, MobUzun) " & _
                " VALUES (@mobadi, @mobresim, @mobbrmucret, @mobgen, @mobuzun)"

            myCmd.Parameters.Add("@mobadi", Me.txtUrunAdi.Text)
            myCmd.Parameters.Add("@mobresim", ResimKonumu)
            myCmd.Parameters.Add("@mobbrmucret", Me.txtUrunBrmFiyat.Text)
            myCmd.Parameters.Add("@mobgen", Me.txtUrunGenislik.Text)
            myCmd.Parameters.Add("@mobuzun", Me.txtUrunUzunluk.Text)

            myConn.Open()
            myCmd.ExecuteNonQuery()

        Catch exOle As OleDbException
            MessageBox.Show(exOle.Message, "OLEDB HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "GENEL HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            'myConn.Close()
            'myCmd.Dispose()
            'myConn.Dispose()
        End Try
    End Sub

    'Var olan ürün düzenlenmesi..
    Private Sub UrunDuzenle()
        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            Dim objListBoxNesnesi As ListBoxNesnesi
            objListBoxNesnesi = CType(Me.lstUrunler.SelectedItem, ListBoxNesnesi)

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = _
                "UPDATE " & Urun_Turu_Tablosu & " SET " & _
                    "MobAdi=@mobadi, " & _
                    "MobResim=@mobresim, " & _
                    "MobBrmAlanUcret=@mobbrmucret, " & _
                    "MobGen=@mobgen, " & _
                    "MobUzun=@mobuzun " & _
                "WHERE MobNo=" & objListBoxNesnesi.No

            myCmd.Parameters.Add("@mobadi", Me.txtUrunAdi.Text)
            myCmd.Parameters.Add("@mobresim", ResimKonumu)
            myCmd.Parameters.Add("@mobbrmucret", Me.txtUrunBrmFiyat.Text)
            myCmd.Parameters.Add("@mobgen", CInt(Me.txtUrunGenislik.Text))
            myCmd.Parameters.Add("@mobuzun", CInt(Me.txtUrunUzunluk.Text))

            myConn.Open()
            myCmd.ExecuteNonQuery()

            MessageBox.Show("Ürün güncelleþtirildi..", "Ürün Düzenlendi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

        Catch exOle As OleDbException
            MessageBox.Show(exOle.Message, "OLEDB HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "GENEL HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            'myConn.Close()
            'myCmd.Dispose()
            'myConn.Dispose()
        End Try
    End Sub

    'Var olan urunun silinmesi..
    Private Sub UrunSil()
        Try

            If Me.lstUrunler.SelectedIndex = -1 Then
                MessageBox.Show("Listeden silinecek ürünü seçiniz..", "Ürün Seçmediniz..", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Try
            End If

            Dim objListboxNesnesi As New ListBoxNesnesi
            objListboxNesnesi = CType(Me.lstUrunler.SelectedItem, ListBoxNesnesi)

            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "DELETE FROM " & Urun_Turu_Tablosu & " WHERE MobNo=@mobno"

            myCmd.Parameters.Add("@mobno", CInt(objListboxNesnesi.No))

            myConn.Open()
            myCmd.ExecuteNonQuery()

            MessageBox.Show("Ürün silinmiþtir..", "Ürün Silindi..", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

            TextboxTemizle()

        Catch exOle As OleDbException
            MessageBox.Show(exOle.Message, "OLEDB HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "GENEL HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            'myConn.Close()
            'myCmd.Dispose()
            'myConn.Dispose()
        End Try
    End Sub

    'Yeni ürün kaydý yapýlýrken seçilen resmin sadece isminin alýnmasý..
    Private Sub ResimAdiAl()
        Dim Deger As String

        For i As Integer = Len(ResimKonumu) To 1 Step -1
            Deger = Mid(ResimKonumu, i, 1)
            If Deger = "\" Then
                ResimAdi = Mid(ResimKonumu, i + 1, Len(ResimKonumu) - i)
                Exit For
            End If
        Next
    End Sub

    'Yeni ürün kaydýnda textboxlarýn temizlenmesi..
    Private Sub TextboxTemizle()
        Me.txtUrunAdi.Text = ""
        Me.txtUrunKodu.Text = ""
        Me.txtUrunBrmFiyat.Text = ""
        Me.txtUrunGenislik.Text = ""
        Me.txtUrunUzunluk.Text = ""
        Me.txtToplam.Text = ""
        Me.lstUrunEkstralar.Items.Clear()
        Me.lstUrunler.Items.Clear()
    End Sub

    'Ürün resminin anti_alias ile kuçültülmesi..
    Private Sub ShrinkImage(ByVal from_pic As PictureBox, ByVal to_pic As PictureBox, Optional ByVal anti_alias As Boolean = False)
        ' Kaynak resmini okuyor..
        Dim from_bm As New Bitmap(from_pic.Image)

        ' Boyutlandýrma yapýyor..
        Dim wid As Integer = Me.picUrunResmi.Width
        Dim hgt As Integer = Me.picUrunResmi.Height
        Dim to_bm As New Bitmap(wid, hgt)

        ' Resim kopyalýyor..
        Dim gr As Graphics = Graphics.FromImage(to_bm)
        If anti_alias Then gr.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
        gr.DrawImage(from_bm, 0, 0, wid - 1, hgt - 1)

        ' Sonuç görüntüleniyor..
        to_pic.Image = to_bm
    End Sub

#End Region


#Region "Olaylar"
    Private Sub mnuKapat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKapat.Click
        Me.Close()
    End Sub

    'Urunler formu kapatýlýnca MdiParent da "Ürün Ýþlemleri" menuitem ýnýn aktifhale getirilmesi..
    Private Sub frmUrunler_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Me.MdiParent.Menu.MenuItems.Item(0).MenuItems.Item(0).Enabled = True
    End Sub

    Private Sub frmUrunler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UrunTuruDoldur()
        Mode = ""
        VisibleAyarlari()
    End Sub

    Private Sub cboUrunTuru_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboUrunTuru.TextChanged

        UrunTuruTextKontrol()

        If ListedeVar = False Then
            Me.cboUrunTuru.Text = "Ürün Türü Seçiniz..."
            Exit Sub
        End If
    End Sub

    Private Sub cboUrunTuru_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUrunTuru.SelectedIndexChanged
        Me.lstUrunler.Items.Clear()

        'Seçilen ürün kategorisne göre Database e ekleme yada düzenleme yada silme yapýldýðýnda gidilecek tabloyu hafýzada tutuyoruz..
        Select Case Me.cboUrunTuru.Text
            Case "Dolaplar"
                Urun_Turu_Tablosu = "tblDolaplar"
            Case "Kapýlar"
                Urun_Turu_Tablosu = "tblKapilar"
            Case "Koltuklar"
                Urun_Turu_Tablosu = "tblKoltuklar"
            Case "Masalar"
                Urun_Turu_Tablosu = "tblMasalar"
            Case "Diðer"
                Urun_Turu_Tablosu = "tblDiger"
        End Select

        UrunListesiDoldur()
    End Sub

    Private Sub lstUrunler_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstUrunler.SelectedIndexChanged
        UrunBilgileri()
        ToplamHesapla()
    End Sub

    Private Sub picUrunResmi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picUrunResmi.Click
        Shell("explorer.exe " & ResimKonumu)
    End Sub

    Private Sub btnUrunBrmFiyat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUrunBrmFiyat.Click
        Dim YeniBirimFiyat As Integer

        'Eger urunbirimfiyati degisimi yapilmak için butona basilip sonra cancel yapilirsa errhandler a gidilecek..
        On Error GoTo errhandler

        YeniBirimFiyat = InputBox("Ürünün yeni birim fiyatýný giriniz: " & vbCrLf & "ÖRN: 6 gibi..")
        If Not IsNumeric(YeniBirimFiyat) = True Then
            MessageBox.Show("Lütfen Yeni Birim Fiyatý Bir Sayý Olarak Giriniz..", "HATALI DEÐER", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            Exit Sub
        End If

        Me.txtUrunBrmFiyat.Text = YeniBirimFiyat

errhandler:

        ToplamHesapla()
    End Sub

    Private Sub btnEkstraEkle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEkstraEkle.Click
        Dim UrunEkstra As String
        Dim UrunEkstraFiyati As Integer

        'Eger ekstra ekleme yapilmak için butona basilip sonra cancel yapilirsa errhandler a gidilecek..
        On Error GoTo errhandler

        UrunEkstra = InputBox("Ürüne Eklenecek Özellik Giriniz: " & vbCrLf & "ÖRN:Farklý Boya, Kulp, Cila vb..")
        UrunEkstraFiyati = InputBox("Ürüne Eklenecek Özelliðin Fiyatýný Giriniz: " & vbCrLf & "ÖRN: 8 vb..")

        If Trim(UrunEkstra) = "" AndAlso IsNumeric(UrunEkstraFiyati) = False Then
            MessageBox.Show("Lütfen Özelliði Ekleyiniz yada Fiyatý Bir Rakam Olarak Giriniz..", "HATALI DEÐER", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            Exit Sub
        End If

        Me.lstUrunEkstralar.Items.Add(UrunEkstra & " --- " & UrunEkstraFiyati)

        Me.txtToplam.Text += UrunEkstraFiyati

errhandler:
    End Sub

    Private Sub btnEkstraSil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEkstraSil.Click
        'Ekstra girilen özellik adý ile fiyatýnýn birbirinden ayrýlmasý..
        UrunEkstraFiyatAyirma()

        Me.lstUrunEkstralar.Items.Remove(Me.lstUrunEkstralar.SelectedItem)
        Me.txtToplam.Text -= EkstraFiyatSil
    End Sub

    Private Sub btnGenHesapla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenHesapla.Click, btnUzunHesapla.Click
        ToplamHesapla()
    End Sub

    Private Sub mnuUrunEkle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUrunEkle.Click
        If Me.cboUrunTuru.SelectedIndex = -1 Then
            MessageBox.Show("Lütfen bir 'Ürün Türü' seçiniz..", "Ürün Türü Seçmediniz..", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        Mode = "Add"
        VisibleAyarlari()
        TextboxTemizle()
        Me.picUrunResmi.Image = Nothing
        ResimKonumu = Nothing
    End Sub

    Private Sub mnuUrunDuzenle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUrunDuzenle.Click
        Mode = "Update"
        VisibleAyarlari()
    End Sub

    Private Sub btnKaydet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKaydet.Click
        If Mode = "Add" Then
            'Eðer programýn bulunduðu dizinde Resimler adlý bir klasör yoksa onu yaratacak..
            If Not System.IO.Directory.Exists(Application.StartupPath & "\Resimler") = True Then
                MkDir(Application.StartupPath & "\Resimler")
            End If

            If ResimKonumu = Nothing Then
                'Eger resim secilmezse urune eklenecek olan resim,bu resim resim yok olarak yazý iceren bir resim olmali..
                ResimKonumu = Application.StartupPath & "\Bos.bmp"
            Else
                'Seçilen resim programýn bulunduðu dizinde Resimler adlý klasöre kopyalanacak..
                FileCopy(ResimKonumu, Application.StartupPath & "\Resimler\" & ResimAdi)
                ResimKonumu = Application.StartupPath & "\Resimler\" & ResimAdi
            End If

            UrunEkle()

        ElseIf Mode = "Update" Then
            UrunDuzenle()
        End If

        Mode = ""
        VisibleAyarlari()
        UrunListesiDoldur()
    End Sub

    Private Sub btnIptal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIptal.Click
        Mode = ""
        VisibleAyarlari()
        UrunListesiDoldur()
    End Sub

    Private Sub mnuUrunSil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUrunSil.Click
        DialogResult = MessageBox.Show(Me.lstUrunler.SelectedItem.ToString & " adlý ürünü silmek istediðinizden emin misiniz?", "Ürün Silme", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

        If Not DialogResult = DialogResult.No Then
            UrunSil()
            UrunListesiDoldur()

            If Not Me.lstUrunler.Items.Count = 0 Then
                Me.lstUrunler.SelectedIndex = 0
            End If
        End If
    End Sub


    Private Sub btnUrunResim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUrunResim.Click
        'Eger resim secimi yapilmak için butona basilip sonra cancel yapilirsa errhandler a gidilecek..
        On Error GoTo errhandler

        Me.OpenFileDialog1.Title = "Resim Seç"
        Me.OpenFileDialog1.Filter = "Bitmap Resmi (*.bmp)|*.bmp|JPEG Resmi(*.jpg;*.jpeg)|*.jpg;*.jpeg|GIF Resmi (*.gif)|*.gif"
        Me.OpenFileDialog1.FilterIndex = 2
        Me.OpenFileDialog1.ShowDialog()
        ResimKonumu = Me.OpenFileDialog1.FileName

        ResimAdiAl()

        Me.picUrunResmi.Image = Image.FromFile(ResimKonumu)

        'Resim boyutlandýrýlýyor..
        ShrinkImage(Me.picUrunResmi, Me.picUrunResmi, True)

errhandler:
    End Sub

    'ContextMenu Urun Ekle..
    Private Sub cmnuUrunEkle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuUrunEkle.Click
        mnuUrunEkle_Click(sender, e)
    End Sub

    'ContextMenu Urun Duzenle..
    Private Sub cmnuUrunDuzenle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuUrunDuzenle.Click
        mnuUrunDuzenle_Click(sender, e)
    End Sub

    'ContextMenu Urun Sil..
    Private Sub cmnuUrunSil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuUrunSil.Click
        mnuUrunSil_Click(sender, e)
    End Sub


#End Region

End Class
