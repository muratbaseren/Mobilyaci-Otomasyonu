#Region "imports"
Imports System
Imports System.Data
Imports System.Data.OleDb
#End Region

Public Class frmSiparisler
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
    Friend WithEvents grpUrunListesi As System.Windows.Forms.GroupBox
    Friend WithEvents lstUrunler As System.Windows.Forms.ListBox
    Friend WithEvents cboUrunTuru As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents picUrunResmi As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnEkstraSil As System.Windows.Forms.Button
    Friend WithEvents btnEkstraEkle As System.Windows.Forms.Button
    Friend WithEvents lstUrunEkstralar As System.Windows.Forms.ListBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtUrunBrmFiyat As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtUrunUzunluk As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtUrunGenislik As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtUrunKodu As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnIptal As System.Windows.Forms.Button
    Friend WithEvents btnKaydet As System.Windows.Forms.Button
    Friend WithEvents grpSiparis As System.Windows.Forms.GroupBox
    Friend WithEvents lstSiparisler As System.Windows.Forms.ListBox
    Friend WithEvents grpUrun As System.Windows.Forms.GroupBox
    Friend WithEvents label As System.Windows.Forms.Label
    Friend WithEvents txtSiparisAdi As System.Windows.Forms.TextBox
    Friend WithEvents MainMnuSiparisler As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents lblSiparis As System.Windows.Forms.Label
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSiparisEkle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSiparisDuzenle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSiparisSil As System.Windows.Forms.MenuItem
    Friend WithEvents mnuKapat As System.Windows.Forms.MenuItem
    Friend WithEvents chkTeslim As System.Windows.Forms.CheckBox
    Friend WithEvents txtSiparisID As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents mnuSiparisIslemler2 As System.Windows.Forms.MenuItem
    Friend WithEvents txtToplam As System.Windows.Forms.TextBox
    Friend WithEvents btnGenHesapla As System.Windows.Forms.Button
    Friend WithEvents btnUzunHesapla As System.Windows.Forms.Button
    Friend WithEvents txtUrunTuru As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents btnUrunBrmFiyat As System.Windows.Forms.Button
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents cmnuSiparisEkle As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuSiparisDuzenle As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuSiparisSil As System.Windows.Forms.MenuItem
    Friend WithEvents txtUrunAdi As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblYTL As System.Windows.Forms.Label
    Friend WithEvents txtGizli As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSiparisler))
        Me.grpSiparis = New System.Windows.Forms.GroupBox
        Me.txtSiparisID = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.chkTeslim = New System.Windows.Forms.CheckBox
        Me.btnIptal = New System.Windows.Forms.Button
        Me.btnKaydet = New System.Windows.Forms.Button
        Me.txtSiparisAdi = New System.Windows.Forms.TextBox
        Me.label = New System.Windows.Forms.Label
        Me.lstSiparisler = New System.Windows.Forms.ListBox
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.cmnuSiparisEkle = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.cmnuSiparisDuzenle = New System.Windows.Forms.MenuItem
        Me.cmnuSiparisSil = New System.Windows.Forms.MenuItem
        Me.lblSiparis = New System.Windows.Forms.Label
        Me.grpUrunListesi = New System.Windows.Forms.GroupBox
        Me.lstUrunler = New System.Windows.Forms.ListBox
        Me.cboUrunTuru = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.grpUrun = New System.Windows.Forms.GroupBox
        Me.txtGizli = New System.Windows.Forms.TextBox
        Me.lblYTL = New System.Windows.Forms.Label
        Me.txtUrunAdi = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.btnUrunBrmFiyat = New System.Windows.Forms.Button
        Me.txtUrunTuru = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.btnUzunHesapla = New System.Windows.Forms.Button
        Me.btnGenHesapla = New System.Windows.Forms.Button
        Me.picUrunResmi = New System.Windows.Forms.PictureBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtToplam = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.btnEkstraSil = New System.Windows.Forms.Button
        Me.btnEkstraEkle = New System.Windows.Forms.Button
        Me.lstUrunEkstralar = New System.Windows.Forms.ListBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtUrunBrmFiyat = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtUrunUzunluk = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtUrunGenislik = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtUrunKodu = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.MainMnuSiparisler = New System.Windows.Forms.MainMenu
        Me.mnuSiparisIslemler2 = New System.Windows.Forms.MenuItem
        Me.mnuSiparisEkle = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuSiparisDuzenle = New System.Windows.Forms.MenuItem
        Me.mnuSiparisSil = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.mnuKapat = New System.Windows.Forms.MenuItem
        Me.grpSiparis.SuspendLayout()
        Me.grpUrun.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpSiparis
        '
        Me.grpSiparis.Controls.Add(Me.txtSiparisID)
        Me.grpSiparis.Controls.Add(Me.Label7)
        Me.grpSiparis.Controls.Add(Me.chkTeslim)
        Me.grpSiparis.Controls.Add(Me.btnIptal)
        Me.grpSiparis.Controls.Add(Me.btnKaydet)
        Me.grpSiparis.Controls.Add(Me.txtSiparisAdi)
        Me.grpSiparis.Controls.Add(Me.label)
        Me.grpSiparis.Controls.Add(Me.lstSiparisler)
        Me.grpSiparis.Controls.Add(Me.lblSiparis)
        Me.grpSiparis.Location = New System.Drawing.Point(608, 16)
        Me.grpSiparis.Name = "grpSiparis"
        Me.grpSiparis.Size = New System.Drawing.Size(216, 472)
        Me.grpSiparis.TabIndex = 1
        Me.grpSiparis.TabStop = False
        '
        'txtSiparisID
        '
        Me.txtSiparisID.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtSiparisID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSiparisID.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtSiparisID.ForeColor = System.Drawing.Color.Brown
        Me.txtSiparisID.Location = New System.Drawing.Point(8, 40)
        Me.txtSiparisID.Name = "txtSiparisID"
        Me.txtSiparisID.ReadOnly = True
        Me.txtSiparisID.Size = New System.Drawing.Size(200, 24)
        Me.txtSiparisID.TabIndex = 28
        Me.txtSiparisID.Text = ""
        Me.txtSiparisID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(8, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(200, 23)
        Me.Label7.TabIndex = 27
        Me.Label7.Text = "Sipariþ Kodu"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkTeslim
        '
        Me.chkTeslim.Location = New System.Drawing.Point(8, 352)
        Me.chkTeslim.Name = "chkTeslim"
        Me.chkTeslim.Size = New System.Drawing.Size(200, 24)
        Me.chkTeslim.TabIndex = 26
        Me.chkTeslim.Text = "Teslim Edildi..."
        '
        'btnIptal
        '
        Me.btnIptal.BackColor = System.Drawing.Color.DarkCyan
        Me.btnIptal.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnIptal.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnIptal.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnIptal.Location = New System.Drawing.Point(8, 440)
        Me.btnIptal.Name = "btnIptal"
        Me.btnIptal.Size = New System.Drawing.Size(96, 23)
        Me.btnIptal.TabIndex = 25
        Me.btnIptal.Text = "Ýptal"
        '
        'btnKaydet
        '
        Me.btnKaydet.BackColor = System.Drawing.Color.DarkCyan
        Me.btnKaydet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnKaydet.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnKaydet.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnKaydet.Location = New System.Drawing.Point(112, 440)
        Me.btnKaydet.Name = "btnKaydet"
        Me.btnKaydet.Size = New System.Drawing.Size(96, 23)
        Me.btnKaydet.TabIndex = 24
        Me.btnKaydet.Text = "Kaydet"
        '
        'txtSiparisAdi
        '
        Me.txtSiparisAdi.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSiparisAdi.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSiparisAdi.Location = New System.Drawing.Point(8, 408)
        Me.txtSiparisAdi.Name = "txtSiparisAdi"
        Me.txtSiparisAdi.Size = New System.Drawing.Size(200, 24)
        Me.txtSiparisAdi.TabIndex = 9
        Me.txtSiparisAdi.Text = ""
        Me.txtSiparisAdi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'label
        '
        Me.label.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.label.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.label.ForeColor = System.Drawing.Color.Navy
        Me.label.Location = New System.Drawing.Point(8, 384)
        Me.label.Name = "label"
        Me.label.Size = New System.Drawing.Size(200, 23)
        Me.label.TabIndex = 8
        Me.label.Text = "Sipariþ Veren Adý ve Soyadý"
        Me.label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lstSiparisler
        '
        Me.lstSiparisler.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstSiparisler.ContextMenu = Me.ContextMenu1
        Me.lstSiparisler.Location = New System.Drawing.Point(8, 96)
        Me.lstSiparisler.Name = "lstSiparisler"
        Me.lstSiparisler.Size = New System.Drawing.Size(200, 249)
        Me.lstSiparisler.TabIndex = 7
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmnuSiparisEkle, Me.MenuItem2, Me.cmnuSiparisDuzenle, Me.cmnuSiparisSil})
        '
        'cmnuSiparisEkle
        '
        Me.cmnuSiparisEkle.Index = 0
        Me.cmnuSiparisEkle.Text = "Sipariþ Ekle"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'cmnuSiparisDuzenle
        '
        Me.cmnuSiparisDuzenle.Index = 2
        Me.cmnuSiparisDuzenle.Text = "Sipariþ Düzenle"
        '
        'cmnuSiparisSil
        '
        Me.cmnuSiparisSil.Index = 3
        Me.cmnuSiparisSil.Text = "Sipariþ Sil"
        '
        'lblSiparis
        '
        Me.lblSiparis.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.lblSiparis.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSiparis.ForeColor = System.Drawing.Color.Navy
        Me.lblSiparis.Location = New System.Drawing.Point(8, 72)
        Me.lblSiparis.Name = "lblSiparis"
        Me.lblSiparis.Size = New System.Drawing.Size(200, 23)
        Me.lblSiparis.TabIndex = 6
        Me.lblSiparis.Text = "Sipariþler"
        Me.lblSiparis.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grpUrunListesi
        '
        Me.grpUrunListesi.Location = New System.Drawing.Point(224, 8)
        Me.grpUrunListesi.Name = "grpUrunListesi"
        Me.grpUrunListesi.Size = New System.Drawing.Size(8, 456)
        Me.grpUrunListesi.TabIndex = 7
        Me.grpUrunListesi.TabStop = False
        '
        'lstUrunler
        '
        Me.lstUrunler.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstUrunler.Location = New System.Drawing.Point(8, 120)
        Me.lstUrunler.Name = "lstUrunler"
        Me.lstUrunler.Size = New System.Drawing.Size(200, 340)
        Me.lstUrunler.TabIndex = 5
        '
        'cboUrunTuru
        '
        Me.cboUrunTuru.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.cboUrunTuru.Location = New System.Drawing.Point(8, 40)
        Me.cboUrunTuru.Name = "cboUrunTuru"
        Me.cboUrunTuru.Size = New System.Drawing.Size(200, 25)
        Me.cboUrunTuru.TabIndex = 4
        Me.cboUrunTuru.Text = "Ürün Seçiniz..."
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
        'grpUrun
        '
        Me.grpUrun.Controls.Add(Me.txtGizli)
        Me.grpUrun.Controls.Add(Me.lblYTL)
        Me.grpUrun.Controls.Add(Me.txtUrunAdi)
        Me.grpUrun.Controls.Add(Me.Label11)
        Me.grpUrun.Controls.Add(Me.btnUrunBrmFiyat)
        Me.grpUrun.Controls.Add(Me.txtUrunTuru)
        Me.grpUrun.Controls.Add(Me.Label10)
        Me.grpUrun.Controls.Add(Me.btnUzunHesapla)
        Me.grpUrun.Controls.Add(Me.btnGenHesapla)
        Me.grpUrun.Controls.Add(Me.picUrunResmi)
        Me.grpUrun.Controls.Add(Me.GroupBox3)
        Me.grpUrun.Controls.Add(Me.GroupBox2)
        Me.grpUrun.Controls.Add(Me.txtToplam)
        Me.grpUrun.Controls.Add(Me.Label9)
        Me.grpUrun.Controls.Add(Me.btnEkstraSil)
        Me.grpUrun.Controls.Add(Me.btnEkstraEkle)
        Me.grpUrun.Controls.Add(Me.lstUrunEkstralar)
        Me.grpUrun.Controls.Add(Me.Label8)
        Me.grpUrun.Controls.Add(Me.txtUrunBrmFiyat)
        Me.grpUrun.Controls.Add(Me.Label6)
        Me.grpUrun.Controls.Add(Me.txtUrunUzunluk)
        Me.grpUrun.Controls.Add(Me.Label5)
        Me.grpUrun.Controls.Add(Me.txtUrunGenislik)
        Me.grpUrun.Controls.Add(Me.Label4)
        Me.grpUrun.Controls.Add(Me.txtUrunKodu)
        Me.grpUrun.Controls.Add(Me.Label3)
        Me.grpUrun.Controls.Add(Me.Label1)
        Me.grpUrun.Controls.Add(Me.Label2)
        Me.grpUrun.Controls.Add(Me.cboUrunTuru)
        Me.grpUrun.Controls.Add(Me.lstUrunler)
        Me.grpUrun.Controls.Add(Me.grpUrunListesi)
        Me.grpUrun.Location = New System.Drawing.Point(8, 16)
        Me.grpUrun.Name = "grpUrun"
        Me.grpUrun.Size = New System.Drawing.Size(576, 472)
        Me.grpUrun.TabIndex = 6
        Me.grpUrun.TabStop = False
        '
        'txtGizli
        '
        Me.txtGizli.Location = New System.Drawing.Point(248, 360)
        Me.txtGizli.Name = "txtGizli"
        Me.txtGizli.Size = New System.Drawing.Size(24, 20)
        Me.txtGizli.TabIndex = 34
        Me.txtGizli.Text = ""
        Me.txtGizli.Visible = False
        '
        'lblYTL
        '
        Me.lblYTL.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lblYTL.Location = New System.Drawing.Point(520, 440)
        Me.lblYTL.Name = "lblYTL"
        Me.lblYTL.Size = New System.Drawing.Size(48, 24)
        Me.lblYTL.TabIndex = 33
        Me.lblYTL.Text = "YTL"
        '
        'txtUrunAdi
        '
        Me.txtUrunAdi.AcceptsTab = True
        Me.txtUrunAdi.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtUrunAdi.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunAdi.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunAdi.ForeColor = System.Drawing.Color.Brown
        Me.txtUrunAdi.Location = New System.Drawing.Point(248, 136)
        Me.txtUrunAdi.Name = "txtUrunAdi"
        Me.txtUrunAdi.ReadOnly = True
        Me.txtUrunAdi.Size = New System.Drawing.Size(152, 24)
        Me.txtUrunAdi.TabIndex = 32
        Me.txtUrunAdi.Text = ""
        Me.txtUrunAdi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(248, 120)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(152, 16)
        Me.Label11.TabIndex = 31
        Me.Label11.Text = "Ürün Adý"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnUrunBrmFiyat
        '
        Me.btnUrunBrmFiyat.BackColor = System.Drawing.Color.DarkCyan
        Me.btnUrunBrmFiyat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUrunBrmFiyat.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnUrunBrmFiyat.Location = New System.Drawing.Point(304, 304)
        Me.btnUrunBrmFiyat.Name = "btnUrunBrmFiyat"
        Me.btnUrunBrmFiyat.Size = New System.Drawing.Size(96, 23)
        Me.btnUrunBrmFiyat.TabIndex = 30
        Me.btnUrunBrmFiyat.Text = "Fiyat Deðiþtir"
        '
        'txtUrunTuru
        '
        Me.txtUrunTuru.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtUrunTuru.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunTuru.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunTuru.ForeColor = System.Drawing.Color.Brown
        Me.txtUrunTuru.Location = New System.Drawing.Point(248, 80)
        Me.txtUrunTuru.Name = "txtUrunTuru"
        Me.txtUrunTuru.ReadOnly = True
        Me.txtUrunTuru.Size = New System.Drawing.Size(152, 24)
        Me.txtUrunTuru.TabIndex = 29
        Me.txtUrunTuru.Text = ""
        Me.txtUrunTuru.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(248, 64)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(152, 16)
        Me.Label10.TabIndex = 28
        Me.Label10.Text = "Ürün Türü"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnUzunHesapla
        '
        Me.btnUzunHesapla.BackColor = System.Drawing.Color.DarkCyan
        Me.btnUzunHesapla.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUzunHesapla.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnUzunHesapla.Location = New System.Drawing.Point(504, 216)
        Me.btnUzunHesapla.Name = "btnUzunHesapla"
        Me.btnUzunHesapla.Size = New System.Drawing.Size(64, 23)
        Me.btnUzunHesapla.TabIndex = 27
        Me.btnUzunHesapla.Text = "Hesapla"
        '
        'btnGenHesapla
        '
        Me.btnGenHesapla.BackColor = System.Drawing.Color.DarkCyan
        Me.btnGenHesapla.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnGenHesapla.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnGenHesapla.Location = New System.Drawing.Point(336, 216)
        Me.btnGenHesapla.Name = "btnGenHesapla"
        Me.btnGenHesapla.Size = New System.Drawing.Size(64, 23)
        Me.btnGenHesapla.TabIndex = 26
        Me.btnGenHesapla.Text = "Hesapla"
        '
        'picUrunResmi
        '
        Me.picUrunResmi.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picUrunResmi.Location = New System.Drawing.Point(416, 16)
        Me.picUrunResmi.Name = "picUrunResmi"
        Me.picUrunResmi.Size = New System.Drawing.Size(152, 144)
        Me.picUrunResmi.TabIndex = 19
        Me.picUrunResmi.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Location = New System.Drawing.Point(248, 384)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(320, 16)
        Me.GroupBox3.TabIndex = 18
        Me.GroupBox3.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Location = New System.Drawing.Point(248, 168)
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
        Me.txtToplam.Location = New System.Drawing.Point(296, 432)
        Me.txtToplam.Name = "txtToplam"
        Me.txtToplam.ReadOnly = True
        Me.txtToplam.Size = New System.Drawing.Size(224, 32)
        Me.txtToplam.TabIndex = 16
        Me.txtToplam.Text = ""
        Me.txtToplam.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(296, 408)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(224, 23)
        Me.Label9.TabIndex = 15
        Me.Label9.Text = "...::: TOPLAM :::..."
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnEkstraSil
        '
        Me.btnEkstraSil.BackColor = System.Drawing.Color.DarkCyan
        Me.btnEkstraSil.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEkstraSil.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnEkstraSil.Location = New System.Drawing.Point(496, 352)
        Me.btnEkstraSil.Name = "btnEkstraSil"
        Me.btnEkstraSil.TabIndex = 14
        Me.btnEkstraSil.Text = "Sil"
        '
        'btnEkstraEkle
        '
        Me.btnEkstraEkle.BackColor = System.Drawing.Color.DarkCyan
        Me.btnEkstraEkle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEkstraEkle.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnEkstraEkle.Location = New System.Drawing.Point(416, 352)
        Me.btnEkstraEkle.Name = "btnEkstraEkle"
        Me.btnEkstraEkle.TabIndex = 13
        Me.btnEkstraEkle.Text = "Ekle"
        '
        'lstUrunEkstralar
        '
        Me.lstUrunEkstralar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstUrunEkstralar.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lstUrunEkstralar.ItemHeight = 17
        Me.lstUrunEkstralar.Location = New System.Drawing.Point(416, 280)
        Me.lstUrunEkstralar.Name = "lstUrunEkstralar"
        Me.lstUrunEkstralar.Size = New System.Drawing.Size(152, 70)
        Me.lstUrunEkstralar.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(416, 256)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(152, 23)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Ürün Ekstralar"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUrunBrmFiyat
        '
        Me.txtUrunBrmFiyat.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtUrunBrmFiyat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunBrmFiyat.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunBrmFiyat.Location = New System.Drawing.Point(248, 280)
        Me.txtUrunBrmFiyat.Name = "txtUrunBrmFiyat"
        Me.txtUrunBrmFiyat.ReadOnly = True
        Me.txtUrunBrmFiyat.Size = New System.Drawing.Size(152, 23)
        Me.txtUrunBrmFiyat.TabIndex = 8
        Me.txtUrunBrmFiyat.Text = ""
        Me.txtUrunBrmFiyat.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(248, 256)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(152, 23)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Ürün Birim Fiyatý"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUrunUzunluk
        '
        Me.txtUrunUzunluk.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunUzunluk.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunUzunluk.Location = New System.Drawing.Point(416, 216)
        Me.txtUrunUzunluk.Name = "txtUrunUzunluk"
        Me.txtUrunUzunluk.Size = New System.Drawing.Size(88, 23)
        Me.txtUrunUzunluk.TabIndex = 6
        Me.txtUrunUzunluk.Text = ""
        Me.txtUrunUzunluk.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(416, 192)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(152, 23)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Ürün Uzunluk"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUrunGenislik
        '
        Me.txtUrunGenislik.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunGenislik.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunGenislik.Location = New System.Drawing.Point(248, 216)
        Me.txtUrunGenislik.Name = "txtUrunGenislik"
        Me.txtUrunGenislik.Size = New System.Drawing.Size(88, 23)
        Me.txtUrunGenislik.TabIndex = 4
        Me.txtUrunGenislik.Text = ""
        Me.txtUrunGenislik.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(248, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(152, 23)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Ürün Geniþlik"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUrunKodu
        '
        Me.txtUrunKodu.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtUrunKodu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrunKodu.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtUrunKodu.ForeColor = System.Drawing.Color.Brown
        Me.txtUrunKodu.Location = New System.Drawing.Point(248, 32)
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
        Me.Label3.Location = New System.Drawing.Point(248, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(152, 16)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Ürün Kodu"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MainMnuSiparisler
        '
        Me.MainMnuSiparisler.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSiparisIslemler2})
        '
        'mnuSiparisIslemler2
        '
        Me.mnuSiparisIslemler2.Index = 0
        Me.mnuSiparisIslemler2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSiparisEkle, Me.MenuItem3, Me.mnuSiparisDuzenle, Me.mnuSiparisSil, Me.MenuItem7, Me.mnuKapat})
        Me.mnuSiparisIslemler2.Text = "&Sipariþ Ýþlemleri"
        '
        'mnuSiparisEkle
        '
        Me.mnuSiparisEkle.Index = 0
        Me.mnuSiparisEkle.Text = "Sipariþ Ekl&e"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "-"
        '
        'mnuSiparisDuzenle
        '
        Me.mnuSiparisDuzenle.Index = 2
        Me.mnuSiparisDuzenle.Text = "Sipariþ Dü&zenle"
        '
        'mnuSiparisSil
        '
        Me.mnuSiparisSil.Index = 3
        Me.mnuSiparisSil.Text = "Sipariþ &Sil"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 4
        Me.MenuItem7.Text = "-"
        '
        'mnuKapat
        '
        Me.mnuKapat.Index = 5
        Me.mnuKapat.Text = "&Kapat"
        '
        'frmSiparisler
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(832, 502)
        Me.Controls.Add(Me.grpUrun)
        Me.Controls.Add(Me.grpSiparis)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMnuSiparisler
        Me.Name = "frmSiparisler"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "...::: Sipariþler :::..."
        Me.grpSiparis.ResumeLayout(False)
        Me.grpUrun.ResumeLayout(False)
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
    Private Ekstralar As String
    Private EkstraFiyatSil As Integer
    Private Mode As String = ""
#End Region



#Region "Siparisler Metodu"
    'Sipariþ listesinin doldurulmasý..
    Private Sub SiparisListesiDoldur()

        Me.lstSiparisler.Items.Clear()

        Try
            Dim objListboxNesnesi As ListBoxNesnesi

            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "SELECT * FROM tblSiparis"

            Dim myDR As OleDbDataReader

            myConn.Open()
            myDR = myCmd.ExecuteReader

            Do While myDR.Read
                objListboxNesnesi = New ListBoxNesnesi(myDR.Item("SiparisAdi").ToString, CInt(myDR.Item("ID")))

                Me.lstSiparisler.Items.Add(objListboxNesnesi)
            Loop

            If Not Me.lstSiparisler.Items.Count = 0 Then
                Me.lstSiparisler.SelectedIndex = 0
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

    'Seçilen Siparisin bilgilerinin textboxlara yazýlmasý..
    Private Sub SiparisBilgileri()
        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            Dim objListBoxNesnesi As ListBoxNesnesi
            objListBoxNesnesi = CType(Me.lstSiparisler.SelectedItem, ListBoxNesnesi)

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "SELECT * FROM tblSiparis WHERE ID=@id"

            myCmd.Parameters.Add("@id", CInt(objListBoxNesnesi.No))

            Dim myDR As OleDbDataReader

            myConn.Open()
            myDR = myCmd.ExecuteReader

            'Urun bilgileri textbox lara dolar..
            Do While myDR.Read
                Me.txtSiparisID.Text = myDR.Item("ID")
                Me.txtSiparisAdi.Text = myDR.Item("SiparisAdi")
                Me.txtUrunAdi.Text = myDR.Item("UrunAdi")
                Me.txtUrunKodu.Text = myDR.Item("UrunNo")
                Me.txtUrunTuru.Text = myDR.Item("UrunTuru")
                Me.txtUrunGenislik.Text = myDR.Item("UrunEn")
                Me.txtUrunUzunluk.Text = myDR.Item("UrunBoy")
                Me.txtUrunBrmFiyat.Text = myDR.Item("UrunBrmUcret")

                ResimKonumu = myDR.Item("UrunResim")

                Me.picUrunResmi.Image = Image.FromFile(myDR.Item("UrunResim"))
                Me.chkTeslim.Checked = myDR.Item("UrunTeslim")

                'Resim boyutlandýrýlýyor..
                ShrinkImage(Me.picUrunResmi, Me.picUrunResmi, True)

                Me.lstUrunEkstralar.Items.Clear()

                'Database de UrunEkstra kolonu boþ olduðunda DBNull hatasý veriyor ama hata  oluþunca yapýlacak iþlemler kýsmýna biþiyazmazsak program biraz takýlýyor ama iþleme devam ediyor..
                Ekstralar = myDR.Item("UrunEkstra")

                'Database de urun ekstrasý boþ deðilse,Urun ekstralarini silirek yerine siparisler tablosundaki ekstralari ayirip ekleyecek..
                EkstraCoz()
            Loop

        Catch exOle As OleDbException
            MessageBox.Show(exOle.Message, "OLEDB HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "GENEL HATA !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            'myConn.Close()
            'myCmd.Dispose()
            'myConn.Dispose()
        End Try
    End Sub

    'Siparisler tablosunda tek cumlecik halinde bulunan ekstralarý ayýrýp ekstralar listbox ina dizilmek için hazirlayacak olan sub.. 
    Private Sub EkstraCoz()
        Dim X As String
        Dim Ekstra As String
        Dim Z As Integer = 1

        For i As Integer = 1 To Len(Ekstralar)
            X = Mid(Ekstralar, i, 1)

            If i = Len(Ekstralar) Then
                Ekstra = Mid(Ekstralar, Z, i - Z + 1)
                Me.lstUrunEkstralar.Items.Add(Ekstra)
                Exit For
            End If

            If X = ";" Then
                Ekstra = Mid(Ekstralar, Z, i - Z)
                Me.lstUrunEkstralar.Items.Add(Ekstra)
                Z = i + 1
            End If
        Next

    End Sub

    Private Sub VisibleAyarlari()
        If Mode = "Add" Then
            Me.btnKaydet.Visible = True
            Me.btnIptal.Visible = True
            Me.lstSiparisler.Enabled = False
            Me.chkTeslim.Checked = False

        ElseIf Mode = "Update" Then
            Me.btnKaydet.Visible = True
            Me.btnIptal.Visible = True
            Me.lstSiparisler.Enabled = False

        Else
            Me.btnKaydet.Visible = False
            Me.btnIptal.Visible = False
            Me.lstSiparisler.Enabled = True
        End If
    End Sub

    Private Sub TextboxTemizle()
        If Not Me.lstUrunler.Items.Count = 0 Then
            Me.lstUrunler.SelectedIndex = 0
        Else
            Me.txtSiparisAdi.Text = Nothing
            Me.txtSiparisID.Text = Nothing
            Me.txtToplam.Text = Nothing
            Me.txtUrunAdi.Text = Nothing
            Me.txtUrunBrmFiyat.Text = Nothing
            Me.txtUrunGenislik.Text = Nothing
            Me.txtUrunKodu.Text = Nothing
            Me.txtUrunTuru.Text = Nothing
            Me.txtUrunUzunluk.Text = Nothing
            Me.picUrunResmi.Image = Nothing
            Me.lstSiparisler.Items.Clear()
            Me.lstUrunEkstralar.Items.Clear()
        End If
    End Sub

    'Yeni siparis ekleme..
    Private Sub SiparisEkle()
        Dim Ekstralar As String = ""

        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "INSERT INTO tblSiparis " & _
                " (SiparisAdi, UrunAdi, UrunTuru, UrunNo, UrunEn, UrunBoy, UrunBrmUcret, UrunEkstra, UrunResim, UrunTeslim) " & _
                " VALUES (@siparisadi, @urunadi, @urunturu, @urunno, @urunen, @urunboy, @urunbrmucret, @urunekstra, @urunresim, @urunteslim)"

            If Not Me.lstUrunEkstralar.Items.Count = 0 Then
                For i As Integer = 0 To Me.lstUrunEkstralar.Items.Count - 1
                    Ekstralar += Me.lstUrunEkstralar.Items(i) & ";"
                Next

                Ekstralar = Mid(Ekstralar, 1, Len(Ekstralar) - 1)
            End If

            myCmd.Parameters.Add("@siparisadi", Me.txtSiparisAdi.Text)
            myCmd.Parameters.Add("@urunadi", Me.txtUrunAdi.Text)
            myCmd.Parameters.Add("@urunturu", Me.txtUrunTuru.Text)
            myCmd.Parameters.Add("@urunno", Me.txtUrunKodu.Text)
            myCmd.Parameters.Add("@urunen", Me.txtUrunGenislik.Text)
            myCmd.Parameters.Add("@urunboy", Me.txtUrunUzunluk.Text)
            myCmd.Parameters.Add("@urunbrmucret", Me.txtUrunBrmFiyat.Text)
            myCmd.Parameters.Add("@urunekstra", Ekstralar)
            myCmd.Parameters.Add("@urunresim", ResimKonumu)
            myCmd.Parameters.Add("@urunteslim", Me.chkTeslim.Checked)

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

    'Sipariþ düzenleme..
    Private Sub SiparisDuzenle()
        Dim Ekstralar As String = Nothing

        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            Dim objListBoxNesnesi As ListBoxNesnesi
            objListBoxNesnesi = CType(Me.lstSiparisler.SelectedItem, ListBoxNesnesi)

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = _
                "UPDATE tblSiparis SET " & _
                    "SiparisAdi=@siparisadi, " & _
                    "UrunAdi=@urunadi, " & _
                    "UrunTuru=@urunturu, " & _
                    "UrunNo=@urunno, " & _
                    "UrunEn=@urunen, " & _
                    "UrunBoy=@urunboy, " & _
                    "UrunBrmUcret=@urunbrmucret, " & _
                    "UrunEkstra=@urunekstra, " & _
                    "UrunResim=@urunresim, " & _
                    "UrunTeslim=@urunteslim " & _
                "WHERE ID=" & objListBoxNesnesi.No

            If Not Me.lstUrunEkstralar.Items.Count = 0 Then
                For i As Integer = 0 To Me.lstUrunEkstralar.Items.Count - 1
                    Ekstralar += Me.lstUrunEkstralar.Items(i) & ";"
                Next

                Ekstralar = Mid(Ekstralar, 1, Len(Ekstralar) - 1)
            End If

            myCmd.Parameters.Add("@siparisadi", Me.txtSiparisAdi.Text)
            myCmd.Parameters.Add("@urunadi", Me.txtUrunAdi.Text)
            myCmd.Parameters.Add("@urunturu", Me.txtUrunTuru.Text)
            myCmd.Parameters.Add("@urunno", Me.txtUrunKodu.Text)
            myCmd.Parameters.Add("@urunen", Me.txtUrunGenislik.Text)
            myCmd.Parameters.Add("@urunboy", Me.txtUrunUzunluk.Text)
            myCmd.Parameters.Add("@urunbrmucret", Me.txtUrunBrmFiyat.Text)
            myCmd.Parameters.Add("@urunekstra", Ekstralar)
            myCmd.Parameters.Add("@urunresim", ResimKonumu)
            myCmd.Parameters.Add("@urunteslim", Me.chkTeslim.Checked)

            myConn.Open()
            myCmd.ExecuteNonQuery()

            MessageBox.Show("Sipariþ güncelleþtirildi..", "Sipariþ Düzenlendi", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

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

    'Var olan siparisin silinmesi..
    Private Sub SiparisSil()
        Try

            If Me.lstSiparisler.SelectedIndex = -1 Then
                MessageBox.Show("Listeden silinecek sipariþi seçiniz..", "Sipariþ Seçmediniz..", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Try
            End If

            Dim objListboxNesnesi As New ListBoxNesnesi
            objListboxNesnesi = CType(Me.lstSiparisler.SelectedItem, ListBoxNesnesi)

            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "DELETE FROM tblSiparis WHERE ID=@id"

            myCmd.Parameters.Add("@id", CInt(objListboxNesnesi.No))

            myConn.Open()
            myCmd.ExecuteNonQuery()

            MessageBox.Show("Sipariþ silinmiþtir..", "Sipariþ Silindi..", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

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

#End Region

#Region "Urunler Metodu"
        'Urun turu combobox ina el ile deger girilirse varliginin kontrolu..
    Private Sub UrunTuruTextKontrol()
        Dim ListedeVar As Boolean = False

        'Eðer Urunturu combo box ýnýn text kýsmýna el ile içeriðinde olmayan bir Urun turu girilirse kontrol edecek, text i sýfýrlýyacak yada database e baðlanacak..
        For i As Integer = 0 To Me.cboUrunTuru.Items.Count - 1
            If Me.cboUrunTuru.Text = Me.cboUrunTuru.Items(i) Then
                ListedeVar = True
            End If
        Next
    End Sub

    'urun turu combobox inin doldurulmasi..
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

    'Secilen urun turune göre urunlerin listelenmesi..
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

    'Secilen urun bilgilerinin textboxlara yazilmasi..
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
            Me.txtUrunTuru.Text = Me.cboUrunTuru.Text

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

    'Urunun birim fiyatina göre hesap yapilmasi..
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

    'Urun ekstrasina girilen özelligin fiyatinin isminden ayrilmasi..
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

    'Urun resminin anti_alias ile kucultulmesi..
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

    Private Sub frmSiparisler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UrunTuruDoldur()
        SiparisListesiDoldur()
        Mode = ""
        VisibleAyarlari()
    End Sub

    Private Sub mnuKapat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKapat.Click
        Me.Close()
    End Sub

    'Sipariþler formu kapatýlýnca MdiParent da "Sipariþ Ýþlemleri" menuitem ýnýn aktif hale getirilmesi..
    Private Sub frmSiparisler_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Me.MdiParent.Menu.MenuItems.Item(0).MenuItems.Item(2).Enabled = True
    End Sub

    Private Sub cboUrunTuru_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUrunTuru.SelectedIndexChanged
        Me.lstUrunler.Items.Clear()

        'Seçilen ürün kategorisne göre Database e ekleme yada düzenleme yada silme yapýldýðýnda gidilecek tabloyu hafýzada tutuyoruz..
        Select Case Me.cboUrunTuru.Text
            Case "Dolaplar"
                txtUrunTuru.Text = "Dolaplar"
                Urun_Turu_Tablosu = "tblDolaplar"
            Case "Kapýlar"
                txtUrunTuru.Text = "Kapýlar"
                Urun_Turu_Tablosu = "tblKapilar"
            Case "Koltuklar"
                txtUrunTuru.Text = "Koltuklar"
                Urun_Turu_Tablosu = "tblKoltuklar"
            Case "Masalar"
                txtUrunTuru.Text = "Masalar"
                Urun_Turu_Tablosu = "tblMasalar"
            Case "Diðer"
                txtUrunTuru.Text = "Diðer"
                Urun_Turu_Tablosu = "tblDiger"
        End Select

        UrunListesiDoldur()
    End Sub

    Private Sub lstUrunler_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstUrunler.SelectedIndexChanged
        UrunBilgileri()
        ToplamHesapla()
    End Sub

    Private Sub cboUrunTuru_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboUrunTuru.TextChanged
        UrunTuruTextKontrol()

        If ListedeVar = False Then
            Me.cboUrunTuru.Text = "Ürün Türü Seçiniz..."
            Exit Sub
        End If
    End Sub

    Private Sub btnEkstraEkle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEkstraEkle.Click
        Dim UrunEkstra As String
        Dim UrunEkstraFiyati As Integer

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

    Private Sub btnUrunBrmFiyat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUrunBrmFiyat.Click
        Dim YeniBirimFiyat As Integer

        On Error GoTo errhandler

        'Eger urun birim fiyatý butonuna týklatilip sonra cancel yapilirsa errhandler a gidilecek..
        YeniBirimFiyat = InputBox("Ürünün yeni birim fiyatýný giriniz: " & vbCrLf & "ÖRN: 6 gibi..")


        If Not IsNumeric(YeniBirimFiyat) = True Then
            MessageBox.Show("Lütfen Yeni Birim Fiyatý Bir Sayý Olarak Giriniz..", "HATALI DEÐER", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            Exit Sub
        End If

        Me.txtUrunBrmFiyat.Text = YeniBirimFiyat

errhandler:
        ToplamHesapla()
    End Sub

    'Hem genislik hem uzunluk hesapla butonuna basilinca devreye girecek sub..
    Private Sub btnGenHesapla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenHesapla.Click, btnUzunHesapla.Click
        ToplamHesapla()
    End Sub

    Private Sub lstSiparisler_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstSiparisler.SelectedIndexChanged
        SiparisBilgileri()
        ToplamHesapla()
    End Sub

    Private Sub mnuSiparisEkle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSiparisEkle.Click
        Mode = "Add"
        VisibleAyarlari()
        TextboxTemizle()
        Me.picUrunResmi.Image = Nothing
        ResimKonumu = Nothing
    End Sub

    Private Sub mnuSiparisDuzenle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSiparisDuzenle.Click
        Mode = "Update"
        VisibleAyarlari()
    End Sub

    Private Sub mnuSiparisSil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSiparisSil.Click
        DialogResult = MessageBox.Show(Me.lstSiparisler.SelectedItem.ToString & " adlý kiþiye ait sipariþi silmek istediðinizden emin misiniz?", "Sipariþ Silme", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

        If Not DialogResult = DialogResult.No Then
            SiparisSil()
            SiparisListesiDoldur()

            If Not Me.lstSiparisler.Items.Count = 0 Then
                Me.lstSiparisler.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub btnKaydet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKaydet.Click
        If Mode = "Add" Then
            SiparisEkle()
        ElseIf Mode = "Update" Then
            SiparisDuzenle()
        End If

        Mode = ""
        VisibleAyarlari()
        SiparisListesiDoldur()
    End Sub

    Private Sub btnIptal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIptal.Click
        Mode = ""
        VisibleAyarlari()
        SiparisListesiDoldur()
    End Sub

    Private Sub cmnuSiparisEkle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuSiparisEkle.Click
        mnuSiparisEkle_Click(sender, e)
    End Sub

    Private Sub cmnuSiparisDuzenle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuSiparisDuzenle.Click
        mnuSiparisDuzenle_Click(sender, e)
    End Sub

    Private Sub cmnuSiparisSil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuSiparisSil.Click
        mnuSiparisSil_Click(sender, e)
    End Sub





#End Region



End Class
