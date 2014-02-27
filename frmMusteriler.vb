#Region "imports"
Imports System
Imports System.Data
Imports System.Data.OleDb
#End Region

Public Class frmMusteriler
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents MainMnuMusteriler As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents grpMusteriler As System.Windows.Forms.GroupBox
    Friend WithEvents grpMusteriBilgileri As System.Windows.Forms.GroupBox
    Friend WithEvents txtMusteriID As System.Windows.Forms.TextBox
    Friend WithEvents txtMusteriAd As System.Windows.Forms.TextBox
    Friend WithEvents txtMusteriSoyad As System.Windows.Forms.TextBox
    Friend WithEvents txtMusteriTel As System.Windows.Forms.TextBox
    Friend WithEvents txtMusteriAdres As System.Windows.Forms.TextBox
    Friend WithEvents txtMusteriNot As System.Windows.Forms.TextBox
    Friend WithEvents lstMusteriler As System.Windows.Forms.ListBox
    Friend WithEvents btnKaydet As System.Windows.Forms.Button
    Friend WithEvents btnIptal As System.Windows.Forms.Button
    Friend WithEvents mnuMusteriEkle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMusteriDuzenle As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMusteriSil As System.Windows.Forms.MenuItem
    Friend WithEvents mnuKapat As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMusteriIslemleri2 As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuMusteriEkle As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuMusteriDuzenle As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuMusteriSil As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMusteriler))
        Me.grpMusteriler = New System.Windows.Forms.GroupBox
        Me.lstMusteriler = New System.Windows.Forms.ListBox
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.cmnuMusteriEkle = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.cmnuMusteriDuzenle = New System.Windows.Forms.MenuItem
        Me.cmnuMusteriSil = New System.Windows.Forms.MenuItem
        Me.Label7 = New System.Windows.Forms.Label
        Me.grpMusteriBilgileri = New System.Windows.Forms.GroupBox
        Me.btnIptal = New System.Windows.Forms.Button
        Me.btnKaydet = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtMusteriNot = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtMusteriAdres = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtMusteriTel = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtMusteriAd = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtMusteriID = New System.Windows.Forms.TextBox
        Me.txtMusteriSoyad = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.MainMnuMusteriler = New System.Windows.Forms.MainMenu
        Me.mnuMusteriIslemleri2 = New System.Windows.Forms.MenuItem
        Me.mnuMusteriEkle = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuMusteriDuzenle = New System.Windows.Forms.MenuItem
        Me.mnuMusteriSil = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuKapat = New System.Windows.Forms.MenuItem
        Me.grpMusteriler.SuspendLayout()
        Me.grpMusteriBilgileri.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMusteriler
        '
        Me.grpMusteriler.Controls.Add(Me.lstMusteriler)
        Me.grpMusteriler.Controls.Add(Me.Label7)
        Me.grpMusteriler.Location = New System.Drawing.Point(16, 16)
        Me.grpMusteriler.Name = "grpMusteriler"
        Me.grpMusteriler.Size = New System.Drawing.Size(216, 280)
        Me.grpMusteriler.TabIndex = 0
        Me.grpMusteriler.TabStop = False
        '
        'lstMusteriler
        '
        Me.lstMusteriler.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstMusteriler.ContextMenu = Me.ContextMenu1
        Me.lstMusteriler.Location = New System.Drawing.Point(8, 40)
        Me.lstMusteriler.Name = "lstMusteriler"
        Me.lstMusteriler.Size = New System.Drawing.Size(200, 236)
        Me.lstMusteriler.Sorted = True
        Me.lstMusteriler.TabIndex = 7
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmnuMusteriEkle, Me.MenuItem4, Me.cmnuMusteriDuzenle, Me.cmnuMusteriSil})
        '
        'cmnuMusteriEkle
        '
        Me.cmnuMusteriEkle.Index = 0
        Me.cmnuMusteriEkle.Text = "Müþteri Ekle"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.Text = "-"
        '
        'cmnuMusteriDuzenle
        '
        Me.cmnuMusteriDuzenle.Index = 2
        Me.cmnuMusteriDuzenle.Text = "Müþteri Düzenle"
        '
        'cmnuMusteriSil
        '
        Me.cmnuMusteriSil.Index = 3
        Me.cmnuMusteriSil.Text = "Müþteri Sil"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(8, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(200, 23)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Müþteriler"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grpMusteriBilgileri
        '
        Me.grpMusteriBilgileri.Controls.Add(Me.btnIptal)
        Me.grpMusteriBilgileri.Controls.Add(Me.btnKaydet)
        Me.grpMusteriBilgileri.Controls.Add(Me.GroupBox3)
        Me.grpMusteriBilgileri.Controls.Add(Me.txtMusteriNot)
        Me.grpMusteriBilgileri.Controls.Add(Me.Label6)
        Me.grpMusteriBilgileri.Controls.Add(Me.txtMusteriAdres)
        Me.grpMusteriBilgileri.Controls.Add(Me.Label5)
        Me.grpMusteriBilgileri.Controls.Add(Me.txtMusteriTel)
        Me.grpMusteriBilgileri.Controls.Add(Me.Label2)
        Me.grpMusteriBilgileri.Controls.Add(Me.txtMusteriAd)
        Me.grpMusteriBilgileri.Controls.Add(Me.Label3)
        Me.grpMusteriBilgileri.Controls.Add(Me.txtMusteriID)
        Me.grpMusteriBilgileri.Controls.Add(Me.txtMusteriSoyad)
        Me.grpMusteriBilgileri.Controls.Add(Me.Label1)
        Me.grpMusteriBilgileri.Controls.Add(Me.Label4)
        Me.grpMusteriBilgileri.Location = New System.Drawing.Point(248, 16)
        Me.grpMusteriBilgileri.Name = "grpMusteriBilgileri"
        Me.grpMusteriBilgileri.Size = New System.Drawing.Size(440, 280)
        Me.grpMusteriBilgileri.TabIndex = 1
        Me.grpMusteriBilgileri.TabStop = False
        '
        'btnIptal
        '
        Me.btnIptal.BackColor = System.Drawing.Color.DarkCyan
        Me.btnIptal.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnIptal.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnIptal.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnIptal.Location = New System.Drawing.Point(8, 240)
        Me.btnIptal.Name = "btnIptal"
        Me.btnIptal.Size = New System.Drawing.Size(96, 23)
        Me.btnIptal.TabIndex = 23
        Me.btnIptal.Text = "Ýptal"
        '
        'btnKaydet
        '
        Me.btnKaydet.BackColor = System.Drawing.Color.DarkCyan
        Me.btnKaydet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnKaydet.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnKaydet.ForeColor = System.Drawing.Color.PaleTurquoise
        Me.btnKaydet.Location = New System.Drawing.Point(112, 240)
        Me.btnKaydet.Name = "btnKaydet"
        Me.btnKaydet.Size = New System.Drawing.Size(96, 23)
        Me.btnKaydet.TabIndex = 22
        Me.btnKaydet.Text = "Kaydet"
        '
        'GroupBox3
        '
        Me.GroupBox3.Location = New System.Drawing.Point(216, 8)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(8, 264)
        Me.GroupBox3.TabIndex = 13
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "GroupBox3"
        '
        'txtMusteriNot
        '
        Me.txtMusteriNot.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtMusteriNot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMusteriNot.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMusteriNot.Location = New System.Drawing.Point(232, 168)
        Me.txtMusteriNot.Multiline = True
        Me.txtMusteriNot.Name = "txtMusteriNot"
        Me.txtMusteriNot.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMusteriNot.Size = New System.Drawing.Size(200, 104)
        Me.txtMusteriNot.TabIndex = 12
        Me.txtMusteriNot.Text = ""
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(232, 144)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(200, 23)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Müþteri Hakkýnda Kýsa Not"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMusteriAdres
        '
        Me.txtMusteriAdres.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtMusteriAdres.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMusteriAdres.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMusteriAdres.Location = New System.Drawing.Point(232, 104)
        Me.txtMusteriAdres.Name = "txtMusteriAdres"
        Me.txtMusteriAdres.Size = New System.Drawing.Size(200, 24)
        Me.txtMusteriAdres.TabIndex = 10
        Me.txtMusteriAdres.Text = ""
        Me.txtMusteriAdres.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(232, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(200, 23)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Müþteri Adres"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMusteriTel
        '
        Me.txtMusteriTel.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtMusteriTel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMusteriTel.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMusteriTel.Location = New System.Drawing.Point(232, 40)
        Me.txtMusteriTel.Name = "txtMusteriTel"
        Me.txtMusteriTel.Size = New System.Drawing.Size(200, 24)
        Me.txtMusteriTel.TabIndex = 8
        Me.txtMusteriTel.Text = ""
        Me.txtMusteriTel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(232, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(200, 23)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Müþteri Telefon Numarasý"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMusteriAd
        '
        Me.txtMusteriAd.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtMusteriAd.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMusteriAd.Font = New System.Drawing.Font("Times New Roman", 10.08!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtMusteriAd.Location = New System.Drawing.Point(8, 104)
        Me.txtMusteriAd.Name = "txtMusteriAd"
        Me.txtMusteriAd.Size = New System.Drawing.Size(200, 24)
        Me.txtMusteriAd.TabIndex = 6
        Me.txtMusteriAd.Text = ""
        Me.txtMusteriAd.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(200, 23)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Müþteri Adý"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMusteriID
        '
        Me.txtMusteriID.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtMusteriID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMusteriID.Font = New System.Drawing.Font("Times New Roman", 10.08!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtMusteriID.ForeColor = System.Drawing.Color.Brown
        Me.txtMusteriID.Location = New System.Drawing.Point(8, 40)
        Me.txtMusteriID.Name = "txtMusteriID"
        Me.txtMusteriID.ReadOnly = True
        Me.txtMusteriID.Size = New System.Drawing.Size(200, 24)
        Me.txtMusteriID.TabIndex = 4
        Me.txtMusteriID.Text = ""
        Me.txtMusteriID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtMusteriSoyad
        '
        Me.txtMusteriSoyad.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtMusteriSoyad.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMusteriSoyad.Font = New System.Drawing.Font("Times New Roman", 10.08!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.txtMusteriSoyad.Location = New System.Drawing.Point(8, 168)
        Me.txtMusteriSoyad.Name = "txtMusteriSoyad"
        Me.txtMusteriSoyad.Size = New System.Drawing.Size(200, 24)
        Me.txtMusteriSoyad.TabIndex = 8
        Me.txtMusteriSoyad.Text = ""
        Me.txtMusteriSoyad.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(200, 23)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Müþteri ID"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(200, 23)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Müþteri Soyadý"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MainMnuMusteriler
        '
        Me.MainMnuMusteriler.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuMusteriIslemleri2})
        '
        'mnuMusteriIslemleri2
        '
        Me.mnuMusteriIslemleri2.Index = 0
        Me.mnuMusteriIslemleri2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuMusteriEkle, Me.MenuItem3, Me.mnuMusteriDuzenle, Me.mnuMusteriSil, Me.MenuItem1, Me.mnuKapat})
        Me.mnuMusteriIslemleri2.Text = "&Müþteri Ýþlemleri"
        '
        'mnuMusteriEkle
        '
        Me.mnuMusteriEkle.Index = 0
        Me.mnuMusteriEkle.Text = "Müþteri Ekl&e"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "-"
        '
        'mnuMusteriDuzenle
        '
        Me.mnuMusteriDuzenle.Index = 2
        Me.mnuMusteriDuzenle.Text = "Müþteri Dü&zenle"
        '
        'mnuMusteriSil
        '
        Me.mnuMusteriSil.Index = 3
        Me.mnuMusteriSil.Text = "Müþteri &Sil"
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
        'frmMusteriler
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(698, 307)
        Me.Controls.Add(Me.grpMusteriBilgileri)
        Me.Controls.Add(Me.grpMusteriler)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMnuMusteriler
        Me.Name = "frmMusteriler"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "...::: Müþteriler :::..."
        Me.grpMusteriler.ResumeLayout(False)
        Me.grpMusteriBilgileri.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Yerel Degiskenler"
    Private myConn As OleDbConnection
    Private myCmd As OleDbCommand
    Private myDR As OleDbDataReader
    Private myConnStr As String = _
        "Provider=Microsoft.ace.oledb.12.0;Data Source=" & Application.StartupPath & "\database.mdb;"

    Private Mode As String = Nothing
#End Region


#Region "Metotlar"

    'Müþteri ekleme, düzenleme veya diðer iþlemlere göre ekran nesnelerinin visibility sinin ayarlanmasý..
    Private Sub VisibleAyarlari()
        If Mode = "Add" OrElse Mode = "Update" Then
            Me.btnKaydet.Visible = True
            Me.btnIptal.Visible = True

        Else
            Me.btnKaydet.Visible = False
            Me.btnIptal.Visible = False
        End If
    End Sub

    Private Sub MusteriListele()
        Try
            Me.lstMusteriler.Items.Clear()

            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "SELECT * FROM tblUyeler "

            Dim myDR As OleDbDataReader

            myConn.Open()
            myDR = myCmd.ExecuteReader

            Do While myDR.Read
                Dim objListBoxNesnesi As ListBoxNesnesi
                objListBoxNesnesi = New ListBoxNesnesi(myDR.Item("Ad").ToString, myDR.Item("ID"))

                Me.lstMusteriler.Items.Add(objListBoxNesnesi)
            Loop

            If Not Me.lstMusteriler.Items.Count = 0 Then
                Me.lstMusteriler.SelectedIndex = 0
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

    'Listeden seçilen müþteri bilgilerinin ekranda gösterilmesi..
    Private Sub MusteriBilgileri()
        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            Dim objListBoxNesnesi As ListBoxNesnesi
            objListBoxNesnesi = CType(Me.lstMusteriler.SelectedItem, ListBoxNesnesi)

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "SELECT * FROM tblUyeler WHERE ID=@id"

            myCmd.Parameters.Add("@id", objListBoxNesnesi.No)

            Dim myDR As OleDbDataReader

            myConn.Open()
            myDR = myCmd.ExecuteReader

            '*******************************
            'Textbox larýn doldurulmasý..
            '*******************************
            Do While myDR.Read
                Me.txtMusteriID.Text = myDR.Item("ID")
                Me.txtMusteriAd.Text = myDR.Item("Ad")
                Me.txtMusteriSoyad.Text = myDR.Item("Soyad")
                Me.txtMusteriAdres.Text = myDR.Item("Adres")
                Me.txtMusteriTel.Text = myDR.Item("Tel")
                Me.txtMusteriNot.Text = myDR.Item("Bilgi")
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

    'Yeni müþteri ekleme..
    Private Sub MusteriEkle()
        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "INSERT INTO tblUyeler " & _
                " (Ad, Soyad, Tel, Adres, Bilgi) " & _
                " VALUES (@ad, @soyad, @tel, @adres, @bilgi)"

            myCmd.Parameters.Add("@ad", Me.txtMusteriAd.Text)
            myCmd.Parameters.Add("@soyad", Me.txtMusteriSoyad.Text)
            myCmd.Parameters.Add("@tel", Me.txtMusteriTel.Text)
            myCmd.Parameters.Add("@adres", Me.txtMusteriAdres.Text)
            myCmd.Parameters.Add("@bilgi", Me.txtMusteriNot.Text)

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

    'Müþteri düzenleme..
    Private Sub MusteriDuzenle()
        Try
            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            Dim objListBoxNesnesi As ListBoxNesnesi
            objListBoxNesnesi = CType(Me.lstMusteriler.SelectedItem, ListBoxNesnesi)

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = _
                "UPDATE tblUyeler SET " & _
                    "Ad=@ad, " & _
                    "Soyad=@soyad, " & _
                    "Tel=@tel, " & _
                    "Adres=@adres, " & _
                    "Bilgi=@bilgi " & _
                "WHERE ID=" & objListBoxNesnesi.No

            myCmd.Parameters.Add("@ad", Me.txtMusteriAd.Text)
            myCmd.Parameters.Add("@soyad", Me.txtMusteriSoyad.Text)
            myCmd.Parameters.Add("@tel", Me.txtMusteriTel.Text)
            myCmd.Parameters.Add("@adres", Me.txtMusteriAdres.Text)
            myCmd.Parameters.Add("@bilgi", Me.txtMusteriNot.Text)

            myConn.Open()
            myCmd.ExecuteNonQuery()

            MessageBox.Show("Müþteri Bilgisi Güncelleþtirildi..", "Müþteri Güncelleme", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

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

    'Müþteri silme..
    Private Sub MusteriSil()
        Try

            If Me.lstMusteriler.SelectedIndex = -1 Then
                MessageBox.Show("Listeden silinecek ürünü seçiniz..", "Ürün Seçmediniz..", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Try
            End If

            Dim objListboxNesnesi As New ListBoxNesnesi
            objListboxNesnesi = CType(Me.lstMusteriler.SelectedItem, ListBoxNesnesi)

            myConn = New OleDbConnection
            myConn.ConnectionString = myConnStr

            myCmd = New OleDbCommand
            myCmd.Connection = myConn
            myCmd.CommandType = CommandType.Text
            myCmd.CommandText = "DELETE FROM tblUyeler WHERE ID=@id"

            myCmd.Parameters.Add("@id", CInt(objListboxNesnesi.No))

            myConn.Open()
            myCmd.ExecuteNonQuery()

            MessageBox.Show("Müþteri silinmiþtir..", "Müþteri Silme", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

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

    'Yeni müþteri kaydýnda textboxlarýn temizlenmesi..
    Private Sub TextboxTemizle()
        Me.txtMusteriID.Text = ""
        Me.txtMusteriAd.Text = ""
        Me.txtMusteriSoyad.Text = ""
        Me.txtMusteriTel.Text = ""
        Me.txtMusteriAdres.Text = ""
        Me.txtMusteriNot.Text = ""
        Me.lstMusteriler.Items.Clear()
    End Sub

#End Region


#Region "Olaylar"

    Private Sub mnuKapat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKapat.Click
        Me.Close()
    End Sub

    'Müþteriler formu kapatýlýnca MdiParent da "Müþteri Ýþlemleri" menuitem ýnýn aktif hale getirilmesi..
    Private Sub frmMusteriler_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Me.MdiParent.Menu.MenuItems.Item(0).MenuItems.Item(1).Enabled = True
    End Sub

    Private Sub frmMusteriler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        VisibleAyarlari()
        MusteriListele()
    End Sub

    Private Sub lstMusteriler_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstMusteriler.SelectedIndexChanged
        MusteriBilgileri()
    End Sub

    Private Sub mnuMusteriEkle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMusteriEkle.Click
        Mode = "Add"
        VisibleAyarlari()
        TextboxTemizle()
    End Sub

    Private Sub mnuMusteriDuzenle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMusteriDuzenle.Click
        Mode = "Update"
        VisibleAyarlari()
    End Sub

    Private Sub mnuMusteriSil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMusteriSil.Click
        DialogResult = MessageBox.Show(Me.lstMusteriler.SelectedItem.ToString & " adlý müþteriyi Silmek istediðinize emin misiniz?", "Silme Ýþlemi Onay", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

        If DialogResult = DialogResult.Yes Then
            MusteriSil()
        End If

        MusteriListele()
    End Sub

    Private Sub btnKaydet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKaydet.Click
        If Mode = "Add" Then
            MusteriEkle()
        ElseIf Mode = "Update" Then
            MusteriDuzenle()
        End If

        Mode = Nothing
        VisibleAyarlari()
        MusteriListele()
    End Sub

    Private Sub btnIptal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIptal.Click
        Mode = Nothing
        VisibleAyarlari()
        MusteriListele()
    End Sub

    Private Sub cmnuMusteriEkle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuMusteriEkle.Click
        mnuMusteriEkle_Click(sender, e)
    End Sub

    Private Sub cmnuMusteriDuzenle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuMusteriDuzenle.Click
        mnuMusteriDuzenle_Click(sender, e)
    End Sub

    Private Sub cmnuMusteriSil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuMusteriSil.Click
        mnuMusteriSil_Click(sender, e)
    End Sub

#End Region

End Class
