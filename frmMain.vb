#Region "imports"
Imports System
Imports System.Data
Imports System.Data.OleDb
#End Region

Public Class frmMain
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
    Friend WithEvents MainMnuAna As System.Windows.Forms.MainMenu
    Friend WithEvents mnuCikis As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuIslemler As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUrunIslemleri1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMusteriIslemleri1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSiparisIslemleri1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.MainMnuAna = New System.Windows.Forms.MainMenu
        Me.mnuIslemler = New System.Windows.Forms.MenuItem
        Me.mnuUrunIslemleri1 = New System.Windows.Forms.MenuItem
        Me.mnuMusteriIslemleri1 = New System.Windows.Forms.MenuItem
        Me.mnuSiparisIslemleri1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.mnuCikis = New System.Windows.Forms.MenuItem
        '
        'MainMnuAna
        '
        Me.MainMnuAna.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuIslemler})
        '
        'mnuIslemler
        '
        Me.mnuIslemler.Index = 0
        Me.mnuIslemler.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuUrunIslemleri1, Me.mnuMusteriIslemleri1, Me.mnuSiparisIslemleri1, Me.MenuItem2, Me.mnuCikis})
        Me.mnuIslemler.Text = "Ýþlemler"
        '
        'mnuUrunIslemleri1
        '
        Me.mnuUrunIslemleri1.Index = 0
        Me.mnuUrunIslemleri1.Text = "&Ürün Ýþlemleri"
        '
        'mnuMusteriIslemleri1
        '
        Me.mnuMusteriIslemleri1.Index = 1
        Me.mnuMusteriIslemleri1.Text = "&Müþteri Ýþlemleri"
        '
        'mnuSiparisIslemleri1
        '
        Me.mnuSiparisIslemleri1.Index = 2
        Me.mnuSiparisIslemleri1.Text = "&Sipariþ Ýþlemleri"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.Text = "-"
        '
        'mnuCikis
        '
        Me.mnuCikis.Index = 4
        Me.mnuCikis.Text = "Çý&kýþ"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(528, 266)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMnuAna
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Mobilya Otomasyonu"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

#End Region

    'Urunler form u açýlýr..
    Private Sub mnuUrunIslemleri_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUrunIslemleri1.Click
        Dim frmUrunler As New frmUrunler

        frmUrunler.MdiParent = Me
        frmUrunler.Show()

        Me.mnuUrunIslemleri1.Enabled = False
    End Sub

    'Müþteriler form u açýlýr..
    Private Sub mnuMusteriIslemleri_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMusteriIslemleri1.Click
        Dim frmMusteriler As New frmMusteriler

        frmMusteriler.MdiParent = Me
        frmMusteriler.Show()

        Me.mnuMusteriIslemleri1.Enabled = False
    End Sub

    'Sipariþler form u açýlýr..
    Private Sub mnuSiparisIslemleri_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSiparisIslemleri1.Click
        Dim frmSiparisler As New frmSiparisler

        frmSiparisler.MdiParent = Me
        frmSiparisler.Show()

        Me.mnuSiparisIslemleri1.Enabled = False
    End Sub

    'Programdan çýkýþ..
    Private Sub mnuCikis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCikis.Click
        DialogResult = MessageBox.Show("Çýkmak istediðinize emin misiniz?", "Çýkýþ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly)

        If DialogResult = DialogResult.Yes Then
            If Not Me.MdiChildren.Length = 0 Then
                For i As Integer = 0 To Me.MdiChildren.Length - 1
                    Me.MdiChildren(0).Close()
                Next
            End If

            End
        Else
            Exit Sub
        End If
    End Sub
End Class
