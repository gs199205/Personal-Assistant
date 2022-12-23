Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class frmPilgrim
    Inherits System.Windows.Forms.UserControl

    Private pageAction As String
    Private selectedItemId As Int32

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
    Friend WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents grdData As System.Windows.Forms.DataGrid
    Friend WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents pnlData As System.Windows.Forms.GroupBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtUrl As System.Windows.Forms.TextBox
    Friend WithEvents txtHomePhone As System.Windows.Forms.TextBox
    Friend WithEvents txtCellPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtWorkPhone As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cmbCategory As System.Windows.Forms.ComboBox
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents NameField As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents Address As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents IDField As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents dtBirthDay As System.Windows.Forms.DateTimePicker
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSearchbyName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents HomePhoneField As System.Windows.Forms.DataGridTextBoxColumn
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.grdData = New System.Windows.Forms.DataGrid
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle
        Me.IDField = New System.Windows.Forms.DataGridTextBoxColumn
        Me.NameField = New System.Windows.Forms.DataGridTextBoxColumn
        Me.Address = New System.Windows.Forms.DataGridTextBoxColumn
        Me.HomePhoneField = New System.Windows.Forms.DataGridTextBoxColumn
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.pnlData = New System.Windows.Forms.GroupBox
        Me.dtBirthDay = New System.Windows.Forms.DateTimePicker
        Me.cmbCategory = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtUrl = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtWorkPhone = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtCellPhone = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtHomePhone = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtSearchbyName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        CType(Me.grdData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.pnlData.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "Name"
        Me.DataGridTextBoxColumn1.MappingName = "Name"
        Me.DataGridTextBoxColumn1.NullText = ""
        Me.DataGridTextBoxColumn1.ReadOnly = True
        Me.DataGridTextBoxColumn1.Width = 75
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Format = ""
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.HeaderText = "Phone"
        Me.DataGridTextBoxColumn2.MappingName = "Phone"
        Me.DataGridTextBoxColumn2.NullText = ""
        Me.DataGridTextBoxColumn2.ReadOnly = True
        Me.DataGridTextBoxColumn2.Width = 75
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.grdData)
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(784, 246)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'grdData
        '
        Me.grdData.DataMember = ""
        Me.grdData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdData.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdData.Location = New System.Drawing.Point(3, 56)
        Me.grdData.Name = "grdData"
        Me.grdData.ReadOnly = True
        Me.grdData.Size = New System.Drawing.Size(778, 187)
        Me.grdData.TabIndex = 5
        Me.grdData.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.DataGrid = Me.grdData
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.IDField, Me.NameField, Me.Address, Me.HomePhoneField})
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = "Address"
        Me.DataGridTableStyle1.ReadOnly = True
        '
        'IDField
        '
        Me.IDField.Format = ""
        Me.IDField.FormatInfo = Nothing
        Me.IDField.MappingName = "ID"
        Me.IDField.NullText = ""
        Me.IDField.ReadOnly = True
        Me.IDField.Width = 0
        '
        'NameField
        '
        Me.NameField.Format = ""
        Me.NameField.FormatInfo = Nothing
        Me.NameField.HeaderText = "Name"
        Me.NameField.MappingName = "name"
        Me.NameField.NullText = ""
        Me.NameField.ReadOnly = True
        Me.NameField.Width = 200
        '
        'Address
        '
        Me.Address.Format = ""
        Me.Address.FormatInfo = Nothing
        Me.Address.HeaderText = "Address"
        Me.Address.MappingName = "Address"
        Me.Address.NullText = ""
        Me.Address.ReadOnly = True
        Me.Address.Width = 300
        '
        'HomePhoneField
        '
        Me.HomePhoneField.Format = ""
        Me.HomePhoneField.FormatInfo = Nothing
        Me.HomePhoneField.HeaderText = "Home Phone"
        Me.HomePhoneField.MappingName = "HomePhone"
        Me.HomePhoneField.NullText = ""
        Me.HomePhoneField.ReadOnly = True
        Me.HomePhoneField.Width = 75
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnCancel)
        Me.Panel1.Controls.Add(Me.btnSave)
        Me.Panel1.Controls.Add(Me.btnDelete)
        Me.Panel1.Controls.Add(Me.btnEdit)
        Me.Panel1.Controls.Add(Me.btnAdd)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(3, 16)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(778, 40)
        Me.Panel1.TabIndex = 6
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(553, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 32)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "&Close"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(481, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 32)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "Cance&l"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(401, 4)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(64, 32)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "&Save"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(321, 4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(64, 32)
        Me.btnDelete.TabIndex = 10
        Me.btnDelete.Text = "&Delete"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.btnEdit.Location = New System.Drawing.Point(241, 4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(64, 32)
        Me.btnEdit.TabIndex = 9
        Me.btnEdit.Text = "&Edit"
        '
        'btnAdd
        '
        Me.btnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom), System.Windows.Forms.AnchorStyles)
        Me.btnAdd.Location = New System.Drawing.Point(161, 4)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(64, 32)
        Me.btnAdd.TabIndex = 8
        Me.btnAdd.Text = "&Add"
        '
        'pnlData
        '
        Me.pnlData.Controls.Add(Me.dtBirthDay)
        Me.pnlData.Controls.Add(Me.cmbCategory)
        Me.pnlData.Controls.Add(Me.Label11)
        Me.pnlData.Controls.Add(Me.Label10)
        Me.pnlData.Controls.Add(Me.txtRemarks)
        Me.pnlData.Controls.Add(Me.Label9)
        Me.pnlData.Controls.Add(Me.txtUrl)
        Me.pnlData.Controls.Add(Me.Label8)
        Me.pnlData.Controls.Add(Me.txtWorkPhone)
        Me.pnlData.Controls.Add(Me.Label7)
        Me.pnlData.Controls.Add(Me.txtCellPhone)
        Me.pnlData.Controls.Add(Me.Label5)
        Me.pnlData.Controls.Add(Me.txtHomePhone)
        Me.pnlData.Controls.Add(Me.Label2)
        Me.pnlData.Controls.Add(Me.txtEmail)
        Me.pnlData.Controls.Add(Me.txtAddress)
        Me.pnlData.Controls.Add(Me.txtName)
        Me.pnlData.Controls.Add(Me.Label6)
        Me.pnlData.Controls.Add(Me.Label4)
        Me.pnlData.Controls.Add(Me.Label3)
        Me.pnlData.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlData.Enabled = False
        Me.pnlData.Location = New System.Drawing.Point(0, 303)
        Me.pnlData.Name = "pnlData"
        Me.pnlData.Size = New System.Drawing.Size(784, 249)
        Me.pnlData.TabIndex = 2
        Me.pnlData.TabStop = False
        '
        'dtBirthDay
        '
        Me.dtBirthDay.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtBirthDay.CustomFormat = "dd / MMM / yyyy"
        Me.dtBirthDay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtBirthDay.Location = New System.Drawing.Point(504, 56)
        Me.dtBirthDay.Name = "dtBirthDay"
        Me.dtBirthDay.ShowCheckBox = True
        Me.dtBirthDay.Size = New System.Drawing.Size(152, 20)
        Me.dtBirthDay.TabIndex = 17
        '
        'cmbCategory
        '
        Me.cmbCategory.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCategory.Location = New System.Drawing.Point(504, 32)
        Me.cmbCategory.Name = "cmbCategory"
        Me.cmbCategory.Size = New System.Drawing.Size(264, 21)
        Me.cmbCategory.TabIndex = 16
        '
        'Label11
        '
        Me.Label11.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label11.Location = New System.Drawing.Point(400, 56)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 16)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "Date of Birth"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.Location = New System.Drawing.Point(400, 32)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(80, 16)
        Me.Label10.TabIndex = 14
        Me.Label10.Text = "Category"
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.BackColor = System.Drawing.Color.White
        Me.txtRemarks.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRemarks.ForeColor = System.Drawing.Color.Maroon
        Me.txtRemarks.Location = New System.Drawing.Point(400, 104)
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(368, 136)
        Me.txtRemarks.TabIndex = 13
        Me.txtRemarks.Text = ""
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.Location = New System.Drawing.Point(400, 80)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(80, 16)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "Remarks"
        '
        'txtUrl
        '
        Me.txtUrl.BackColor = System.Drawing.Color.White
        Me.txtUrl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUrl.ForeColor = System.Drawing.Color.Maroon
        Me.txtUrl.Location = New System.Drawing.Point(112, 141)
        Me.txtUrl.Name = "txtUrl"
        Me.txtUrl.Size = New System.Drawing.Size(272, 20)
        Me.txtUrl.TabIndex = 11
        Me.txtUrl.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(16, 143)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 16)
        Me.Label8.TabIndex = 10
        Me.Label8.Text = "Home Page"
        '
        'txtWorkPhone
        '
        Me.txtWorkPhone.BackColor = System.Drawing.Color.White
        Me.txtWorkPhone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtWorkPhone.ForeColor = System.Drawing.Color.Maroon
        Me.txtWorkPhone.Location = New System.Drawing.Point(112, 216)
        Me.txtWorkPhone.Name = "txtWorkPhone"
        Me.txtWorkPhone.Size = New System.Drawing.Size(160, 20)
        Me.txtWorkPhone.TabIndex = 9
        Me.txtWorkPhone.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 218)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 16)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Work Phone"
        '
        'txtCellPhone
        '
        Me.txtCellPhone.BackColor = System.Drawing.Color.White
        Me.txtCellPhone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCellPhone.ForeColor = System.Drawing.Color.Maroon
        Me.txtCellPhone.Location = New System.Drawing.Point(112, 191)
        Me.txtCellPhone.Name = "txtCellPhone"
        Me.txtCellPhone.Size = New System.Drawing.Size(160, 20)
        Me.txtCellPhone.TabIndex = 7
        Me.txtCellPhone.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 193)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 16)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Cell Phone"
        '
        'txtHomePhone
        '
        Me.txtHomePhone.BackColor = System.Drawing.Color.White
        Me.txtHomePhone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHomePhone.ForeColor = System.Drawing.Color.Maroon
        Me.txtHomePhone.Location = New System.Drawing.Point(112, 166)
        Me.txtHomePhone.Name = "txtHomePhone"
        Me.txtHomePhone.Size = New System.Drawing.Size(160, 20)
        Me.txtHomePhone.TabIndex = 5
        Me.txtHomePhone.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Home Phone"
        '
        'txtEmail
        '
        Me.txtEmail.BackColor = System.Drawing.Color.White
        Me.txtEmail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEmail.ForeColor = System.Drawing.Color.Maroon
        Me.txtEmail.Location = New System.Drawing.Point(112, 116)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(272, 20)
        Me.txtEmail.TabIndex = 3
        Me.txtEmail.Text = ""
        '
        'txtAddress
        '
        Me.txtAddress.BackColor = System.Drawing.Color.White
        Me.txtAddress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAddress.ForeColor = System.Drawing.Color.Maroon
        Me.txtAddress.Location = New System.Drawing.Point(112, 57)
        Me.txtAddress.Multiline = True
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(272, 54)
        Me.txtAddress.TabIndex = 1
        Me.txtAddress.Text = ""
        '
        'txtName
        '
        Me.txtName.BackColor = System.Drawing.Color.White
        Me.txtName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtName.ForeColor = System.Drawing.Color.Maroon
        Me.txtName.Location = New System.Drawing.Point(112, 32)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(272, 20)
        Me.txtName.TabIndex = 0
        Me.txtName.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 118)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 16)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "Email"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 23)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Address"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Name"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtSearchbyName)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox2.Location = New System.Drawing.Point(0, 246)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(784, 57)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Search engine"
        '
        'txtSearchbyName
        '
        Me.txtSearchbyName.Location = New System.Drawing.Point(113, 22)
        Me.txtSearchbyName.Name = "txtSearchbyName"
        Me.txtSearchbyName.Size = New System.Drawing.Size(188, 20)
        Me.txtSearchbyName.TabIndex = 0
        Me.txtSearchbyName.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Search by Name:"
        '
        'frmPilgrim
        '
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.pnlData)
        Me.Name = "frmPilgrim"
        Me.Size = New System.Drawing.Size(784, 552)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.grdData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.pnlData.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        btnCancel.Enabled = False
        btnSave.Enabled = False

        btnAdd.Enabled = True
        btnEdit.Enabled = True
        btnDelete.Enabled = True
        pnlData.Enabled = False

        txtSearchbyName.Text = ""

        LoadCurrentItem()
    End Sub

    Private Sub frmPilgrim_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadGrid()
        txtSearchbyName.Focus()
    End Sub

    Private Sub LoadGrid()
        Dim data As DataTable = AddressManager.GetAddresses(" status <> " & eStatus.Deleted & " AND name like '%" & txtSearchbyName.Text & "%'")
        grdData.DataSource = data

        Dim dsn As String = System.Configuration.ConfigurationSettings.AppSettings("dsn")
        Utils.PopulateCombo(cmbCategory, dsn, "Category", "CategoryName", "CategoryId", 1)

        LoadCurrentItem()

        If grdData.CurrentRowIndex = -1 Then
            btnCancel.Enabled = False
            btnEdit.Enabled = False
            btnSave.Enabled = False
            btnDelete.Enabled = False
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        pnlData.Enabled = True
        pageAction = "ADD"

        btnSave.Enabled = True
        btnCancel.Enabled = True

        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False

        ClearData()
        txtName.Focus()
    End Sub

    Public Function ClearData()
        txtName.Text = ""
        txtAddress.Text = ""
        txtEmail.Text = ""
        txtHomePhone.Text = ""
        txtCellPhone.Text = ""
        txtWorkPhone.Text = ""
        txtUrl.Text = ""
        txtRemarks.Text = ""

        cmbCategory.SelectedIndex = 0

        dtBirthDay.Value = DateTime.Today
        dtBirthDay.Checked = False
    End Function

    Private Sub LoadCurrentItem()
        If grdData.CurrentRowIndex() = -1 Then
            Return
        End If

        Dim Id As Int32 = Int32.Parse(grdData.Item(grdData.CurrentRowIndex(), 0).ToString())
        Dim ad As Address = AddressManager.GetAddress(Id)

        If ad Is Nothing Then
            Return
        End If

        txtName.Text = ad.Name
        txtAddress.Text = ad.Address
        txtEmail.Text = ad.Email
        txtUrl.Text = ad.URL
        txtHomePhone.Text = ad.HomePhone
        txtCellPhone.Text = ad.CellPhone
        txtWorkPhone.Text = ad.WorkPhone
        txtRemarks.Text = ad.Remarks

        cmbCategory.SelectedValue = ad.CategoryId

        ' Bad way of checking if date was saved or not... to be corrected later.
        ' BBut it works for now !!!
        If ad.DateOfBirth.Date.Year = "1899" Then
            dtBirthDay.Checked = False
        Else
            dtBirthDay.Checked = True
            dtBirthDay.Value = ad.DateOfBirth
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If pageAction = "ADD" Then
            Dim ad As New Address
            ad.ID = IdManager.GetNextId("Address", "ID")
            ad.Name = Utils.ReplaceEscapeChars(txtName.Text)
            ad.Address = Utils.ReplaceEscapeChars(txtAddress.Text)
            ad.Email = Utils.ReplaceEscapeChars(txtEmail.Text)
            ad.URL = Utils.ReplaceEscapeChars(txtUrl.Text)
            ad.HomePhone = Utils.ReplaceEscapeChars(txtHomePhone.Text)
            ad.CellPhone = Utils.ReplaceEscapeChars(txtCellPhone.Text)
            ad.WorkPhone = Utils.ReplaceEscapeChars(txtWorkPhone.Text)
            ad.Remarks = Utils.ReplaceEscapeChars(txtRemarks.Text)
            ad.Status = eStatus.Approved

            ad.CategoryId = cmbCategory.SelectedValue

            If dtBirthDay.Checked Then
                ad.DateOfBirth = dtBirthDay.Value
            Else
                ad.DateOfBirth = Date.MinValue
            End If

            AddressManager.CreateAddress(ad)

            MessageBox.Show("Saved Successfully", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Dim Id As Int32 = Int32.Parse(grdData.Item(grdData.CurrentRowIndex(), 0).ToString())

            Dim ad As Address = AddressManager.GetAddress(Id)

            ad.Name = Utils.ReplaceEscapeChars(txtName.Text)
            ad.Address = Utils.ReplaceEscapeChars(txtAddress.Text)
            ad.Email = Utils.ReplaceEscapeChars(txtEmail.Text)
            ad.URL = Utils.ReplaceEscapeChars(txtUrl.Text)
            ad.HomePhone = Utils.ReplaceEscapeChars(txtHomePhone.Text)
            ad.CellPhone = Utils.ReplaceEscapeChars(txtCellPhone.Text)
            ad.WorkPhone = Utils.ReplaceEscapeChars(txtWorkPhone.Text)
            ad.Remarks = Utils.ReplaceEscapeChars(txtRemarks.Text)
            ad.Status = eStatus.Approved

            ad.CategoryId = cmbCategory.SelectedValue

            If dtBirthDay.Checked Then
                ad.DateOfBirth = dtBirthDay.Value
            Else
                ad.DateOfBirth = Date.MinValue
            End If

            AddressManager.UpdateAddress(ad)

            MessageBox.Show("Saved Successfully", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        btnCancel.Enabled = False
        btnSave.Enabled = False

        btnAdd.Enabled = True
        btnEdit.Enabled = True
        btnDelete.Enabled = True

        pnlData.Enabled = False

        txtSearchbyName.Text = ""
        LoadGrid()
    End Sub

    Private Sub txtSearchbyName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadGrid()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Visible = False
    End Sub

    Private Sub grdData_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdData.CurrentCellChanged
        grdData.Select(grdData.CurrentRowIndex)
        LoadCurrentItem()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If grdData.CurrentRowIndex = -1 Then
            MessageBox.Show("You must select a record to edit.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        pnlData.Enabled = True
        pageAction = "EDIT"

        btnAdd.Enabled = False
        btnCancel.Enabled = True
        btnEdit.Enabled = False
        btnSave.Enabled = True
        btnDelete.Enabled = False
    End Sub


    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If grdData.CurrentRowIndex = -1 Then
            MessageBox.Show("You must select a record to edit.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If (MessageBox.Show("Are you sure you want to delete current item ?", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No) Then
            Return
        End If

        Dim Id As Int32 = Int32.Parse(grdData.Item(grdData.CurrentRowIndex(), 0).ToString())
        Dim ad As Address = AddressManager.GetAddress(Id)

        If ad Is Nothing Then
            Return
        End If

        ad.Status = eStatus.Deleted

        AddressManager.UpdateAddress(ad)
        LoadGrid()

        MessageBox.Show("Successfully Deleted.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub txtSearchbyName_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchbyName.TextChanged
        LoadGrid()
    End Sub
End Class
