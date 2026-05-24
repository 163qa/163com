Attribute VB_Name = "模块1"
'Attribute VB_Name = "ModPropertyOA"
' 注意：如果通过"文件>导入文件"导入，请取消上行注释并删除本行
' 如果通过复制粘贴方式使用，请保持上行注释状态
Option Explicit

'==========================================================
' 物业公司 OA 管理系统 - 登录注册模块
' 使用方法:
'   1. 在Excel中按 Alt+F11 打开VBA编辑器
'   2. 菜单: 文件 > 导入文件 > 选择本文件
'   3. 运行宏 SetupOASystem 进行初始安装
'   4. 运行宏 LaunchSystem 启动系统
'
' 前置条件:
'   文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置
'   勾选【信任对VBA工程对象模型的访问】
'==========================================================

Private Const CT_MSFORM As Long = 3

Public gCurrentUser As String
Public gCurrentRole As String
Public gEditWorkID As String
Public gViewWorkID As String
Public gEditNoticeID As String
Public gEditApprovalID As String
Public gViewApprovalID As String
Public gEditInspectID As String
Public gEditHygieneID As String
Public gViewHygieneID As String
Public gEditTripID As String
Public gViewTripID As String
Public gEditRepairID As String
Public gViewRepairID As String
Public gEditComplaintID As String
Public gViewComplaintID As String
Public gEditFeeRow As Long
Public gEditAttendApplyID As String
Public gViewAttendApplyID As String
Public gEditScheduleRow As Long
Public gEditPersonnelRow As Long
Public gEditParkingRow As Long
Public gViewParkingID As String

' ---------- 工具函数 ----------

Public Function GetConfigProp(pName As String) As String
    On Error Resume Next
    GetConfigProp = ThisWorkbook.CustomDocumentProperties(pName).Value
    On Error GoTo 0
End Function

Public Sub SetConfigProp(pName As String, pVal As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(pName).Delete
    On Error GoTo 0
    ThisWorkbook.CustomDocumentProperties.Add pName, False, 4, pVal
End Sub

Public Function SheetExists(sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function FormExists(fName As String) As Boolean
    Dim vbc As Object
    On Error Resume Next
    Set vbc = ThisWorkbook.VBProject.VBComponents(fName)
    On Error GoTo 0
    FormExists = Not vbc Is Nothing
End Function

Public Function CheckVBAccess() As Boolean
    Dim n As Long
    On Error Resume Next
    n = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        MsgBox "请先启用VBA项目访问权限：" & vbCrLf & _
               "文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置" & vbCrLf & _
               "勾选【信任对VBA工程对象模型的访问】", vbExclamation
        CheckVBAccess = False
    Else
        CheckVBAccess = True
    End If
    On Error GoTo 0
End Function

' ---------- 安装主入口 ----------

Public Sub SetupOASystem()
    If Not CheckVBAccess() Then Exit Sub
    Application.ScreenUpdating = False
    CreateUserSheet
    CreateUserManagerForm
    CreateWorkSheet
    CreateNoticeSheet
    CreateLoginForm
    CreateRegisterForm
    CreateDashboardForm
    CreateWorkManagerForm
    CreateWorkEditForm
    CreateWorkViewForm
    CreateNoticesForm
    CreateNoticeEditForm
    CreateApprovalSheet
    CreateApprovalsForm
    CreateApprovalEditForm
    CreateApprovalViewForm
    CreateInspectSheet
    CreateInspectManagerForm
    CreateInspectEditForm
    CreateHygieneSheet
    CreateHygieneManagerForm
    CreateHygieneEditForm
    CreateHygieneViewForm
    CreateTripSheet
    CreateTripManagerForm
    CreateTripEditForm
    CreateTripViewForm
    CreateAttendSheet
    CreateAttendForm
    CreateRepairSheet
    CreateRepairsForm
    CreateRepairEditForm
    CreateRepairViewForm
    CreateComplaintSheet
    CreateComplaintsForm
    CreateComplaintEditForm
    CreateFeeSheet
    CreateFeeManagerForm
    CreateFeeEditForm
    CreateParkingSheet
    CreateParkingManagerForm
    CreateParkingEditForm
    CreateParkingViewForm
    CreateAttendApplySheet
    CreateScheduleSheet
    CreatePersonnelSheet
    CreateAttendStatsSheet
    CreateHRMainForm
    CreateAttendApplyEditForm
    CreateAttendApplyViewForm
    CreateScheduleEditForm
    CreatePersonnelEditForm
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox "物业OA系统安装完成！" & vbCrLf & _
           "请运行宏【LaunchSystem】启动系统。" & vbCrLf & _
           "默认管理员: admin / admin123", vbInformation
End Sub

' ---------- 创建用户管理表 ----------

Private Sub CreateUserSheet()
    Dim ws As Worksheet
    If SheetExists("用户管理") Then
        Set ws = ThisWorkbook.Sheets("用户管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "用户管理"
    End If
    With ws
        .Range("A1").Value = "姓名"
        .Range("B1").Value = "账号"
        .Range("C1").Value = "密码"
        .Range("D1").Value = "角色"
        .Range("E1").Value = "注册时间"
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(70, 130, 180)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
        If .Range("A2").Value = "" Then
            .Range("A2").Value = "系统管理员"
            .Range("B2").Value = "admin"
            .Range("C2").Value = "admin123"
            .Range("D2").Value = "管理员"
            .Range("E2").Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
        End If
        .Columns("A:E").AutoFit
    End With
End Sub

' ---------- 创建用户管理窗体 ----------

Private Sub CreateUserManagerForm()
    Dim oldForm As String
    oldForm = GetConfigProp("UserManagerFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "UserManagerFormName", actualName

    vbc.Properties("Caption") = "用户管理"
    vbc.Properties("Width") = 470
    vbc.Properties("Height") = 320
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 用户列表
    Set ctl = dsg.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstUsers"
    ctl.Left = 12: ctl.Top = 12: ctl.Width = 200: ctl.Height = 300

    ' 姓名
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label1": ctl.Caption = "姓名："
    ctl.Left = 224: ctl.Top = 24: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtName"
    ctl.Left = 294: ctl.Top = 24: ctl.Width = 150: ctl.Height = 24

    ' 账号
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label2": ctl.Caption = "账号："
    ctl.Left = 224: ctl.Top = 54: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtAccount"
    ctl.Left = 294: ctl.Top = 54: ctl.Width = 150: ctl.Height = 24

    ' 密码
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label3": ctl.Caption = "密码："
    ctl.Left = 224: ctl.Top = 84: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtPassword"
    ctl.Left = 294: ctl.Top = 84: ctl.Width = 150: ctl.Height = 24
    ctl.PasswordChar = "*"

    ' 角色
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label4": ctl.Caption = "角色："
    ctl.Left = 224: ctl.Top = 114: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboRole"
    ctl.Left = 294: ctl.Top = 114: ctl.Width = 150: ctl.Height = 24
    ctl.Style = 2

    ' 按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdAddUser": ctl.Caption = "新增"
    ctl.Left = 224: ctl.Top = 160: ctl.Width = 80: ctl.Height = 28
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdEditUser": ctl.Caption = "修改"
    ctl.Left = 314: ctl.Top = 160: ctl.Width = 80: ctl.Height = 28
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDeleteUser": ctl.Caption = "删除用户"
    ctl.Left = 224: ctl.Top = 200: ctl.Width = 170: ctl.Height = 28
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 320: ctl.Top = 250: ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf

    ' --- Initialize ---
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    LoadRoles" & vbCrLf
    c = c & "    LoadUserList" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf

    ' --- LoadRoles: 加载角色下拉（预设 + 工作表中已有角色） ---
    c = c & "Private Sub LoadRoles()" & vbCrLf
    c = c & "    cboRole.Clear" & vbCrLf
    c = c & "    Dim dict As Object" & vbCrLf
    c = c & "    Set dict = CreateObject(" & q & "Scripting.Dictionary" & q & ")" & vbCrLf
    c = c & "    dict.Add " & q & "普通员工" & q & ", 1" & vbCrLf
    c = c & "    dict.Add " & q & "部门主管" & q & ", 1" & vbCrLf
    c = c & "    dict.Add " & q & "管理员" & q & ", 1" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim rl As String" & vbCrLf
    c = c & "        rl = Trim(CStr(ws.Cells(i, 4).Value))" & vbCrLf
    c = c & "        If rl <> " & q & q & " And Not dict.Exists(rl) Then" & vbCrLf
    c = c & "            dict.Add rl, 1" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    Dim k As Variant" & vbCrLf
    c = c & "    For Each k In dict.Keys" & vbCrLf
    c = c & "        cboRole.AddItem CStr(k)" & vbCrLf
    c = c & "    Next k" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf

    ' --- LoadUserList: 加载所有已注册用户到左侧列表 ---
    c = c & "Private Sub LoadUserList()" & vbCrLf
    c = c & "    lstUsers.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim uName As String, uAcct As String, uRole As String" & vbCrLf
    c = c & "        uName = CStr(ws.Cells(i, 1).Value)" & vbCrLf
    c = c & "        uAcct = CStr(ws.Cells(i, 2).Value)" & vbCrLf
    c = c & "        uRole = CStr(ws.Cells(i, 4).Value)" & vbCrLf
    c = c & "        lstUsers.AddItem uName & " & q & " (" & q & " & uAcct & " & q & ") [" & q & " & uRole & " & q & "]" & q & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf

    ' --- lstUsers_Click: 选中用户后填充右侧字段 ---
    c = c & "Private Sub lstUsers_Click()" & vbCrLf
    c = c & "    If lstUsers.ListIndex < 0 Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim r As Long" & vbCrLf
    c = c & "    r = lstUsers.ListIndex + 2" & vbCrLf
    c = c & "    txtName.Text = CStr(ws.Cells(r, 1).Value)" & vbCrLf
    c = c & "    txtAccount.Text = CStr(ws.Cells(r, 2).Value)" & vbCrLf
    c = c & "    txtPassword.Text = CStr(ws.Cells(r, 3).Value)" & vbCrLf
    c = c & "    Dim rl As String" & vbCrLf
    c = c & "    rl = Trim(CStr(ws.Cells(r, 4).Value))" & vbCrLf
    c = c & "    cboRole.ListIndex = -1" & vbCrLf
    c = c & "    Dim j As Long" & vbCrLf
    c = c & "    Dim matched As Boolean" & vbCrLf
    c = c & "    matched = False" & vbCrLf
    c = c & "    For j = 0 To cboRole.ListCount - 1" & vbCrLf
    c = c & "        If CStr(cboRole.List(j, 0)) = rl Then" & vbCrLf
    c = c & "            cboRole.ListIndex = j" & vbCrLf
    c = c & "            matched = True" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next j" & vbCrLf
    c = c & "    If Not matched And rl <> " & q & q & " Then" & vbCrLf
    c = c & "        cboRole.AddItem rl" & vbCrLf
    c = c & "        cboRole.ListIndex = cboRole.ListCount - 1" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf

    ' --- cmdAddUser_Click ---
    c = c & "Private Sub cmdAddUser_Click()" & vbCrLf
    c = c & "    If Trim(txtName.Text) = " & q & q & " Or Trim(txtAccount.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "姓名和账号不能为空！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(txtPassword.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "密码不能为空！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If cboRole.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择角色！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    ' 检查账号是否重复
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 2).Value) = Trim(txtAccount.Text) Then" & vbCrLf
    c = c & "            MsgBox " & q & "该账号已存在！" & q & ", vbExclamation" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    Dim nr As Long" & vbCrLf
    c = c & "    nr = lr + 1" & vbCrLf
    c = c & "    ws.Cells(nr, 1).Value = Trim(txtName.Text)" & vbCrLf
    c = c & "    ws.Cells(nr, 2).Value = Trim(txtAccount.Text)" & vbCrLf
    c = c & "    ws.Cells(nr, 3).Value = Trim(txtPassword.Text)" & vbCrLf
    c = c & "    ws.Cells(nr, 4).Value = cboRole.Text" & vbCrLf
    c = c & "    ws.Cells(nr, 5).Value = Format(Now, " & q & "yyyy-mm-dd hh:mm:ss" & q & ")" & vbCrLf
    c = c & "    MsgBox " & q & "用户新增成功！" & q & ", vbInformation" & vbCrLf
    c = c & "    ClearFields" & vbCrLf
    c = c & "    LoadRoles" & vbCrLf
    c = c & "    LoadUserList" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf

    ' --- cmdEditUser_Click ---
    c = c & "Private Sub cmdEditUser_Click()" & vbCrLf
    c = c & "    If lstUsers.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择要修改的用户！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(txtName.Text) = " & q & q & " Or Trim(txtAccount.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "姓名和账号不能为空！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If cboRole.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择角色！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim r As Long" & vbCrLf
    c = c & "    r = lstUsers.ListIndex + 2" & vbCrLf
    ' 检查修改后的账号是否与其他用户重复
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If i <> r And CStr(ws.Cells(i, 2).Value) = Trim(txtAccount.Text) Then" & vbCrLf
    c = c & "            MsgBox " & q & "该账号已被其他用户使用！" & q & ", vbExclamation" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    ws.Cells(r, 1).Value = Trim(txtName.Text)" & vbCrLf
    c = c & "    ws.Cells(r, 2).Value = Trim(txtAccount.Text)" & vbCrLf
    c = c & "    If Trim(txtPassword.Text) <> " & q & q & " Then" & vbCrLf
    c = c & "        ws.Cells(r, 3).Value = Trim(txtPassword.Text)" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    ws.Cells(r, 4).Value = cboRole.Text" & vbCrLf
    c = c & "    MsgBox " & q & "用户信息已修改！" & q & ", vbInformation" & vbCrLf
    c = c & "    LoadRoles" & vbCrLf
    c = c & "    LoadUserList" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf

    ' --- cmdDeleteUser_Click ---
    c = c & "Private Sub cmdDeleteUser_Click()" & vbCrLf
    c = c & "    If lstUsers.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择要删除的用户！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim r As Long" & vbCrLf
    c = c & "    r = lstUsers.ListIndex + 2" & vbCrLf
    c = c & "    Dim uAcct As String" & vbCrLf
    c = c & "    uAcct = CStr(ws.Cells(r, 2).Value)" & vbCrLf
    c = c & "    If uAcct = " & q & "admin" & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "不能删除默认管理员账号！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If uAcct = gCurrentUser Then" & vbCrLf
    c = c & "        MsgBox " & q & "不能删除当前登录用户！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim ans As VbMsgBoxResult" & vbCrLf
    c = c & "    ans = MsgBox(" & q & "确定要删除用户【" & q & " & ws.Cells(r, 1).Value & " & q & "】？此操作不可恢复！" & q & ", vbYesNo + vbExclamation)" & vbCrLf
    c = c & "    If ans = vbNo Then Exit Sub" & vbCrLf
    c = c & "    ws.Rows(r).Delete" & vbCrLf
    c = c & "    MsgBox " & q & "用户已删除！" & q & ", vbInformation" & vbCrLf
    c = c & "    ClearFields" & vbCrLf
    c = c & "    LoadRoles" & vbCrLf
    c = c & "    LoadUserList" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf

    ' --- ClearFields ---
    c = c & "Private Sub ClearFields()" & vbCrLf
    c = c & "    txtName.Text = " & q & q & vbCrLf
    c = c & "    txtAccount.Text = " & q & q & vbCrLf
    c = c & "    txtPassword.Text = " & q & q & vbCrLf
    c = c & "    cboRole.ListIndex = -1" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf

    ' --- cmdClose_Click ---
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 显示用户管理窗体 ----------

Public Sub ShowUserManagerForm()
    Dim fName As String
    fName = GetConfigProp("UserManagerFormName")
    If fName = "" Or Not FormExists("UserManagerFormName") Then
        MsgBox "用户管理窗体不存在！请先运行SetupOASystem。", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("UserManagerFormName")
    frm.Show
End Sub

' ---------- 创建工作管理表 ----------

Private Sub CreateWorkSheet()
    Dim ws As Worksheet
    If SheetExists("工作管理") Then
        Set ws = ThisWorkbook.Sheets("工作管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "工作管理"
    End If
    With ws
        .Range("A1").Value = "工作ID"
        .Range("B1").Value = "工作标题"
        .Range("C1").Value = "工作描述"
        .Range("D1").Value = "工作类型"
        .Range("E1").Value = "优先级"
        .Range("F1").Value = "负责人"
        .Range("G1").Value = "创建人"
        .Range("H1").Value = "状态"
        .Range("I1").Value = "创建时间"
        .Range("J1").Value = "截止时间"
        .Range("K1").Value = "完成时间"
        .Range("A1:K1").Font.Bold = True
        .Range("A1:K1").Interior.Color = RGB(70, 130, 180)
        .Range("A1:K1").Font.Color = RGB(255, 255, 255)
        .Columns("A:K").AutoFit
    End With
End Sub

' ---------- 创建信息管理表 ----------

Private Sub CreateNoticeSheet()
    Dim ws As Worksheet
    If SheetExists("信息管理") Then
        Set ws = ThisWorkbook.Sheets("信息管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "信息管理"
    End If
    With ws
        .Range("A1").Value = "公告ID"
        .Range("B1").Value = "标题"
        .Range("C1").Value = "内容"
        .Range("D1").Value = "类型"
        .Range("E1").Value = "发布部门"
        .Range("F1").Value = "发布人"
        .Range("G1").Value = "对象"
        .Range("H1").Value = "是否置顶"
        .Range("I1").Value = "紧急程度"
        .Range("J1").Value = "生效时间"
        .Range("K1").Value = "截止时间"
        .Range("L1").Value = "发布时间"
        .Range("A1:L1").Font.Bold = True
        .Range("A1:L1").Interior.Color = RGB(70, 130, 180)
        .Range("A1:L1").Font.Color = RGB(255, 255, 255)
        .Columns("A:L").AutoFit
    End With
End Sub

Private Sub CreateLoginForm()
    Dim oldForm As String
    oldForm = GetConfigProp("LoginFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "LoginFormName", actualName

    ' 设置窗体属性
    vbc.Properties("Caption") = "物业OA系统"
    vbc.Properties("Width") = 360
    vbc.Properties("Height") = 240
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer

    ' Label1 - 标题
    Dim ctl As Object
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label1"
    ctl.Caption = "物业 OA 系统 - 登录"
    ctl.Left = 48: ctl.Top = 36
    ctl.Width = 200: ctl.Height = 24
    ctl.Font.Size = 14
    ctl.Font.Bold = True

    ' Label2 - 用户名标签
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label2"
    ctl.Caption = "用户名："
    ctl.Left = 48: ctl.Top = 72
    ctl.Width = 60: ctl.Height = 18

    ' txtUser
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtUser"
    ctl.Left = 120: ctl.Top = 72
    ctl.Width = 150: ctl.Height = 24

    ' Label3 - 密码标签
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label3"
    ctl.Caption = "密码："
    ctl.Left = 48: ctl.Top = 108
    ctl.Width = 60: ctl.Height = 18

    ' txtPass
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtPass"
    ctl.Left = 120: ctl.Top = 108
    ctl.Width = 150: ctl.Height = 24
    ctl.PasswordChar = "*"

    ' cmdLogin
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdLogin"
    ctl.Caption = "登录"
    ctl.Left = 80: ctl.Top = 156
    ctl.Width = 80: ctl.Height = 28

    ' cmdExit
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdExit"
    ctl.Caption = "退出"
    ctl.Left = 180: ctl.Top = 156
    ctl.Width = 80: ctl.Height = 28

    ' lblRegister - 注册链接
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRegister"
    ctl.Caption = "没有账号？点击注册"
    ctl.Left = 110: ctl.Top = 195
    ctl.Width = 130: ctl.Height = 16
    ctl.ForeColor = RGB(0, 0, 255)
    ctl.Font.Underline = True
    ctl.Font.Size = 9

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "Private loginAttempts As Integer" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    loginAttempts = 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdLogin_Click()" & vbCrLf
    c = c & "    Dim u As String, p As String" & vbCrLf
    c = c & "    u = Trim(txtUser.Text)" & vbCrLf
    c = c & "    p = Trim(txtPass.Text)" & vbCrLf
    c = c & "    If u = " & q & q & " Or p = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请输入用户名和密码！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If AuthenticateUser(u, p) Then" & vbCrLf
    c = c & "        Me.Hide" & vbCrLf
    c = c & "        Unload Me" & vbCrLf
    c = c & "        ShowDashForm" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        loginAttempts = loginAttempts + 1" & vbCrLf
    c = c & "        If loginAttempts >= 3 Then" & vbCrLf
    c = c & "            MsgBox " & q & "登录失败次数过多，系统退出！" & q & ", vbCritical" & vbCrLf
    c = c & "            Unload Me" & vbCrLf
    c = c & "        Else" & vbCrLf
    c = c & "            Dim remain As Integer" & vbCrLf
    c = c & "            remain = 3 - loginAttempts" & vbCrLf
    c = c & "            MsgBox " & q & "用户名或密码错误！剩余" & q
    c = c & " & remain & " & q & "次机会" & q & ", vbWarning" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdExit_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub lblRegister_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowRegForm" & vbCrLf
    c = c & "    Dim fName As String" & vbCrLf
    c = c & "    fName = GetConfigProp(" & q & "LoginFormName" & q & ")" & vbCrLf
    c = c & "    Dim frm As Object" & vbCrLf
    c = c & "    Set frm = VBA.UserForms.Add(fName)" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "    frm.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub txtUser_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    c = c & "    If KeyCode = 13 Then txtPass.SetFocus" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub txtPass_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    c = c & "    If KeyCode = 13 Then cmdLogin_Click" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建注册窗体 (占位) ----------

Private Sub CreateRegisterForm()
    Dim oldForm As String
    oldForm = GetConfigProp("RegisterFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "RegisterFormName", actualName

    vbc.Properties("Caption") = "用户注册"
    vbc.Properties("Width") = 280
    vbc.Properties("Height") = 320
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' lblName
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblName"
    ctl.Caption = "姓名："
    ctl.Left = 20: ctl.Top = 20
    ctl.Width = 70: ctl.Height = 18

    ' txtName
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtName"
    ctl.Left = 100: ctl.Top = 20
    ctl.Width = 140: ctl.Height = 20

    ' lblAccount
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAccount"
    ctl.Caption = "账号："
    ctl.Left = 20: ctl.Top = 60
    ctl.Width = 70: ctl.Height = 18

    ' txtAccount
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtAccount"
    ctl.Left = 100: ctl.Top = 60
    ctl.Width = 140: ctl.Height = 20

    ' lblPwd
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPwd"
    ctl.Caption = "密码："
    ctl.Left = 20: ctl.Top = 95
    ctl.Width = 70: ctl.Height = 18

    ' txtPwd
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtPwd"
    ctl.Left = 100: ctl.Top = 95
    ctl.Width = 140: ctl.Height = 20
    ctl.PasswordChar = "*"

    ' lblConfirmPwd
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblConfirmPwd"
    ctl.Caption = "确认密码："
    ctl.Left = 20: ctl.Top = 130
    ctl.Width = 70: ctl.Height = 18

    ' txtConfirmPwd
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtConfirmPwd"
    ctl.Left = 100: ctl.Top = 130
    ctl.Width = 140: ctl.Height = 20
    ctl.PasswordChar = "*"

    ' lblRole
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRole"
    ctl.Caption = "角色："
    ctl.Left = 20: ctl.Top = 165
    ctl.Width = 70: ctl.Height = 18

    ' cboRole
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboRole"
    ctl.Left = 100: ctl.Top = 165
    ctl.Width = 140: ctl.Height = 20
    ctl.Style = 2

    ' cmdSave
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave"
    ctl.Caption = "保存"
    ctl.Left = 60: ctl.Top = 220
    ctl.Width = 60: ctl.Height = 24

    ' cmdCancel
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel"
    ctl.Caption = "取消"
    ctl.Left = 160: ctl.Top = 220
    ctl.Width = 60: ctl.Height = 24

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    With cboRole" & vbCrLf
    c = c & "        .AddItem " & q & "管理员" & q & vbCrLf
    c = c & "        .AddItem " & q & "部门主管" & q & vbCrLf
    c = c & "        .AddItem " & q & "行政部" & q & vbCrLf
    c = c & "        .AddItem " & q & "财务部" & q & vbCrLf
    c = c & "        .AddItem " & q & "销售部" & q & vbCrLf
    c = c & "        .AddItem " & q & "工程部" & q & vbCrLf
    c = c & "        .AddItem " & q & "客服部" & q & vbCrLf
    c = c & "        .AddItem " & q & "保安部" & q & vbCrLf
    c = c & "        .AddItem " & q & "保洁部" & q & vbCrLf
    c = c & "        .ListIndex = -1" & vbCrLf
    c = c & "    End With" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    Dim sName As String, sAcct As String" & vbCrLf
    c = c & "    Dim sPwd As String, sConf As String, sRole As String" & vbCrLf
    c = c & "    sName = Trim(txtName.Text)" & vbCrLf
    c = c & "    sAcct = Trim(txtAccount.Text)" & vbCrLf
    c = c & "    sPwd = Trim(txtPwd.Text)" & vbCrLf
    c = c & "    sConf = Trim(txtConfirmPwd.Text)" & vbCrLf
    c = c & "    If cboRole.ListIndex >= 0 Then" & vbCrLf
    c = c & "        sRole = cboRole.Text" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        sRole = " & q & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If sName = " & q & q & " Or sAcct = " & q & q
    c = c & " Or sPwd = " & q & q & " Or sRole = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写所有必填项！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If sPwd <> sConf Then" & vbCrLf
    c = c & "        MsgBox " & q & "两次输入的密码不一致！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Len(sPwd) < 6 Then" & vbCrLf
    c = c & "        MsgBox " & q & "密码长度不能少于6位！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If AccountExists(sAcct) Then" & vbCrLf
    c = c & "        MsgBox " & q & "该账号已存在！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If sRole = " & q & "管理员" & q & " And AdminExists() Then" & vbCrLf
    c = c & "        MsgBox " & q & "系统仅允许一个管理员账号！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    RegisterNewUser sName, sAcct, sPwd, sRole" & vbCrLf
    c = c & "    MsgBox " & q & "注册成功！请返回登录。" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建仪表盘窗体 (占位) ----------

Private Sub CreateDashboardForm()
    Dim oldForm As String
    oldForm = GetConfigProp("DashboardFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "DashboardFormName", actualName

    vbc.Properties("Caption") = "物业OA系统 - 仪表盘"
    vbc.Properties("Width") = 1000
    vbc.Properties("Height") = 600
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 欢迎标签
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblWelcome"
    ctl.Caption = "欢迎，"
    ctl.Left = 24: ctl.Top = 24
    ctl.Width = 300: ctl.Height = 24
    ctl.Font.Size = 14
    ctl.Font.Bold = True

    ' 角色权限标签
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRoleInfo"
    ctl.Caption = ""
    ctl.Left = 24: ctl.Top = 52
    ctl.Width = 500: ctl.Height = 18
    ctl.ForeColor = RGB(70, 130, 180)
    ctl.Font.Size = 9

    ' ===== 区域1: 统计数据 =====
    Set ctl = dsg.Controls.Add("Forms.Frame.1")
    ctl.Name = "fraStats"
    ctl.Caption = "统计数据"
    ctl.Left = 24: ctl.Top = 78
    ctl.Width = 940: ctl.Height = 70
    ctl.Font.Bold = True

    Dim statsFrame As Object
    Set statsFrame = ctl

    Call AddDashStatLabel(dsg, statsFrame, "lblStatUsers", "注册用户: 0", 20, 22, 140, 30)
    Call AddDashStatLabel(dsg, statsFrame, "lblStatPending", "待办工作: 0", 170, 22, 140, 30)
    Call AddDashStatLabel(dsg, statsFrame, "lblStatDoing", "办理中: 0", 320, 22, 140, 30)
    Call AddDashStatLabel(dsg, statsFrame, "lblStatDone", "已完成: 0", 470, 22, 140, 30)
    Call AddDashStatLabel(dsg, statsFrame, "lblStatApproval", "待审批: 0", 620, 22, 140, 30)
    Call AddDashStatLabel(dsg, statsFrame, "lblStatRepair", "待报修: 0", 770, 22, 140, 30)

    ' ===== 区域2: 业务模块按钮 =====
    Set ctl = dsg.Controls.Add("Forms.Frame.1")
    ctl.Name = "fraModules"
    ctl.Caption = "业务模块"
    ctl.Left = 24: ctl.Top = 156
    ctl.Width = 940: ctl.Height = 200
    ctl.Font.Bold = True

    Dim modFrame As Object
    Set modFrame = ctl

    ' 13个模块按钮, 5列 x 3行布局
    Dim btnNames(1 To 13) As String
    Dim btnCaptions(1 To 13) As String
    btnNames(1) = "cmdTrip": btnCaptions(1) = "个人行程"
    btnNames(2) = "cmdAttend": btnCaptions(2) = "个人考勤"
    btnNames(3) = "cmdInfo": btnCaptions(3) = "信息管理"
    btnNames(4) = "cmdWork": btnCaptions(4) = "工作管理"
    btnNames(5) = "cmdInspect": btnCaptions(5) = "巡检管理"
    btnNames(6) = "cmdHygiene": btnCaptions(6) = "卫生管理"
    btnNames(7) = "cmdFee": btnCaptions(7) = "费用管理"
    btnNames(8) = "cmdApproval": btnCaptions(8) = "审批管理"
    btnNames(9) = "cmdRepair": btnCaptions(9) = "报修管理"
    btnNames(10) = "cmdComplaint": btnCaptions(10) = "投诉建议"
    btnNames(11) = "cmdParking": btnCaptions(11) = "停车管理"
    btnNames(12) = "cmdHR": btnCaptions(12) = "人力资源"
    btnNames(13) = "cmdUserMgr": btnCaptions(13) = "用户管理"

    Dim bIdx As Long
    Dim bRow As Long, bCol As Long
    Dim bLeft As Long, bTop As Long
    Dim bW As Long, bH As Long
    bW = 140: bH = 40

    For bIdx = 1 To 13
        bRow = (bIdx - 1) \ 5
        bCol = (bIdx - 1) Mod 5
        bLeft = 30 + bCol * (bW + 24)
        bTop = 26 + bRow * (bH + 14)
        Set ctl = modFrame.Controls.Add("Forms.CommandButton.1")
        ctl.Name = btnNames(bIdx)
        ctl.Caption = btnCaptions(bIdx)
        ctl.Left = bLeft: ctl.Top = bTop
        ctl.Width = bW: ctl.Height = bH
        ctl.Font.Size = 10
    Next bIdx

    ' ===== 区域3: 待审批/待处理事项 =====
    Set ctl = dsg.Controls.Add("Forms.Frame.1")
    ctl.Name = "fraPending"
    ctl.Caption = "待办事项"
    ctl.Left = 24: ctl.Top = 366
    ctl.Width = 450: ctl.Height = 170
    ctl.Font.Bold = True

    Dim pendFrame As Object
    Set pendFrame = ctl

    Set ctl = pendFrame.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPendTitle"
    ctl.Caption = "待处理工作："
    ctl.Left = 12: ctl.Top = 20
    ctl.Width = 120: ctl.Height = 16
    ctl.Font.Bold = True

    Set ctl = pendFrame.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstPendingItems"
    ctl.Left = 12: ctl.Top = 40
    ctl.Width = 420: ctl.Height = 110

    ' 待审批区域
    Set ctl = dsg.Controls.Add("Forms.Frame.1")
    ctl.Name = "fraApproval"
    ctl.Caption = "待审批事项"
    ctl.Left = 490: ctl.Top = 366
    ctl.Width = 474: ctl.Height = 170
    ctl.Font.Bold = True

    Dim apprvFrame As Object
    Set apprvFrame = ctl

    Set ctl = apprvFrame.Controls.Add("Forms.Label.1")
    ctl.Name = "lblApprvTitle"
    ctl.Caption = "待审批列表："
    ctl.Left = 12: ctl.Top = 20
    ctl.Width = 120: ctl.Height = 16
    ctl.Font.Bold = True

    Set ctl = apprvFrame.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstApprovalItems"
    ctl.Left = 12: ctl.Top = 40
    ctl.Width = 444: ctl.Height = 110

    ' 底部注销按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdLogout"
    ctl.Caption = "注销登录"
    ctl.Left = 880: ctl.Top = 546
    ctl.Width = 90: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    lblWelcome.Caption = " & q & "欢迎，" & q & " & gCurrentUser" & vbCrLf
    c = c & "    Dim isHigh As Boolean" & vbCrLf
    c = c & "    isHigh = IsHighPrivilege(gCurrentRole)" & vbCrLf
    c = c & "    If isHigh Then" & vbCrLf
    c = c & "        lblRoleInfo.Caption = " & q & "角色: " & q
    c = c & " & gCurrentRole & " & q & " | 权限: 最高权限(增删改查)" & q & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        lblRoleInfo.Caption = " & q & "角色: " & q
    c = c & " & gCurrentRole & " & q & " | 权限: 普通用户(仅查看)" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    cmdUserMgr.Enabled = isHigh" & vbCrLf
    c = c & "    LoadStats" & vbCrLf
    c = c & "    LoadPendingItems" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadStats()" & vbCrLf
    c = c & "    Dim wsU As Worksheet, wsW As Worksheet" & vbCrLf
    c = c & "    Dim uCnt As Long, pCnt As Long, dCnt As Long, fCnt As Long" & vbCrLf
    c = c & "    Set wsU = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    uCnt = Application.Max(wsU.Cells(wsU.Rows.Count, 1).End(xlUp).Row - 1, 0)" & vbCrLf
    c = c & "    lblStatUsers.Caption = " & q & "注册用户: " & q & " & uCnt" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set wsW = ThisWorkbook.Sheets(" & q & "工作管理" & q & ")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If wsW Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = wsW.Cells(wsW.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    pCnt = 0: dCnt = 0: fCnt = 0" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Select Case wsW.Cells(i, 8).Value" & vbCrLf
    c = c & "            Case " & q & "待办" & q & ": pCnt = pCnt + 1" & vbCrLf
    c = c & "            Case " & q & "办理中" & q & ": dCnt = dCnt + 1" & vbCrLf
    c = c & "            Case " & q & "已完成" & q & ": fCnt = fCnt + 1" & vbCrLf
    c = c & "        End Select" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    lblStatPending.Caption = " & q & "待办工作: " & q & " & pCnt" & vbCrLf
    c = c & "    lblStatDoing.Caption = " & q & "办理中: " & q & " & dCnt" & vbCrLf
    c = c & "    lblStatDone.Caption = " & q & "已完成: " & q & " & fCnt" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadPendingItems()" & vbCrLf
    c = c & "    lstPendingItems.Clear" & vbCrLf
    c = c & "    lstApprovalItems.Clear" & vbCrLf
    c = c & "    Dim wsW As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set wsW = ThisWorkbook.Sheets(" & q & "工作管理" & q & ")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If wsW Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    Dim isHigh As Boolean" & vbCrLf
    c = c & "    isHigh = IsHighPrivilege(gCurrentRole)" & vbCrLf
    c = c & "    lr = wsW.Cells(wsW.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim st As String, assignee As String" & vbCrLf
    c = c & "        st = wsW.Cells(i, 8).Value" & vbCrLf
    c = c & "        assignee = wsW.Cells(i, 6).Value" & vbCrLf
    c = c & "        If st = " & q & "待办" & q & " Or st = " & q & "办理中" & q & " Then" & vbCrLf
    c = c & "            If isHigh Or assignee = gCurrentUser Then" & vbCrLf
    c = c & "                Dim info As String" & vbCrLf
    c = c & "                info = wsW.Cells(i, 2).Value & " & q & " | " & q
    c = c & " & assignee & " & q & " | " & q & " & st" & vbCrLf
    c = c & "                lstPendingItems.AddItem info" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdWork_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowWorkManagerForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdUserMgr_Click()" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "您没有用户管理权限！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowUserManagerForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdTrip_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowTripManagerForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdAttend_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowAttendForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdInfo_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowNoticesForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdInspect_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowInspectManagerForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdHygiene_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowHygieneManagerForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdFee_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowFeeManagerForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdApproval_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowApprovalsForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdRepair_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowRepairsForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdComplaint_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowComplaintsForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdParking_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowParkingManagerForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub cmdHR_Click()" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "    ShowHRMainForm" & vbCrLf
    c = c & "    Me.Show" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdLogout_Click()" & vbCrLf
    c = c & "    gCurrentUser = " & q & q & vbCrLf
    c = c & "    gCurrentRole = " & q & q & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "    LaunchSystem" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    c = c & "    If CloseMode = 0 Then" & vbCrLf
    c = c & "        gCurrentUser = " & q & q & vbCrLf
    c = c & "        gCurrentRole = " & q & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

Private Sub AddDashStatLabel(dsg As Object, frm As Object, sName As String, sCap As String, sL As Long, sT As Long, sW As Long, sH As Long)
    Dim ctl As Object
    Set ctl = frm.Controls.Add("Forms.Label.1")
    ctl.Name = sName
    ctl.Caption = sCap
    ctl.Left = sL: ctl.Top = sT
    ctl.Width = sW: ctl.Height = sH
    ctl.Font.Size = 11
    ctl.Font.Bold = True
    ctl.ForeColor = RGB(50, 50, 120)
End Sub

' ---------- 创建工作管理窗体 ----------

Private Sub CreateWorkManagerForm()
    Dim oldForm As String
    oldForm = GetConfigProp("WorkManagerFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "WorkManagerFormName", actualName

    vbc.Properties("Caption") = "工作管理"
    vbc.Properties("Width") = 560
    vbc.Properties("Height") = 480
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 切换卡片: 我的工作 / 所有工作
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdMyWork"
    ctl.Caption = "我的工作"
    ctl.Left = 12: ctl.Top = 6
    ctl.Width = 90: ctl.Height = 26
    ctl.BackColor = RGB(70, 130, 180)
    ctl.ForeColor = RGB(255, 255, 255)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdAllWork"
    ctl.Caption = "所有工作"
    ctl.Left = 110: ctl.Top = 6
    ctl.Width = 90: ctl.Height = 26

    ' MultiPage
    Set ctl = dsg.Controls.Add("Forms.MultiPage.1")
    ctl.Name = "MultiPage1"
    ctl.Left = 12: ctl.Top = 38
    ctl.Width = 530: ctl.Height = 370

    Dim mp As Object
    Set mp = ctl
    mp.Pages(0).Caption = "待办工作"
    mp.Pages.Add
    mp.Pages(1).Caption = "办理中"
    mp.Pages.Add
    mp.Pages(2).Caption = "已完成"

    ' Page 0 - 待办列表
    Dim pg0 As Object
    Set pg0 = mp.Pages(0)

    Set ctl = pg0.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstPending"
    ctl.Left = 12: ctl.Top = 12
    ctl.Width = 498: ctl.Height = 230
    ctl.ColumnCount = 6
    ctl.ColumnWidths = "40;120;80;60;80;80"

    Set ctl = pg0.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPendHdr"
    ctl.Caption = "ID | 标题 | 类型 | 优先级 | 负责人 | 截止时间"
    ctl.Left = 12: ctl.Top = 244
    ctl.Width = 498: ctl.Height = 14
    ctl.Font.Size = 8
    ctl.ForeColor = RGB(120, 120, 120)

    Set ctl = pg0.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNewTask"
    ctl.Caption = "新建工作"
    ctl.Left = 12: ctl.Top = 264
    ctl.Width = 80: ctl.Height = 28

    Set ctl = pg0.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdViewTask"
    ctl.Caption = "查看"
    ctl.Left = 102: ctl.Top = 264
    ctl.Width = 70: ctl.Height = 28

    Set ctl = pg0.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdEditTask"
    ctl.Caption = "编辑"
    ctl.Left = 182: ctl.Top = 264
    ctl.Width = 70: ctl.Height = 28

    Set ctl = pg0.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDeleteTask"
    ctl.Caption = "删除"
    ctl.Left = 262: ctl.Top = 264
    ctl.Width = 70: ctl.Height = 28

    ' Page 1 - 办理中列表
    Dim pg1 As Object
    Set pg1 = mp.Pages(1)

    Set ctl = pg1.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstDoing"
    ctl.Left = 12: ctl.Top = 12
    ctl.Width = 498: ctl.Height = 260
    ctl.ColumnCount = 6
    ctl.ColumnWidths = "40;120;80;60;80;80"

    Set ctl = pg1.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdViewDoing"
    ctl.Caption = "查看"
    ctl.Left = 12: ctl.Top = 280
    ctl.Width = 70: ctl.Height = 28

    ' Page 2 - 已完成列表
    Dim pg2 As Object
    Set pg2 = mp.Pages(2)

    Set ctl = pg2.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstDone"
    ctl.Left = 12: ctl.Top = 12
    ctl.Width = 498: ctl.Height = 260
    ctl.ColumnCount = 6
    ctl.ColumnWidths = "40;120;80;60;80;80"

    Set ctl = pg2.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdViewDone"
    ctl.Caption = "查看"
    ctl.Left = 12: ctl.Top = 280
    ctl.Width = 70: ctl.Height = 28

    ' 返回按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdBack"
    ctl.Caption = "返回"
    ctl.Left = 460: ctl.Top = 416
    ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "Private showAll As Boolean" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    showAll = False" & vbCrLf
    c = c & "    Dim isHigh As Boolean" & vbCrLf
    c = c & "    isHigh = IsHighPrivilege(gCurrentRole)" & vbCrLf
    c = c & "    cmdAllWork.Visible = isHigh" & vbCrLf
    c = c & "    Dim canEdit As Boolean" & vbCrLf
    c = c & "    canEdit = isHigh" & vbCrLf
    c = c & "    cmdNewTask.Enabled = isHigh" & vbCrLf
    c = c & "    cmdEditTask.Enabled = isHigh" & vbCrLf
    c = c & "    cmdDeleteTask.Enabled = isHigh" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdMyWork_Click()" & vbCrLf
    c = c & "    showAll = False" & vbCrLf
    c = c & "    cmdMyWork.BackColor = RGB(70, 130, 180)" & vbCrLf
    c = c & "    cmdMyWork.ForeColor = RGB(255, 255, 255)" & vbCrLf
    c = c & "    cmdAllWork.BackColor = &H8000000F" & vbCrLf
    c = c & "    cmdAllWork.ForeColor = RGB(0, 0, 0)" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdAllWork_Click()" & vbCrLf
    c = c & "    showAll = True" & vbCrLf
    c = c & "    cmdAllWork.BackColor = RGB(70, 130, 180)" & vbCrLf
    c = c & "    cmdAllWork.ForeColor = RGB(255, 255, 255)" & vbCrLf
    c = c & "    cmdMyWork.BackColor = &H8000000F" & vbCrLf
    c = c & "    cmdMyWork.ForeColor = RGB(0, 0, 0)" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub RefreshAllLists()" & vbCrLf
    c = c & "    LoadWorkList lstPending, " & q & "待办" & q & vbCrLf
    c = c & "    LoadWorkList lstDoing, " & q & "办理中" & q & vbCrLf
    c = c & "    LoadWorkList lstDone, " & q & "已完成" & q & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadWorkList(lst As MSForms.ListBox, filterStatus As String)" & vbCrLf
    c = c & "    lst.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "工作管理" & q & ")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long, i As Long, idx As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    idx = 0" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If ws.Cells(i, 8).Value = filterStatus Then" & vbCrLf
    c = c & "            Dim canShow As Boolean" & vbCrLf
    c = c & "            canShow = False" & vbCrLf
    c = c & "            If showAll Then" & vbCrLf
    c = c & "                canShow = True" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                If ws.Cells(i, 6).Value = gCurrentUser Then canShow = True" & vbCrLf
    c = c & "                If ws.Cells(i, 7).Value = gCurrentUser Then canShow = True" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            If canShow Then" & vbCrLf
    c = c & "                lst.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "                lst.List(idx, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "                lst.List(idx, 2) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "                lst.List(idx, 3) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "                lst.List(idx, 4) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "                lst.List(idx, 5) = CStr(ws.Cells(i, 10).Value)" & vbCrLf
    c = c & "                idx = idx + 1" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Function GetSelectedWorkID(lst As MSForms.ListBox) As String" & vbCrLf
    c = c & "    GetSelectedWorkID = " & q & q & vbCrLf
    c = c & "    If lst.ListIndex >= 0 Then" & vbCrLf
    c = c & "        GetSelectedWorkID = lst.List(lst.ListIndex, 0)" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Function" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNewTask_Click()" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "您没有新建工作的权限！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditWorkID = " & q & q & vbCrLf
    c = c & "    ShowWorkEditForm" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdViewTask_Click()" & vbCrLf
    c = c & "    Dim wid As String" & vbCrLf
    c = c & "    wid = GetSelectedWorkID(lstPending)" & vbCrLf
    c = c & "    If wid = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条工作记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewWorkID = wid" & vbCrLf
    c = c & "    ShowWorkViewForm" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEditTask_Click()" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "您没有编辑权限！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim wid As String" & vbCrLf
    c = c & "    wid = GetSelectedWorkID(lstPending)" & vbCrLf
    c = c & "    If wid = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条工作记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditWorkID = wid" & vbCrLf
    c = c & "    ShowWorkEditForm" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDeleteTask_Click()" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "您没有删除权限！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim wid As String" & vbCrLf
    c = c & "    wid = GetSelectedWorkID(lstPending)" & vbCrLf
    c = c & "    If wid = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条工作记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确认删除该工作？" & q & ", vbQuestion + vbYesNo) = vbYes Then" & vbCrLf
    c = c & "        DeleteWorkByID wid" & vbCrLf
    c = c & "        RefreshAllLists" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdViewDoing_Click()" & vbCrLf
    c = c & "    Dim wid As String" & vbCrLf
    c = c & "    wid = GetSelectedWorkID(lstDoing)" & vbCrLf
    c = c & "    If wid = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条工作记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewWorkID = wid" & vbCrLf
    c = c & "    ShowWorkViewForm" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdViewDone_Click()" & vbCrLf
    c = c & "    Dim wid As String" & vbCrLf
    c = c & "    wid = GetSelectedWorkID(lstDone)" & vbCrLf
    c = c & "    If wid = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条工作记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewWorkID = wid" & vbCrLf
    c = c & "    ShowWorkViewForm" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdBack_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建工作编辑窗体 ----------

Private Sub CreateWorkEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("WorkEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "WorkEditFormName", actualName

    vbc.Properties("Caption") = "任务/工单编辑"
    vbc.Properties("Width") = 460
    vbc.Properties("Height") = 380
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 类型框架
    Set ctl = dsg.Controls.Add("Forms.Frame.1")
    ctl.Name = "fraType"
    ctl.Caption = "选择类型"
    ctl.Left = 24: ctl.Top = 10
    ctl.Width = 400: ctl.Height = 48

    Dim typeFrame As Object
    Set typeFrame = ctl

    Set ctl = typeFrame.Controls.Add("Forms.OptionButton.1")
    ctl.Name = "optTask"
    ctl.Caption = "任务"
    ctl.Left = 20: ctl.Top = 18
    ctl.Width = 80: ctl.Height = 18
    ctl.Value = True

    Set ctl = typeFrame.Controls.Add("Forms.OptionButton.1")
    ctl.Name = "optWorkOrder"
    ctl.Caption = "工单"
    ctl.Left = 120: ctl.Top = 18
    ctl.Width = 80: ctl.Height = 18

    ' 标题
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label1"
    ctl.Caption = "标题："
    ctl.Left = 24: ctl.Top = 70
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtTitle"
    ctl.Left = 84: ctl.Top = 70
    ctl.Width = 340: ctl.Height = 24

    ' 描述
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label2"
    ctl.Caption = "描述："
    ctl.Left = 24: ctl.Top = 102
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtDesc"
    ctl.Left = 84: ctl.Top = 102
    ctl.Width = 340: ctl.Height = 60
    ctl.MultiLine = True
    ctl.ScrollBars = 2

    ' 工作类型
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblWorkType"
    ctl.Caption = "工作类型："
    ctl.Left = 24: ctl.Top = 174
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboWorkType"
    ctl.Left = 84: ctl.Top = 174
    ctl.Width = 130: ctl.Height = 22
    ctl.Style = 2

    ' 优先级
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPriority"
    ctl.Caption = "优先级："
    ctl.Left = 230: ctl.Top = 174
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboPriority"
    ctl.Left = 294: ctl.Top = 174
    ctl.Width = 130: ctl.Height = 22
    ctl.Style = 2

    ' 负责人
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label3"
    ctl.Caption = "负责人："
    ctl.Left = 24: ctl.Top = 208
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboAssignee"
    ctl.Left = 84: ctl.Top = 208
    ctl.Width = 150: ctl.Height = 24
    ctl.Style = 0

    ' 截止日期
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label4"
    ctl.Caption = "截止日期："
    ctl.Left = 24: ctl.Top = 244
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtDueDate"
    ctl.Left = 84: ctl.Top = 244
    ctl.Width = 150: ctl.Height = 24
    ctl.Text = ""

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDateHint"
    ctl.Caption = "(格式: yyyy-mm-dd)"
    ctl.Left = 240: ctl.Top = 248
    ctl.Width = 120: ctl.Height = 14
    ctl.ForeColor = RGB(150, 150, 150)
    ctl.Font.Size = 8

    ' 保存/取消
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave"
    ctl.Caption = "保存"
    ctl.Left = 120: ctl.Top = 290
    ctl.Width = 80: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel"
    ctl.Caption = "取消"
    ctl.Left = 240: ctl.Top = 290
    ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboWorkType.AddItem " & q & "维修" & q & vbCrLf
    c = c & "    cboWorkType.AddItem " & q & "保洁" & q & vbCrLf
    c = c & "    cboWorkType.AddItem " & q & "检查" & q & vbCrLf
    c = c & "    cboWorkType.AddItem " & q & "其他" & q & vbCrLf
    c = c & "    cboPriority.AddItem " & q & "高" & q & vbCrLf
    c = c & "    cboPriority.AddItem " & q & "中" & q & vbCrLf
    c = c & "    cboPriority.AddItem " & q & "低" & q & vbCrLf
    c = c & "    cboPriority.ListIndex = 1" & vbCrLf
    c = c & "    LoadUserList" & vbCrLf
    c = c & "    txtDueDate.Text = Format(Date + 7, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    If gEditWorkID <> " & q & q & " Then LoadWorkData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadUserList()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        cboAssignee.AddItem ws.Cells(i, 2).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadWorkData()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "工作管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gEditWorkID Then" & vbCrLf
    c = c & "            txtTitle.Text = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            txtDesc.Text = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            cboWorkType.Text = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            cboPriority.Text = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            cboAssignee.Text = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            txtDueDate.Text = CStr(ws.Cells(i, 10).Value)" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    Dim sTitle As String, sDesc As String" & vbCrLf
    c = c & "    Dim sType As String, sPri As String" & vbCrLf
    c = c & "    Dim sAssignee As String, sDue As String" & vbCrLf
    c = c & "    sTitle = Trim(txtTitle.Text)" & vbCrLf
    c = c & "    sDesc = Trim(txtDesc.Text)" & vbCrLf
    c = c & "    sType = cboWorkType.Text" & vbCrLf
    c = c & "    sPri = cboPriority.Text" & vbCrLf
    c = c & "    sAssignee = Trim(cboAssignee.Text)" & vbCrLf
    c = c & "    sDue = Trim(txtDueDate.Text)" & vbCrLf
    c = c & "    If sTitle = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请输入工作标题！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If sType = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择工作类型！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If sAssignee = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择负责人！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If gEditWorkID <> " & q & q & " Then" & vbCrLf
    c = c & "        UpdateWorkRecord gEditWorkID, sTitle, sDesc, sType, sPri, sAssignee, sDue" & vbCrLf
    c = c & "        MsgBox " & q & "工作信息已更新！" & q & ", vbInformation" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        AddWorkRecord sTitle, sDesc, sType, sPri, sAssignee, sDue" & vbCrLf
    c = c & "        MsgBox " & q & "工作已创建并派发！" & q & ", vbInformation" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建工作查看窗体 ----------

Private Sub CreateWorkViewForm()
    Dim oldForm As String
    oldForm = GetConfigProp("WorkViewFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "WorkViewFormName", actualName

    vbc.Properties("Caption") = "工作详情"
    vbc.Properties("Width") = 440
    vbc.Properties("Height") = 380
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yPos As Long
    yPos = 20

    ' 工作标题
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTitleCap"
    ctl.Caption = "工作标题："
    ctl.Left = 20: ctl.Top = yPos
    ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTitleVal"
    ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yPos
    ctl.Width = 310: ctl.Height = 18

    yPos = yPos + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTypeCap"
    ctl.Caption = "工作类型："
    ctl.Left = 20: ctl.Top = yPos
    ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTypeVal"
    ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yPos
    ctl.Width = 100: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPriCap"
    ctl.Caption = "优先级："
    ctl.Left = 220: ctl.Top = yPos
    ctl.Width = 60: ctl.Height = 18
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPriVal"
    ctl.Caption = ""
    ctl.Left = 290: ctl.Top = yPos
    ctl.Width = 100: ctl.Height = 18

    yPos = yPos + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDescCap"
    ctl.Caption = "工作描述："
    ctl.Left = 20: ctl.Top = yPos
    ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtDescView"
    ctl.Left = 100: ctl.Top = yPos
    ctl.Width = 310: ctl.Height = 60
    ctl.MultiLine = True
    ctl.Locked = True
    ctl.ScrollBars = 2

    yPos = yPos + 70
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAssigneeCap"
    ctl.Caption = "负责人："
    ctl.Left = 20: ctl.Top = yPos
    ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAssigneeVal"
    ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yPos
    ctl.Width = 100: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCreatorCap"
    ctl.Caption = "创建人："
    ctl.Left = 220: ctl.Top = yPos
    ctl.Width = 60: ctl.Height = 18
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCreatorVal"
    ctl.Caption = ""
    ctl.Left = 290: ctl.Top = yPos
    ctl.Width = 100: ctl.Height = 18

    yPos = yPos + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStatusCap"
    ctl.Caption = "当前状态："
    ctl.Left = 20: ctl.Top = yPos
    ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStatusVal"
    ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yPos
    ctl.Width = 100: ctl.Height = 20
    ctl.Font.Size = 11
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDueCap"
    ctl.Caption = "截止时间："
    ctl.Left = 220: ctl.Top = yPos
    ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDueVal"
    ctl.Caption = ""
    ctl.Left = 290: ctl.Top = yPos
    ctl.Width = 120: ctl.Height = 18

    yPos = yPos + 40
    ' 操作按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSetDoing"
    ctl.Caption = "办理中"
    ctl.Left = 40: ctl.Top = yPos
    ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(255, 165, 0)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSetDone"
    ctl.Caption = "已完成"
    ctl.Left = 170: ctl.Top = yPos
    ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(60, 179, 113)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose"
    ctl.Caption = "关闭"
    ctl.Left = 300: ctl.Top = yPos
    ctl.Width = 100: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    If gViewWorkID = " & q & q & " Then" & vbCrLf
    c = c & "        Unload Me" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "工作管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gViewWorkID Then" & vbCrLf
    c = c & "            lblTitleVal.Caption = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            txtDescView.Text = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            lblTypeVal.Caption = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lblPriVal.Caption = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lblAssigneeVal.Caption = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            lblCreatorVal.Caption = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            lblStatusVal.Caption = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            lblDueVal.Caption = CStr(ws.Cells(i, 10).Value)" & vbCrLf
    c = c & "            Dim st As String" & vbCrLf
    c = c & "            st = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            If st = " & q & "待办" & q & " Then" & vbCrLf
    c = c & "                lblStatusVal.ForeColor = RGB(255, 140, 0)" & vbCrLf
    c = c & "                cmdSetDoing.Enabled = True" & vbCrLf
    c = c & "                cmdSetDone.Enabled = True" & vbCrLf
    c = c & "            ElseIf st = " & q & "办理中" & q & " Then" & vbCrLf
    c = c & "                lblStatusVal.ForeColor = RGB(30, 144, 255)" & vbCrLf
    c = c & "                cmdSetDoing.Enabled = False" & vbCrLf
    c = c & "                cmdSetDone.Enabled = True" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lblStatusVal.ForeColor = RGB(60, 179, 113)" & vbCrLf
    c = c & "                cmdSetDoing.Enabled = False" & vbCrLf
    c = c & "                cmdSetDone.Enabled = False" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSetDoing_Click()" & vbCrLf
    c = c & "    UpdateWorkStatus gViewWorkID, " & q & "办理中" & q & vbCrLf
    c = c & "    MsgBox " & q & "已标记为办理中！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSetDone_Click()" & vbCrLf
    c = c & "    UpdateWorkStatus gViewWorkID, " & q & "已完成" & q & vbCrLf
    c = c & "    MsgBox " & q & "已标记为已完成！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建公告列表窗体 ----------

Private Sub CreateNoticesForm()
    Dim oldForm As String
    oldForm = GetConfigProp("NoticesFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "NoticesFormName", actualName

    vbc.Properties("Caption") = "信息管理 - 公告列表"
    vbc.Properties("Width") = 520
    vbc.Properties("Height") = 420
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 公告列表
    Set ctl = dsg.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstNotices"
    ctl.Left = 12: ctl.Top = 12
    ctl.Width = 490: ctl.Height = 320
    ctl.ColumnCount = 7
    ctl.ColumnWidths = "50;140;60;60;70;50;50"

    ' 列头标签
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblHdr"
    ctl.Caption = "ID | 标题 | 类型 | 部门 | 对象 | 置顶 | 紧急"
    ctl.Left = 12: ctl.Top = 334
    ctl.Width = 490: ctl.Height = 14
    ctl.Font.Size = 8
    ctl.ForeColor = RGB(120, 120, 120)

    ' 按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNewNotice"
    ctl.Caption = "新建公告"
    ctl.Left = 12: ctl.Top = 354
    ctl.Width = 100: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdViewNotice"
    ctl.Caption = "查看/编辑"
    ctl.Left = 120: ctl.Top = 354
    ctl.Width = 100: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDeleteNotice"
    ctl.Caption = "删除"
    ctl.Left = 228: ctl.Top = 354
    ctl.Width = 100: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdBack"
    ctl.Caption = "返回"
    ctl.Left = 420: ctl.Top = 354
    ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Dim isHigh As Boolean" & vbCrLf
    c = c & "    isHigh = IsHighPrivilege(gCurrentRole)" & vbCrLf
    c = c & "    cmdNewNotice.Enabled = isHigh" & vbCrLf
    c = c & "    cmdViewNotice.Enabled = True" & vbCrLf
    c = c & "    cmdDeleteNotice.Enabled = isHigh" & vbCrLf
    c = c & "    RefreshList" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub RefreshList()" & vbCrLf
    c = c & "    lstNotices.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "信息管理" & q & ")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long, i As Long, idx As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    idx = 0" & vbCrLf
    c = c & "    Dim isHigh As Boolean" & vbCrLf
    c = c & "    isHigh = IsHighPrivilege(gCurrentRole)" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim target As String" & vbCrLf
    c = c & "        target = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "        Dim canSee As Boolean" & vbCrLf
    c = c & "        canSee = False" & vbCrLf
    c = c & "        If isHigh Then" & vbCrLf
    c = c & "            canSee = True" & vbCrLf
    c = c & "        ElseIf target = " & q & "所有人" & q & " Then" & vbCrLf
    c = c & "            canSee = True" & vbCrLf
    c = c & "        ElseIf InStr(target, gCurrentUser) > 0 Then" & vbCrLf
    c = c & "            canSee = True" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        If canSee Then" & vbCrLf
    c = c & "            lstNotices.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lstNotices.List(idx, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lstNotices.List(idx, 2) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lstNotices.List(idx, 3) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lstNotices.List(idx, 4) = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            If ws.Cells(i, 8).Value = " & q & "是" & q & " Then" & vbCrLf
    c = c & "                lstNotices.List(idx, 5) = " & q & "[置顶]" & q & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lstNotices.List(idx, 5) = " & q & q & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            lstNotices.List(idx, 6) = ws.Cells(i, 9).Value" & vbCrLf
    c = c & "            idx = idx + 1" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNewNotice_Click()" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "您没有新建公告的权限！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditNoticeID = " & q & q & vbCrLf
    c = c & "    ShowNoticeEditForm" & vbCrLf
    c = c & "    RefreshList" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdViewNotice_Click()" & vbCrLf
    c = c & "    If lstNotices.ListIndex = -1 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条公告！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditNoticeID = lstNotices.List(lstNotices.ListIndex, 0)" & vbCrLf
    c = c & "    ShowNoticeEditForm" & vbCrLf
    c = c & "    RefreshList" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDeleteNotice_Click()" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "您没有删除权限！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If lstNotices.ListIndex = -1 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条公告！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim nid As String" & vbCrLf
    c = c & "    nid = lstNotices.List(lstNotices.ListIndex, 0)" & vbCrLf
    c = c & "    If MsgBox(" & q & "确认删除该公告？" & q & ", vbQuestion + vbYesNo) = vbYes Then" & vbCrLf
    c = c & "        DeleteNoticeByID nid" & vbCrLf
    c = c & "        RefreshList" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdBack_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建公告编辑窗体 ----------

Private Sub CreateNoticeEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("NoticeEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "NoticeEditFormName", actualName

    vbc.Properties("Caption") = "公告编辑"
    vbc.Properties("Width") = 440
    vbc.Properties("Height") = 420
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 标题
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label1"
    ctl.Caption = "标题："
    ctl.Left = 24: ctl.Top = 24
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtTitle"
    ctl.Left = 84: ctl.Top = 24
    ctl.Width = 320: ctl.Height = 24

    ' 内容
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label2"
    ctl.Caption = "内容："
    ctl.Left = 24: ctl.Top = 54
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtContent"
    ctl.Left = 84: ctl.Top = 54
    ctl.Width = 320: ctl.Height = 80
    ctl.MultiLine = True
    ctl.ScrollBars = 2

    ' 类型
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblType"
    ctl.Caption = "类型："
    ctl.Left = 24: ctl.Top = 144
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboType"
    ctl.Left = 84: ctl.Top = 144
    ctl.Width = 130: ctl.Height = 22
    ctl.Style = 2

    ' 发布部门
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label3"
    ctl.Caption = "发布部门："
    ctl.Left = 230: ctl.Top = 144
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboDept"
    ctl.Left = 296: ctl.Top = 144
    ctl.Width = 108: ctl.Height = 22
    ctl.Style = 2

    ' 发布人
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label4"
    ctl.Caption = "发布人："
    ctl.Left = 24: ctl.Top = 176
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtAuthor"
    ctl.Left = 84: ctl.Top = 176
    ctl.Width = 130: ctl.Height = 22
    ctl.Enabled = False

    ' 对象
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTarget"
    ctl.Caption = "对象："
    ctl.Left = 230: ctl.Top = 176
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboTarget"
    ctl.Left = 296: ctl.Top = 176
    ctl.Width = 108: ctl.Height = 22
    ctl.Style = 0

    ' 紧急公告
    Set ctl = dsg.Controls.Add("Forms.CheckBox.1")
    ctl.Name = "chkUrgent"
    ctl.Caption = "紧急公告"
    ctl.Left = 24: ctl.Top = 210
    ctl.Width = 100: ctl.Height = 18

    ' 置顶
    Set ctl = dsg.Controls.Add("Forms.CheckBox.1")
    ctl.Name = "chkTop"
    ctl.Caption = "置顶显示"
    ctl.Left = 140: ctl.Top = 210
    ctl.Width = 100: ctl.Height = 18

    ' 紧急程度
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblUrgLevel"
    ctl.Caption = "紧急程度："
    ctl.Left = 260: ctl.Top = 210
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboUrgLevel"
    ctl.Left = 326: ctl.Top = 208
    ctl.Width = 78: ctl.Height = 22
    ctl.Style = 2

    ' 生效时间
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label5"
    ctl.Caption = "生效时间："
    ctl.Left = 24: ctl.Top = 244
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtStart"
    ctl.Left = 84: ctl.Top = 244
    ctl.Width = 120: ctl.Height = 24

    ' 截止时间
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label6"
    ctl.Caption = "截止时间："
    ctl.Left = 230: ctl.Top = 244
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtEnd"
    ctl.Left = 296: ctl.Top = 244
    ctl.Width = 120: ctl.Height = 24

    ' 日期格式提示
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDateFmt"
    ctl.Caption = "(日期格式: yyyy-mm-dd)"
    ctl.Left = 84: ctl.Top = 272
    ctl.Width = 200: ctl.Height = 14
    ctl.ForeColor = RGB(150, 150, 150)
    ctl.Font.Size = 8

    ' 保存/取消
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave"
    ctl.Caption = "保存"
    ctl.Left = 100: ctl.Top = 300
    ctl.Width = 80: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel"
    ctl.Caption = "取消"
    ctl.Left = 240: ctl.Top = 300
    ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboType.AddItem " & q & "通知" & q & vbCrLf
    c = c & "    cboType.AddItem " & q & "公告" & q & vbCrLf
    c = c & "    cboType.AddItem " & q & "紧急通知" & q & vbCrLf
    c = c & "    cboType.ListIndex = 0" & vbCrLf
    c = c & "    cboDept.AddItem " & q & "客服部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "工程部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "安保部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "保洁部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "综合部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "行政部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "财务部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "销售部" & q & vbCrLf
    c = c & "    cboUrgLevel.AddItem " & q & "普通" & q & vbCrLf
    c = c & "    cboUrgLevel.AddItem " & q & "重要" & q & vbCrLf
    c = c & "    cboUrgLevel.AddItem " & q & "紧急" & q & vbCrLf
    c = c & "    cboUrgLevel.ListIndex = 0" & vbCrLf
    c = c & "    LoadTargetList" & vbCrLf
    c = c & "    txtAuthor.Text = gCurrentUser" & vbCrLf
    c = c & "    txtStart.Text = Format(Date, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    txtEnd.Text = Format(Date + 30, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    If gEditNoticeID <> " & q & q & " Then" & vbCrLf
    c = c & "        LoadNoticeData" & vbCrLf
    c = c & "        If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "            txtTitle.Locked = True" & vbCrLf
    c = c & "            txtContent.Locked = True" & vbCrLf
    c = c & "            cboType.Enabled = False" & vbCrLf
    c = c & "            cboDept.Enabled = False" & vbCrLf
    c = c & "            cboTarget.Enabled = False" & vbCrLf
    c = c & "            chkUrgent.Enabled = False" & vbCrLf
    c = c & "            chkTop.Enabled = False" & vbCrLf
    c = c & "            cboUrgLevel.Enabled = False" & vbCrLf
    c = c & "            txtStart.Locked = True" & vbCrLf
    c = c & "            txtEnd.Locked = True" & vbCrLf
    c = c & "            cmdSave.Enabled = False" & vbCrLf
    c = c & "            Me.Caption = " & q & "公告查看（只读）" & q & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadTargetList()" & vbCrLf
    c = c & "    cboTarget.AddItem " & q & "所有人" & q & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim uName As String, uAcct As String" & vbCrLf
    c = c & "        uName = ws.Cells(i, 1).Value" & vbCrLf
    c = c & "        uAcct = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "        cboTarget.AddItem uName & " & q & " (" & q & " & uAcct & " & q & ")" & q & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    cboTarget.ListIndex = 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadNoticeData()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "信息管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gEditNoticeID Then" & vbCrLf
    c = c & "            txtTitle.Text = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            txtContent.Text = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            cboType.Text = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            cboDept.Text = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            txtAuthor.Text = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            cboTarget.Text = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            chkTop.Value = (ws.Cells(i, 8).Value = " & q & "是" & q & ")" & vbCrLf
    c = c & "            cboUrgLevel.Text = ws.Cells(i, 9).Value" & vbCrLf
    c = c & "            chkUrgent.Value = (ws.Cells(i, 9).Value = " & q & "紧急" & q & ")" & vbCrLf
    c = c & "            txtStart.Text = CStr(ws.Cells(i, 10).Value)" & vbCrLf
    c = c & "            txtEnd.Text = CStr(ws.Cells(i, 11).Value)" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub chkUrgent_Click()" & vbCrLf
    c = c & "    If chkUrgent.Value Then" & vbCrLf
    c = c & "        cboUrgLevel.Text = " & q & "紧急" & q & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        cboUrgLevel.Text = " & q & "普通" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    Dim sTitle As String, sContent As String" & vbCrLf
    c = c & "    Dim sType As String, sDept As String" & vbCrLf
    c = c & "    Dim sTarget As String, sUrg As String" & vbCrLf
    c = c & "    Dim sStart As String, sEnd2 As String" & vbCrLf
    c = c & "    Dim stopp As String" & vbCrLf
    c = c & "    sTitle = Trim(txtTitle.Text)" & vbCrLf
    c = c & "    sContent = Trim(txtContent.Text)" & vbCrLf
    c = c & "    sType = cboType.Text" & vbCrLf
    c = c & "    sDept = cboDept.Text" & vbCrLf
    c = c & "    sTarget = cboTarget.Text" & vbCrLf
    c = c & "    sUrg = cboUrgLevel.Text" & vbCrLf
    c = c & "    sStart = Trim(txtStart.Text)" & vbCrLf
    c = c & "    sEnd2 = Trim(txtEnd.Text)" & vbCrLf
    c = c & "    If chkTop.Value Then stopp = " & q & "是" & q
    c = c & " Else stopp = " & q & "否" & q & vbCrLf
    c = c & "    If sTitle = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请输入公告标题！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If sContent = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请输入公告内容！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If sDept = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择发布部门！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If gEditNoticeID <> " & q & q & " Then" & vbCrLf
    c = c & "        UpdateNoticeRecord gEditNoticeID, sTitle, sContent, sType, sDept, sTarget, stopp, sUrg, sStart, sEnd2" & vbCrLf
    c = c & "        MsgBox " & q & "公告已更新！" & q & ", vbInformation" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        AddNoticeRecord sTitle, sContent, sType, sDept, sTarget, stopp, sUrg, sStart, sEnd2" & vbCrLf
    c = c & "        MsgBox " & q & "公告已发布！" & q & ", vbInformation" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建审批管理表 ----------

Private Sub CreateApprovalSheet()
    Dim ws As Worksheet
    If SheetExists("审批管理") Then
        Set ws = ThisWorkbook.Sheets("审批管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "审批管理"
    End If
    With ws
        .Range("A1").Value = "申请ID"
        .Range("B1").Value = "申请类型"
        .Range("C1").Value = "子类型"
        .Range("D1").Value = "申请人"
        .Range("E1").Value = "审批人"
        .Range("F1").Value = "事由"
        .Range("G1").Value = "开始日期"
        .Range("H1").Value = "结束日期"
        .Range("I1").Value = "天数"
        .Range("J1").Value = "金额"
        .Range("K1").Value = "物品/项目"
        .Range("L1").Value = "状态"
        .Range("M1").Value = "审批意见"
        .Range("N1").Value = "申请时间"
        .Range("O1").Value = "审批时间"
        .Range("A1:O1").Font.Bold = True
        .Range("A1:O1").Interior.Color = RGB(70, 130, 180)
        .Range("A1:O1").Font.Color = RGB(255, 255, 255)
        .Columns("A:O").AutoFit
    End With
End Sub

' ---------- 创建审批中心窗体 ----------

Private Sub CreateApprovalsForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ApprovalsFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ApprovalsFormName", actualName

    vbc.Properties("Caption") = "审批管理"
    vbc.Properties("Width") = 560
    vbc.Properties("Height") = 500
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 切换卡片
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdMyApply"
    ctl.Caption = "我的申请"
    ctl.Left = 12: ctl.Top = 6
    ctl.Width = 90: ctl.Height = 26
    ctl.BackColor = RGB(70, 130, 180)
    ctl.ForeColor = RGB(255, 255, 255)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdPendApproval"
    ctl.Caption = "待审批"
    ctl.Left = 110: ctl.Top = 6
    ctl.Width = 90: ctl.Height = 26

    ' MultiPage
    Set ctl = dsg.Controls.Add("Forms.MultiPage.1")
    ctl.Name = "MultiPage1"
    ctl.Left = 12: ctl.Top = 38
    ctl.Width = 530: ctl.Height = 370

    Dim mp As Object
    Set mp = ctl
    mp.Pages(0).Caption = "待审批"
    mp.Pages.Add
    mp.Pages(1).Caption = "已同意"
    mp.Pages.Add
    mp.Pages(2).Caption = "已拒绝"

    ' Page 0 - 待审批列表
    Dim pg0 As Object
    Set pg0 = mp.Pages(0)

    Set ctl = pg0.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstPending"
    ctl.Left = 12: ctl.Top = 12
    ctl.Width = 498: ctl.Height = 230
    ctl.ColumnCount = 6
    ctl.ColumnWidths = "50;80;60;70;70;80"

    Set ctl = pg0.Controls.Add("Forms.Label.1")
    ctl.Name = "lblHdr0"
    ctl.Caption = "ID | 类型 | 子类型 | 申请人 | 审批人 | 状态"
    ctl.Left = 12: ctl.Top = 244
    ctl.Width = 498: ctl.Height = 14
    ctl.Font.Size = 8
    ctl.ForeColor = RGB(120, 120, 120)

    Set ctl = pg0.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNewApply"
    ctl.Caption = "新建申请"
    ctl.Left = 12: ctl.Top = 264
    ctl.Width = 80: ctl.Height = 28

    Set ctl = pg0.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdViewApply"
    ctl.Caption = "查看"
    ctl.Left = 102: ctl.Top = 264
    ctl.Width = 70: ctl.Height = 28

    Set ctl = pg0.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdEditApply"
    ctl.Caption = "编辑"
    ctl.Left = 182: ctl.Top = 264
    ctl.Width = 70: ctl.Height = 28

    Set ctl = pg0.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDeleteApply"
    ctl.Caption = "删除"
    ctl.Left = 262: ctl.Top = 264
    ctl.Width = 70: ctl.Height = 28

    ' Page 1 - 已同意
    Dim pg1 As Object
    Set pg1 = mp.Pages(1)

    Set ctl = pg1.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstApproved"
    ctl.Left = 12: ctl.Top = 12
    ctl.Width = 498: ctl.Height = 270
    ctl.ColumnCount = 6
    ctl.ColumnWidths = "50;80;60;70;70;80"

    Set ctl = pg1.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdViewApproved"
    ctl.Caption = "查看"
    ctl.Left = 12: ctl.Top = 286
    ctl.Width = 70: ctl.Height = 28

    ' Page 2 - 已拒绝
    Dim pg2 As Object
    Set pg2 = mp.Pages(2)

    Set ctl = pg2.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstRejected"
    ctl.Left = 12: ctl.Top = 12
    ctl.Width = 498: ctl.Height = 270
    ctl.ColumnCount = 6
    ctl.ColumnWidths = "50;80;60;70;70;80"

    Set ctl = pg2.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdViewRejected"
    ctl.Caption = "查看"
    ctl.Left = 12: ctl.Top = 286
    ctl.Width = 70: ctl.Height = 28

    ' 返回按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdBack"
    ctl.Caption = "返回"
    ctl.Left = 460: ctl.Top = 420
    ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "Private showPend As Boolean" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    showPend = False" & vbCrLf
    c = c & "    Dim isHigh As Boolean" & vbCrLf
    c = c & "    isHigh = IsHighPrivilege(gCurrentRole)" & vbCrLf
    c = c & "    cmdPendApproval.Visible = isHigh" & vbCrLf
    c = c & "    cmdEditApply.Enabled = isHigh" & vbCrLf
    c = c & "    cmdDeleteApply.Enabled = isHigh" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdMyApply_Click()" & vbCrLf
    c = c & "    showPend = False" & vbCrLf
    c = c & "    cmdMyApply.BackColor = RGB(70, 130, 180)" & vbCrLf
    c = c & "    cmdMyApply.ForeColor = RGB(255, 255, 255)" & vbCrLf
    c = c & "    cmdPendApproval.BackColor = &H8000000F" & vbCrLf
    c = c & "    cmdPendApproval.ForeColor = RGB(0, 0, 0)" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdPendApproval_Click()" & vbCrLf
    c = c & "    showPend = True" & vbCrLf
    c = c & "    cmdPendApproval.BackColor = RGB(70, 130, 180)" & vbCrLf
    c = c & "    cmdPendApproval.ForeColor = RGB(255, 255, 255)" & vbCrLf
    c = c & "    cmdMyApply.BackColor = &H8000000F" & vbCrLf
    c = c & "    cmdMyApply.ForeColor = RGB(0, 0, 0)" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub RefreshAllLists()" & vbCrLf
    c = c & "    LoadApprovalList lstPending, " & q & "待审批" & q & vbCrLf
    c = c & "    LoadApprovalList lstApproved, " & q & "已同意" & q & vbCrLf
    c = c & "    LoadApprovalList lstRejected, " & q & "已拒绝" & q & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadApprovalList(lst As MSForms.ListBox, filterSt As String)" & vbCrLf
    c = c & "    lst.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "审批管理" & q & ")" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If ws Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim lr As Long, i As Long, idx As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    idx = 0" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If ws.Cells(i, 12).Value = filterSt Then" & vbCrLf
    c = c & "            Dim canShow As Boolean" & vbCrLf
    c = c & "            canShow = False" & vbCrLf
    c = c & "            If showPend Then" & vbCrLf
    c = c & "                canShow = True" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                If ws.Cells(i, 4).Value = gCurrentUser Then canShow = True" & vbCrLf
    c = c & "                If IsHighPrivilege(gCurrentRole) Then canShow = True" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            If canShow Then" & vbCrLf
    c = c & "                lst.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "                lst.List(idx, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "                lst.List(idx, 2) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "                lst.List(idx, 3) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "                lst.List(idx, 4) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "                lst.List(idx, 5) = ws.Cells(i, 12).Value" & vbCrLf
    c = c & "                idx = idx + 1" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Function GetSelID(lst As MSForms.ListBox) As String" & vbCrLf
    c = c & "    GetSelID = " & q & q & vbCrLf
    c = c & "    If lst.ListIndex >= 0 Then GetSelID = lst.List(lst.ListIndex, 0)" & vbCrLf
    c = c & "End Function" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNewApply_Click()" & vbCrLf
    c = c & "    gEditApprovalID = " & q & q & vbCrLf
    c = c & "    ShowApprovalEditForm" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdViewApply_Click()" & vbCrLf
    c = c & "    Dim aid As String" & vbCrLf
    c = c & "    aid = GetSelID(lstPending)" & vbCrLf
    c = c & "    If aid = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewApprovalID = aid" & vbCrLf
    c = c & "    ShowApprovalViewForm" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEditApply_Click()" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "您没有编辑权限！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim aid As String" & vbCrLf
    c = c & "    aid = GetSelID(lstPending)" & vbCrLf
    c = c & "    If aid = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditApprovalID = aid" & vbCrLf
    c = c & "    ShowApprovalEditForm" & vbCrLf
    c = c & "    RefreshAllLists" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDeleteApply_Click()" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "您没有删除权限！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim aid As String" & vbCrLf
    c = c & "    aid = GetSelID(lstPending)" & vbCrLf
    c = c & "    If aid = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择一条记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确认删除？" & q & ", vbQuestion + vbYesNo) = vbYes Then" & vbCrLf
    c = c & "        DeleteApprovalByID aid" & vbCrLf
    c = c & "        RefreshAllLists" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdViewApproved_Click()" & vbCrLf
    c = c & "    Dim aid As String" & vbCrLf
    c = c & "    aid = GetSelID(lstApproved)" & vbCrLf
    c = c & "    If aid = " & q & q & " Then Exit Sub" & vbCrLf
    c = c & "    gViewApprovalID = aid" & vbCrLf
    c = c & "    ShowApprovalViewForm" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdViewRejected_Click()" & vbCrLf
    c = c & "    Dim aid As String" & vbCrLf
    c = c & "    aid = GetSelID(lstRejected)" & vbCrLf
    c = c & "    If aid = " & q & q & " Then Exit Sub" & vbCrLf
    c = c & "    gViewApprovalID = aid" & vbCrLf
    c = c & "    ShowApprovalViewForm" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdBack_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建审批申请/编辑窗体 ----------

Private Sub CreateApprovalEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ApprovalEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ApprovalEditFormName", actualName

    vbc.Properties("Caption") = "审批申请/编辑"
    vbc.Properties("Width") = 500
    vbc.Properties("Height") = 400
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 申请类型
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label1"
    ctl.Caption = "申请类型："
    ctl.Left = 24: ctl.Top = 24
    ctl.Width = 70: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboType"
    ctl.Left = 96: ctl.Top = 24
    ctl.Width = 150: ctl.Height = 24
    ctl.Style = 2

    ' 子类型
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblSubType"
    ctl.Caption = "子类型："
    ctl.Left = 270: ctl.Top = 24
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboSubType"
    ctl.Left = 336: ctl.Top = 24
    ctl.Width = 130: ctl.Height = 24
    ctl.Style = 2

    ' 申请详情框架
    Set ctl = dsg.Controls.Add("Forms.Frame.1")
    ctl.Name = "fraDynamic"
    ctl.Caption = "申请详情"
    ctl.Left = 24: ctl.Top = 58
    ctl.Width = 452: ctl.Height = 220

    Dim fra As Object
    Set fra = ctl

    ' 开始时间
    Set ctl = fra.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStart"
    ctl.Caption = "开始时间："
    ctl.Left = 16: ctl.Top = 24
    ctl.Width = 70: ctl.Height = 18

    Set ctl = fra.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtStart"
    ctl.Left = 90: ctl.Top = 24
    ctl.Width = 110: ctl.Height = 22

    ' 结束时间
    Set ctl = fra.Controls.Add("Forms.Label.1")
    ctl.Name = "lblEnd"
    ctl.Caption = "结束时间："
    ctl.Left = 220: ctl.Top = 24
    ctl.Width = 70: ctl.Height = 18

    Set ctl = fra.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtEnd2"
    ctl.Left = 294: ctl.Top = 24
    ctl.Width = 110: ctl.Height = 22

    ' 天数
    Set ctl = fra.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDays"
    ctl.Caption = "天数："
    ctl.Left = 16: ctl.Top = 54
    ctl.Width = 50: ctl.Height = 18

    Set ctl = fra.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtDays"
    ctl.Left = 90: ctl.Top = 54
    ctl.Width = 60: ctl.Height = 22
    ctl.Enabled = False

    ' 事由
    Set ctl = fra.Controls.Add("Forms.Label.1")
    ctl.Name = "lblReason"
    ctl.Caption = "事由："
    ctl.Left = 16: ctl.Top = 84
    ctl.Width = 50: ctl.Height = 18

    Set ctl = fra.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtReason"
    ctl.Left = 90: ctl.Top = 84
    ctl.Width = 340: ctl.Height = 55
    ctl.MultiLine = True
    ctl.ScrollBars = 2

    ' 金额
    Set ctl = fra.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAmount"
    ctl.Caption = "金额："
    ctl.Left = 16: ctl.Top = 150
    ctl.Width = 50: ctl.Height = 18

    Set ctl = fra.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtAmount"
    ctl.Left = 90: ctl.Top = 150
    ctl.Width = 110: ctl.Height = 22

    ' 物品/项目
    Set ctl = fra.Controls.Add("Forms.Label.1")
    ctl.Name = "lblItem"
    ctl.Caption = "物品/项目："
    ctl.Left = 220: ctl.Top = 150
    ctl.Width = 70: ctl.Height = 18

    Set ctl = fra.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtItem"
    ctl.Left = 294: ctl.Top = 150
    ctl.Width = 136: ctl.Height = 22

    ' 申请人
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblApplicant"
    ctl.Caption = "申请人："
    ctl.Left = 24: ctl.Top = 288
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtApplicant"
    ctl.Left = 90: ctl.Top = 288
    ctl.Width = 110: ctl.Height = 22
    ctl.Enabled = False

    ' 审批人
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblApprover"
    ctl.Caption = "审批人："
    ctl.Left = 220: ctl.Top = 288
    ctl.Width = 60: ctl.Height = 18

    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtApprover"
    ctl.Left = 286: ctl.Top = 288
    ctl.Width = 110: ctl.Height = 22
    ctl.Enabled = False

    ' 保存/取消
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave"
    ctl.Caption = "保存"
    ctl.Left = 140: ctl.Top = 330
    ctl.Width = 80: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel"
    ctl.Caption = "取消"
    ctl.Left = 260: ctl.Top = 330
    ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboType.AddItem " & q & "请假申请" & q & vbCrLf
    c = c & "    cboType.AddItem " & q & "报销申请" & q & vbCrLf
    c = c & "    cboType.AddItem " & q & "采购申请" & q & vbCrLf
    c = c & "    cboType.AddItem " & q & "用章申请" & q & vbCrLf
    c = c & "    cboType.AddItem " & q & "其他" & q & vbCrLf
    c = c & "    cboType.ListIndex = 0" & vbCrLf
    c = c & "    txtApplicant.Text = gCurrentUser" & vbCrLf
    c = c & "    txtApprover.Text = GetDefaultApprover()" & vbCrLf
    c = c & "    txtStart.Text = Format(Date, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    txtEnd2.Text = Format(Date + 1, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    CalcDays" & vbCrLf
    c = c & "    UpdateSubType" & vbCrLf
    c = c & "    ToggleFields" & vbCrLf
    c = c & "    If gEditApprovalID <> " & q & q & " Then LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cboType_Change()" & vbCrLf
    c = c & "    UpdateSubType" & vbCrLf
    c = c & "    ToggleFields" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UpdateSubType()" & vbCrLf
    c = c & "    cboSubType.Clear" & vbCrLf
    c = c & "    Select Case cboType.Text" & vbCrLf
    c = c & "        Case " & q & "请假申请" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "外出" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "事假" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "年假" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "病假" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "婚假" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "产假" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "丧假" & q & vbCrLf
    c = c & "        Case " & q & "报销申请" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "差旅" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "办公" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "其他" & q & vbCrLf
    c = c & "        Case " & q & "用章申请" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "公章" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "合同章" & q & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "财务章" & q & vbCrLf
    c = c & "        Case Else" & vbCrLf
    c = c & "            cboSubType.AddItem " & q & "其他" & q & vbCrLf
    c = c & "    End Select" & vbCrLf
    c = c & "    If cboSubType.ListCount > 0 Then cboSubType.ListIndex = 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub ToggleFields()" & vbCrLf
    c = c & "    Dim t As String" & vbCrLf
    c = c & "    t = cboType.Text" & vbCrLf
    c = c & "    Dim showDate As Boolean, showMoney As Boolean" & vbCrLf
    c = c & "    showDate = (t = " & q & "请假申请" & q & " Or t = " & q & "用章申请" & q & ")" & vbCrLf
    c = c & "    showMoney = (t = " & q & "报销申请" & q & " Or t = " & q & "采购申请" & q & ")" & vbCrLf
    c = c & "    lblStart.Visible = showDate" & vbCrLf
    c = c & "    txtStart.Visible = showDate" & vbCrLf
    c = c & "    lblEnd.Visible = (t = " & q & "请假申请" & q & ")" & vbCrLf
    c = c & "    txtEnd2.Visible = (t = " & q & "请假申请" & q & ")" & vbCrLf
    c = c & "    lblDays.Visible = (t = " & q & "请假申请" & q & ")" & vbCrLf
    c = c & "    txtDays.Visible = (t = " & q & "请假申请" & q & ")" & vbCrLf
    c = c & "    lblAmount.Visible = showMoney" & vbCrLf
    c = c & "    txtAmount.Visible = showMoney" & vbCrLf
    c = c & "    lblItem.Visible = showMoney" & vbCrLf
    c = c & "    txtItem.Visible = showMoney" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub txtStart_Change()" & vbCrLf
    c = c & "    CalcDays" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub txtEnd2_Change()" & vbCrLf
    c = c & "    CalcDays" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub CalcDays()" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Dim d1 As Date, d2 As Date" & vbCrLf
    c = c & "    d1 = CDate(txtStart.Text)" & vbCrLf
    c = c & "    d2 = CDate(txtEnd2.Text)" & vbCrLf
    c = c & "    If Err.Number = 0 And d2 >= d1 Then" & vbCrLf
    c = c & "        txtDays.Text = CStr(DateDiff(" & q & "d" & q & ", d1, d2) + 1)" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        txtDays.Text = " & q & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "审批管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gEditApprovalID Then" & vbCrLf
    c = c & "            cboType.Text = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            UpdateSubType" & vbCrLf
    c = c & "            cboSubType.Text = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            txtApplicant.Text = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            txtApprover.Text = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            txtReason.Text = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            txtStart.Text = CStr(ws.Cells(i, 7).Value)" & vbCrLf
    c = c & "            txtEnd2.Text = CStr(ws.Cells(i, 8).Value)" & vbCrLf
    c = c & "            txtDays.Text = CStr(ws.Cells(i, 9).Value)" & vbCrLf
    c = c & "            txtAmount.Text = CStr(ws.Cells(i, 10).Value)" & vbCrLf
    c = c & "            txtItem.Text = ws.Cells(i, 11).Value" & vbCrLf
    c = c & "            ToggleFields" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If cboType.Text = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择申请类型！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(txtReason.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写事由！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If gEditApprovalID <> " & q & q & " Then" & vbCrLf
    c = c & "        UpdateApprovalRecord gEditApprovalID, cboType.Text, cboSubType.Text, txtReason.Text, txtStart.Text, txtEnd2.Text, txtDays.Text, txtAmount.Text, txtItem.Text" & vbCrLf
    c = c & "        MsgBox " & q & "申请已更新！" & q & ", vbInformation" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        AddApprovalRecord cboType.Text, cboSubType.Text, txtApplicant.Text, txtApprover.Text, txtReason.Text, txtStart.Text, txtEnd2.Text, txtDays.Text, txtAmount.Text, txtItem.Text" & vbCrLf
    c = c & "        MsgBox " & q & "申请已提交！" & q & ", vbInformation" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建审批查看窗体 ----------

Private Sub CreateApprovalViewForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ApprovalViewFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ApprovalViewFormName", actualName

    vbc.Properties("Caption") = "审批详情"
    vbc.Properties("Width") = 440
    vbc.Properties("Height") = 380
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 20

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTypeCap": ctl.Caption = "申请类型："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTypeVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblSubCap": ctl.Caption = "子类型："
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblSubVal": ctl.Caption = ""
    ctl.Left = 290: ctl.Top = yy: ctl.Width = 120: ctl.Height = 18

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAppCap": ctl.Caption = "申请人："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAppVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblApprvCap": ctl.Caption = "审批人："
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblApprvVal": ctl.Caption = ""
    ctl.Left = 290: ctl.Top = yy: ctl.Width = 120: ctl.Height = 18

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblReasonCap": ctl.Caption = "事由："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 50: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtReasonV": ctl.Left = 100: ctl.Top = yy
    ctl.Width = 310: ctl.Height = 55: ctl.MultiLine = True
    ctl.Locked = True: ctl.ScrollBars = 2

    yy = yy + 62
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDateCap": ctl.Caption = "日期："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 50: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDateVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 200: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAmtCap": ctl.Caption = "金额："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 50: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAmtVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblItemCap": ctl.Caption = "物品/项目："
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblItemVal": ctl.Caption = ""
    ctl.Left = 290: ctl.Top = yy: ctl.Width = 120: ctl.Height = 18

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStCap": ctl.Caption = "状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 50: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20
    ctl.Font.Size = 11: ctl.Font.Bold = True

    yy = yy + 36
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdApprove"
    ctl.Caption = "同意"
    ctl.Left = 40: ctl.Top = yy
    ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(60, 179, 113)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdReject"
    ctl.Caption = "拒绝"
    ctl.Left = 170: ctl.Top = yy
    ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(220, 80, 80)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose"
    ctl.Caption = "关闭"
    ctl.Left = 300: ctl.Top = yy
    ctl.Width = 100: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    If gViewApprovalID = " & q & q & " Then Unload Me: Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "审批管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gViewApprovalID Then" & vbCrLf
    c = c & "            lblTypeVal.Caption = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lblSubVal.Caption = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            lblAppVal.Caption = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lblApprvVal.Caption = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            txtReasonV.Text = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            lblDateVal.Caption = ws.Cells(i, 7).Value & " & q & " ~ " & q & " & ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            lblAmtVal.Caption = ws.Cells(i, 10).Value" & vbCrLf
    c = c & "            lblItemVal.Caption = ws.Cells(i, 11).Value" & vbCrLf
    c = c & "            lblStVal.Caption = ws.Cells(i, 12).Value" & vbCrLf
    c = c & "            Dim st As String" & vbCrLf
    c = c & "            st = ws.Cells(i, 12).Value" & vbCrLf
    c = c & "            If st = " & q & "待审批" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(255, 140, 0)" & vbCrLf
    c = c & "                Dim canApprove As Boolean" & vbCrLf
    c = c & "                canApprove = IsHighPrivilege(gCurrentRole)" & vbCrLf
    c = c & "                cmdApprove.Enabled = canApprove" & vbCrLf
    c = c & "                cmdReject.Enabled = canApprove" & vbCrLf
    c = c & "            ElseIf st = " & q & "已同意" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(60, 179, 113)" & vbCrLf
    c = c & "                cmdApprove.Enabled = False" & vbCrLf
    c = c & "                cmdReject.Enabled = False" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(220, 80, 80)" & vbCrLf
    c = c & "                cmdApprove.Enabled = False" & vbCrLf
    c = c & "                cmdReject.Enabled = False" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdApprove_Click()" & vbCrLf
    c = c & "    UpdateApprovalStatus gViewApprovalID, " & q & "已同意" & q & vbCrLf
    c = c & "    MsgBox " & q & "已同意该申请！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdReject_Click()" & vbCrLf
    c = c & "    UpdateApprovalStatus gViewApprovalID, " & q & "已拒绝" & q & vbCrLf
    c = c & "    MsgBox " & q & "已拒绝该申请！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建巡检管理表 ----------

Private Sub CreateInspectSheet()
    Dim ws As Worksheet
    If SheetExists("巡检管理") Then
        Set ws = ThisWorkbook.Sheets("巡检管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "巡检管理"
    End If
    With ws
        .Range("A1").Value = "巡检编号"
        .Range("B1").Value = "巡检员"
        .Range("C1").Value = "巡检区域"
        .Range("D1").Value = "巡检时间"
        .Range("E1").Value = "巡检状态"
        .Range("F1").Value = "巡检备注"
        .Range("G1").Value = "图片路径"
        .Range("H1").Value = "创建时间"
        .Range("A1:H1").Font.Bold = True
        .Columns("A:H").AutoFit
    End With
End Sub

' ---------- 创建巡检管理窗体 ----------

Private Sub CreateInspectManagerForm()
    Dim oldForm As String
    oldForm = GetConfigProp("InspectManagerFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "InspectManagerFormName", actualName

    vbc.Properties("Caption") = "巡检管理"
    vbc.Properties("Width") = 680
    vbc.Properties("Height") = 500
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 状态统计标签
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStats": ctl.Caption = "状态统计: 加载中..."
    ctl.Left = 20: ctl.Top = 10: ctl.Width = 500: ctl.Height = 18
    ctl.Font.Size = 10: ctl.Font.Bold = True

    ' 巡检任务列表
    Set ctl = dsg.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstInspect"
    ctl.Left = 20: ctl.Top = 35: ctl.Width = 630: ctl.Height = 350
    ctl.ColumnCount = 6
    ctl.ColumnWidths = "60;70;100;110;70;200"
    ctl.ColumnHeads = False

    ' 列头标签
    Dim headers As Variant
    headers = Array("编号", "巡检员", "巡检区域", "巡检时间", "状态", "备注")
    Dim hLeft As Variant
    hLeft = Array(20, 80, 150, 250, 360, 430)
    Dim hWid As Variant
    hWid = Array(60, 70, 100, 110, 70, 200)
    Dim hi As Long
    For hi = 0 To 5
        Set ctl = dsg.Controls.Add("Forms.Label.1")
        ctl.Name = "lblH" & hi
        ctl.Caption = headers(hi)
        ctl.Left = hLeft(hi): ctl.Top = 35: ctl.Width = hWid(hi): ctl.Height = 14
        ctl.Font.Bold = True: ctl.Font.Size = 8
    Next hi

    ' 按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNew": ctl.Caption = "新建巡检"
    ctl.Left = 20: ctl.Top = 400: ctl.Width = 90: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdEdit": ctl.Caption = "编辑"
    ctl.Left = 120: ctl.Top = 400: ctl.Width = 70: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDelete": ctl.Caption = "删除"
    ctl.Left = 200: ctl.Top = 400: ctl.Width = 70: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdRefresh": ctl.Caption = "刷新"
    ctl.Left = 280: ctl.Top = 400: ctl.Width = 70: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 560: ctl.Top = 400: ctl.Width = 90: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    lstInspect.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "巡检管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim cNormal As Long, cAbnormal As Long, cPending As Long" & vbCrLf
    c = c & "    cNormal = 0: cAbnormal = 0: cPending = 0" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        lstInspect.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "        lstInspect.List(lstInspect.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "        lstInspect.List(lstInspect.ListCount - 1, 2) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "        lstInspect.List(lstInspect.ListCount - 1, 3) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "        lstInspect.List(lstInspect.ListCount - 1, 4) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "        lstInspect.List(lstInspect.ListCount - 1, 5) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "        Dim st As String" & vbCrLf
    c = c & "        st = CStr(ws.Cells(i, 5).Value)" & vbCrLf
    c = c & "        If st = " & q & "正常" & q & " Then cNormal = cNormal + 1" & vbCrLf
    c = c & "        If st = " & q & "异常" & q & " Then cAbnormal = cAbnormal + 1" & vbCrLf
    c = c & "        If st = " & q & "待巡检" & q & " Then cPending = cPending + 1" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    lblStats.Caption = " & q & "状态统计:  正常: " & q & " & cNormal & " & q & "  |  异常: " & q & " & cAbnormal & " & q & "  |  待巡检: " & q & " & cPending & " & q & "  |  总计: " & q & " & (cNormal + cAbnormal + cPending)" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNew_Click()" & vbCrLf
    c = c & "    gEditInspectID = " & q & q & vbCrLf
    c = c & "    ShowInspectEditForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEdit_Click()" & vbCrLf
    c = c & "    If lstInspect.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条巡检记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim selID As String" & vbCrLf
    c = c & "    selID = lstInspect.List(lstInspect.ListIndex, 0)" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        Dim creator As String" & vbCrLf
    c = c & "        creator = lstInspect.List(lstInspect.ListIndex, 1)" & vbCrLf
    c = c & "        If creator <> gCurrentUser Then" & vbCrLf
    c = c & "            MsgBox " & q & "只能编辑自己创建的巡检任务！" & q & ", vbExclamation" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditInspectID = selID" & vbCrLf
    c = c & "    ShowInspectEditForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelete_Click()" & vbCrLf
    c = c & "    If lstInspect.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条巡检记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        Dim creator As String" & vbCrLf
    c = c & "        creator = lstInspect.List(lstInspect.ListIndex, 1)" & vbCrLf
    c = c & "        If creator <> gCurrentUser Then" & vbCrLf
    c = c & "            MsgBox " & q & "只能删除自己创建的巡检任务！" & q & ", vbExclamation" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该巡检记录？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        DeleteInspectByID lstInspect.List(lstInspect.ListIndex, 0)" & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdRefresh_Click()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建巡检编辑窗体 ----------

Private Sub CreateInspectEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("InspectEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "InspectEditFormName", actualName

    vbc.Properties("Caption") = "巡检任务编辑"
    vbc.Properties("Width") = 420
    vbc.Properties("Height") = 380
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 20

    ' 巡检员
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblInspector": ctl.Caption = "巡检员："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtInspector"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20
    ctl.Locked = True

    yy = yy + 32
    ' 巡检区域
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblArea": ctl.Caption = "巡检区域："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboArea"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20
    ctl.Style = 0

    yy = yy + 32
    ' 巡检时间
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTime": ctl.Caption = "巡检时间："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtTime"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20

    yy = yy + 32
    ' 巡检状态
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStatus": ctl.Caption = "巡检状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboStatus"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20
    ctl.Style = 2

    yy = yy + 32
    ' 巡检备注
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRemark": ctl.Caption = "巡检备注："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtRemark"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 70
    ctl.MultiLine = True: ctl.ScrollBars = 2

    yy = yy + 80
    ' 图片路径
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblImage": ctl.Caption = "图片上传："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtImagePath"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 200: ctl.Height = 20
    ctl.Locked = True
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdBrowse": ctl.Caption = "浏览..."
    ctl.Left = 310: ctl.Top = yy: ctl.Width = 70: ctl.Height = 20

    yy = yy + 42
    ' 保存/取消按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboArea.AddItem " & q & "A栋大堂" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "B栋大堂" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "地下车库" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "园林绿化" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "消防通道" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "设备机房" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "外围围墙" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "电梯间" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "待巡检" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "正常" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "异常" & q & vbCrLf
    c = c & "    txtInspector.Text = gCurrentUser" & vbCrLf
    c = c & "    txtTime.Text = Format(Now, " & q & "yyyy-mm-dd hh:mm" & q & ")" & vbCrLf
    c = c & "    cboStatus.ListIndex = 0" & vbCrLf
    c = c & "    If gEditInspectID <> " & q & q & " Then" & vbCrLf
    c = c & "        Me.Caption = " & q & "编辑巡检任务" & q & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Me.Caption = " & q & "新建巡检任务" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "巡检管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gEditInspectID Then" & vbCrLf
    c = c & "            txtInspector.Text = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            cboArea.Text = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            txtTime.Text = CStr(ws.Cells(i, 4).Value)" & vbCrLf
    c = c & "            cboStatus.Text = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            txtRemark.Text = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            txtImagePath.Text = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdBrowse_Click()" & vbCrLf
    c = c & "    Dim fd As Object" & vbCrLf
    c = c & "    Set fd = Application.FileDialog(1)" & vbCrLf
    c = c & "    fd.Title = " & q & "选择巡检图片" & q & vbCrLf
    c = c & "    fd.Filters.Clear" & vbCrLf
    c = c & "    fd.Filters.Add " & q & "图片文件" & q & ", " & q & "*.jpg;*.jpeg;*.png;*.bmp;*.gif" & q & vbCrLf
    c = c & "    If fd.Show = -1 Then" & vbCrLf
    c = c & "        txtImagePath.Text = fd.SelectedItems(1)" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(cboArea.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择或输入巡检区域！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(txtTime.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写巡检时间！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If gEditInspectID <> " & q & q & " Then" & vbCrLf
    c = c & "        UpdateInspectRecord gEditInspectID, txtInspector.Text, cboArea.Text, txtTime.Text, cboStatus.Text, txtRemark.Text, txtImagePath.Text" & vbCrLf
    c = c & "        MsgBox " & q & "巡检记录已更新！" & q & ", vbInformation" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        AddInspectRecord txtInspector.Text, cboArea.Text, txtTime.Text, cboStatus.Text, txtRemark.Text, txtImagePath.Text" & vbCrLf
    c = c & "        MsgBox " & q & "巡检任务已创建！" & q & ", vbInformation" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 巡检管理窗体显示 ----------

Public Sub ShowInspectManagerForm()
    Dim fName As String
    fName = GetConfigProp("InspectManagerFormName")
    If fName = "" Or Not FormExists("InspectManagerFormName") Then
        MsgBox "巡检管理窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("InspectManagerFormName")
    frm.Show
End Sub

Public Sub ShowInspectEditForm()
    Dim fName As String
    fName = GetConfigProp("InspectEditFormName")
    If fName = "" Or Not FormExists("InspectEditFormName") Then
        MsgBox "巡检编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("InspectEditFormName")
    frm.Show
End Sub

' ---------- 巡检管理数据操作函数 ----------

Private Function GenerateInspectID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("巡检管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "XJ" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateInspectID = "XJ" & Format(maxID + 1, "0000")
End Function

Public Sub AddInspectRecord(sInspector As String, sArea As String, sTime As String, sStatus As String, sRemark As String, sImage As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("巡检管理")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateInspectID()
    ws.Cells(newRow, 2).Value = sInspector
    ws.Cells(newRow, 3).Value = sArea
    ws.Cells(newRow, 4).Value = sTime
    ws.Cells(newRow, 5).Value = sStatus
    ws.Cells(newRow, 6).Value = sRemark
    ws.Cells(newRow, 7).Value = sImage
    ws.Cells(newRow, 8).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Columns("A:H").AutoFit
End Sub

Public Sub UpdateInspectRecord(iID As String, sInspector As String, sArea As String, sTime As String, sStatus As String, sRemark As String, sImage As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("巡检管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = iID Then
            ws.Cells(i, 2).Value = sInspector
            ws.Cells(i, 3).Value = sArea
            ws.Cells(i, 4).Value = sTime
            ws.Cells(i, 5).Value = sStatus
            ws.Cells(i, 6).Value = sRemark
            ws.Cells(i, 7).Value = sImage
            Exit For
        End If
    Next i
    ws.Columns("A:H").AutoFit
End Sub

Public Sub DeleteInspectByID(iID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("巡检管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = iID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

' ---------- 创建卫生管理表 ----------

Private Sub CreateHygieneSheet()
    Dim ws As Worksheet
    If SheetExists("卫生管理") Then
        Set ws = ThisWorkbook.Sheets("卫生管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "卫生管理"
    End If
    With ws
        .Range("A1").Value = "任务编号"
        .Range("B1").Value = "保洁任务"
        .Range("C1").Value = "保洁区域"
        .Range("D1").Value = "保洁员"
        .Range("E1").Value = "保洁时间"
        .Range("F1").Value = "任务状态"
        .Range("G1").Value = "备注"
        .Range("H1").Value = "图片路径"
        .Range("I1").Value = "创建时间"
        .Range("J1").Value = "完成时间"
        .Range("A1:J1").Font.Bold = True
        .Columns("A:J").AutoFit
    End With
End Sub

' ---------- 创建卫生管理窗体 ----------

Private Sub CreateHygieneManagerForm()
    Dim oldForm As String
    oldForm = GetConfigProp("HygieneManagerFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "HygieneManagerFormName", actualName

    vbc.Properties("Caption") = "卫生管理"
    vbc.Properties("Width") = 700
    vbc.Properties("Height") = 520
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 状态统计标签
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStats": ctl.Caption = "状态统计: 加载中..."
    ctl.Left = 20: ctl.Top = 10: ctl.Width = 550: ctl.Height = 18
    ctl.Font.Size = 10: ctl.Font.Bold = True

    ' 保洁任务列表
    Set ctl = dsg.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstHygiene"
    ctl.Left = 20: ctl.Top = 35: ctl.Width = 650: ctl.Height = 370
    ctl.ColumnCount = 7
    ctl.ColumnWidths = "60;110;90;70;100;60;130"
    ctl.ColumnHeads = False

    ' 按钮行
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNew": ctl.Caption = "新建任务"
    ctl.Left = 20: ctl.Top = 420: ctl.Width = 90: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdView": ctl.Caption = "查看"
    ctl.Left = 120: ctl.Top = 420: ctl.Width = 70: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdEdit": ctl.Caption = "编辑"
    ctl.Left = 200: ctl.Top = 420: ctl.Width = 70: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDelete": ctl.Caption = "删除"
    ctl.Left = 280: ctl.Top = 420: ctl.Width = 70: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdRefresh": ctl.Caption = "刷新"
    ctl.Left = 360: ctl.Top = 420: ctl.Width = 70: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 580: ctl.Top = 420: ctl.Width = 90: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    lstHygiene.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "卫生管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim cDone As Long, cUndone As Long, cPending As Long" & vbCrLf
    c = c & "    cDone = 0: cUndone = 0: cPending = 0" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        lstHygiene.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "        lstHygiene.List(lstHygiene.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "        lstHygiene.List(lstHygiene.ListCount - 1, 2) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "        lstHygiene.List(lstHygiene.ListCount - 1, 3) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "        lstHygiene.List(lstHygiene.ListCount - 1, 4) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "        lstHygiene.List(lstHygiene.ListCount - 1, 5) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "        lstHygiene.List(lstHygiene.ListCount - 1, 6) = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "        Dim st As String" & vbCrLf
    c = c & "        st = CStr(ws.Cells(i, 6).Value)" & vbCrLf
    c = c & "        If st = " & q & "已完成" & q & " Then cDone = cDone + 1" & vbCrLf
    c = c & "        If st = " & q & "未完成" & q & " Then cUndone = cUndone + 1" & vbCrLf
    c = c & "        If st = " & q & "待完成" & q & " Then cPending = cPending + 1" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    lblStats.Caption = " & q & "状态统计:  已完成: " & q & " & cDone & " & q & "  |  未完成: " & q & " & cUndone & " & q & "  |  待完成: " & q & " & cPending & " & q & "  |  总计: " & q & " & (cDone + cUndone + cPending)" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNew_Click()" & vbCrLf
    c = c & "    gEditHygieneID = " & q & q & vbCrLf
    c = c & "    ShowHygieneEditForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdView_Click()" & vbCrLf
    c = c & "    If lstHygiene.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条保洁记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewHygieneID = lstHygiene.List(lstHygiene.ListIndex, 0)" & vbCrLf
    c = c & "    ShowHygieneViewForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEdit_Click()" & vbCrLf
    c = c & "    If lstHygiene.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条保洁记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim selID As String" & vbCrLf
    c = c & "    selID = lstHygiene.List(lstHygiene.ListIndex, 0)" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        Dim creator As String" & vbCrLf
    c = c & "        creator = lstHygiene.List(lstHygiene.ListIndex, 3)" & vbCrLf
    c = c & "        If creator <> gCurrentUser Then" & vbCrLf
    c = c & "            MsgBox " & q & "只能编辑自己创建的保洁任务！" & q & ", vbExclamation" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditHygieneID = selID" & vbCrLf
    c = c & "    ShowHygieneEditForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelete_Click()" & vbCrLf
    c = c & "    If lstHygiene.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条保洁记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        Dim creator As String" & vbCrLf
    c = c & "        creator = lstHygiene.List(lstHygiene.ListIndex, 3)" & vbCrLf
    c = c & "        If creator <> gCurrentUser Then" & vbCrLf
    c = c & "            MsgBox " & q & "只能删除自己创建的保洁任务！" & q & ", vbExclamation" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该保洁记录？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        DeleteHygieneByID lstHygiene.List(lstHygiene.ListIndex, 0)" & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdRefresh_Click()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建卫生编辑窗体 ----------

Private Sub CreateHygieneEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("HygieneEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "HygieneEditFormName", actualName

    vbc.Properties("Caption") = "保洁任务编辑"
    vbc.Properties("Width") = 420
    vbc.Properties("Height") = 400
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 20

    ' 保洁任务
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTask": ctl.Caption = "保洁任务："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtTask"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20

    yy = yy + 32
    ' 保洁区域
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblArea": ctl.Caption = "保洁区域："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboArea"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20
    ctl.Style = 0

    yy = yy + 32
    ' 保洁员
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCleaner": ctl.Caption = "保洁员："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtCleaner"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20
    ctl.Locked = True

    yy = yy + 32
    ' 保洁时间
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTime": ctl.Caption = "保洁时间："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtTime"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20

    yy = yy + 32
    ' 任务状态
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStatus": ctl.Caption = "任务状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboStatus"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20
    ctl.Style = 2

    yy = yy + 32
    ' 备注
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRemark": ctl.Caption = "备注："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtRemark"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 60
    ctl.MultiLine = True: ctl.ScrollBars = 2

    yy = yy + 70
    ' 图片路径
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblImage": ctl.Caption = "图片上传："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtImagePath"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 200: ctl.Height = 20
    ctl.Locked = True
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdBrowse": ctl.Caption = "浏览..."
    ctl.Left = 310: ctl.Top = yy: ctl.Width = 70: ctl.Height = 20

    yy = yy + 40
    ' 保存/取消
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboArea.AddItem " & q & "A栋大堂" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "B栋大堂" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "地下车库" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "园林绿化区" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "楼道走廊" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "公共卫生间" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "电梯间" & q & vbCrLf
    c = c & "    cboArea.AddItem " & q & "办公区域" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "待完成" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "未完成" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "已完成" & q & vbCrLf
    c = c & "    txtCleaner.Text = gCurrentUser" & vbCrLf
    c = c & "    txtTime.Text = Format(Now, " & q & "yyyy-mm-dd hh:mm" & q & ")" & vbCrLf
    c = c & "    cboStatus.ListIndex = 0" & vbCrLf
    c = c & "    If gEditHygieneID <> " & q & q & " Then" & vbCrLf
    c = c & "        Me.Caption = " & q & "编辑保洁任务" & q & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Me.Caption = " & q & "新建保洁任务" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "卫生管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gEditHygieneID Then" & vbCrLf
    c = c & "            txtTask.Text = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            cboArea.Text = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            txtCleaner.Text = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            txtTime.Text = CStr(ws.Cells(i, 5).Value)" & vbCrLf
    c = c & "            cboStatus.Text = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            txtRemark.Text = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            txtImagePath.Text = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdBrowse_Click()" & vbCrLf
    c = c & "    Dim fd As Object" & vbCrLf
    c = c & "    Set fd = Application.FileDialog(1)" & vbCrLf
    c = c & "    fd.Title = " & q & "选择保洁图片" & q & vbCrLf
    c = c & "    fd.Filters.Clear" & vbCrLf
    c = c & "    fd.Filters.Add " & q & "图片文件" & q & ", " & q & "*.jpg;*.jpeg;*.png;*.bmp;*.gif" & q & vbCrLf
    c = c & "    If fd.Show = -1 Then" & vbCrLf
    c = c & "        txtImagePath.Text = fd.SelectedItems(1)" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(txtTask.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写保洁任务！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(cboArea.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择或输入保洁区域！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If gEditHygieneID <> " & q & q & " Then" & vbCrLf
    c = c & "        UpdateHygieneRecord gEditHygieneID, txtTask.Text, cboArea.Text, txtCleaner.Text, txtTime.Text, cboStatus.Text, txtRemark.Text, txtImagePath.Text" & vbCrLf
    c = c & "        MsgBox " & q & "保洁任务已更新！" & q & ", vbInformation" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        AddHygieneRecord txtTask.Text, cboArea.Text, txtCleaner.Text, txtTime.Text, cboStatus.Text, txtRemark.Text, txtImagePath.Text" & vbCrLf
    c = c & "        MsgBox " & q & "保洁任务已创建！" & q & ", vbInformation" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建卫生查看窗体 ----------

Private Sub CreateHygieneViewForm()
    Dim oldForm As String
    oldForm = GetConfigProp("HygieneViewFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "HygieneViewFormName", actualName

    vbc.Properties("Caption") = "保洁任务详情"
    vbc.Properties("Width") = 420
    vbc.Properties("Height") = 360
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 20

    ' 任务名称
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTaskCap": ctl.Caption = "保洁任务："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTaskVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAreaCap": ctl.Caption = "保洁区域："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAreaVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCleanerCap": ctl.Caption = "保洁员："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCleanerVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTimeCap": ctl.Caption = "保洁时间："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTimeVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStCap": ctl.Caption = "当前状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20
    ctl.Font.Size = 11: ctl.Font.Bold = True

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRemarkCap": ctl.Caption = "备注："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtRemarkV"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 50
    ctl.MultiLine = True: ctl.Locked = True: ctl.ScrollBars = 2

    yy = yy + 60
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblImgCap": ctl.Caption = "图片路径："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblImgVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 18

    yy = yy + 36
    ' 操作按钮: 进行中、已完成、关闭
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdInProgress"
    ctl.Caption = "进行中"
    ctl.Left = 30: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(255, 165, 0)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdComplete"
    ctl.Caption = "已完成"
    ctl.Left = 150: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(60, 179, 113)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose"
    ctl.Caption = "关闭"
    ctl.Left = 270: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    If gViewHygieneID = " & q & q & " Then Unload Me: Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "卫生管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gViewHygieneID Then" & vbCrLf
    c = c & "            lblTaskVal.Caption = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lblAreaVal.Caption = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            lblCleanerVal.Caption = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lblTimeVal.Caption = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lblStVal.Caption = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            txtRemarkV.Text = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            lblImgVal.Caption = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            Dim st As String" & vbCrLf
    c = c & "            st = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            If st = " & q & "已完成" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(60, 179, 113)" & vbCrLf
    c = c & "                cmdInProgress.Enabled = False" & vbCrLf
    c = c & "                cmdComplete.Enabled = False" & vbCrLf
    c = c & "            ElseIf st = " & q & "未完成" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(220, 80, 80)" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(255, 140, 0)" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdInProgress_Click()" & vbCrLf
    c = c & "    UpdateHygieneStatus gViewHygieneID, " & q & "未完成" & q & vbCrLf
    c = c & "    MsgBox " & q & "状态已更新为【进行中/未完成】！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdComplete_Click()" & vbCrLf
    c = c & "    UpdateHygieneStatus gViewHygieneID, " & q & "已完成" & q & vbCrLf
    c = c & "    MsgBox " & q & "任务已标记为【已完成】！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 卫生管理窗体显示 ----------

Public Sub ShowHygieneManagerForm()
    Dim fName As String
    fName = GetConfigProp("HygieneManagerFormName")
    If fName = "" Or Not FormExists("HygieneManagerFormName") Then
        MsgBox "卫生管理窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("HygieneManagerFormName")
    frm.Show
End Sub

Public Sub ShowHygieneEditForm()
    Dim fName As String
    fName = GetConfigProp("HygieneEditFormName")
    If fName = "" Or Not FormExists("HygieneEditFormName") Then
        MsgBox "保洁编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("HygieneEditFormName")
    frm.Show
End Sub

Public Sub ShowHygieneViewForm()
    Dim fName As String
    fName = GetConfigProp("HygieneViewFormName")
    If fName = "" Or Not FormExists("HygieneViewFormName") Then
        MsgBox "保洁查看窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("HygieneViewFormName")
    frm.Show
End Sub

' ---------- 卫生管理数据操作函数 ----------

Private Function GenerateHygieneID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("卫生管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "WS" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateHygieneID = "WS" & Format(maxID + 1, "0000")
End Function

Public Sub AddHygieneRecord(sTask As String, sArea As String, sCleaner As String, sTime As String, sStatus As String, sRemark As String, sImage As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("卫生管理")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateHygieneID()
    ws.Cells(newRow, 2).Value = sTask
    ws.Cells(newRow, 3).Value = sArea
    ws.Cells(newRow, 4).Value = sCleaner
    ws.Cells(newRow, 5).Value = sTime
    ws.Cells(newRow, 6).Value = sStatus
    ws.Cells(newRow, 7).Value = sRemark
    ws.Cells(newRow, 8).Value = sImage
    ws.Cells(newRow, 9).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(newRow, 10).Value = ""
    ws.Columns("A:J").AutoFit
End Sub

Public Sub UpdateHygieneRecord(hID As String, sTask As String, sArea As String, sCleaner As String, sTime As String, sStatus As String, sRemark As String, sImage As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("卫生管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = hID Then
            ws.Cells(i, 2).Value = sTask
            ws.Cells(i, 3).Value = sArea
            ws.Cells(i, 4).Value = sCleaner
            ws.Cells(i, 5).Value = sTime
            ws.Cells(i, 6).Value = sStatus
            ws.Cells(i, 7).Value = sRemark
            ws.Cells(i, 8).Value = sImage
            Exit For
        End If
    Next i
    ws.Columns("A:J").AutoFit
End Sub

Public Sub UpdateHygieneStatus(hID As String, newStatus As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("卫生管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = hID Then
            ws.Cells(i, 6).Value = newStatus
            If newStatus = "已完成" Then
                ws.Cells(i, 10).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub DeleteHygieneByID(hID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("卫生管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = hID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

' ---------- 创建个人行程表 ----------

Private Sub CreateTripSheet()
    Dim ws As Worksheet
    If SheetExists("个人行程") Then
        Set ws = ThisWorkbook.Sheets("个人行程")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "个人行程"
    End If
    With ws
        .Range("A1").Value = "行程编号"
        .Range("B1").Value = "用户"
        .Range("C1").Value = "行程标题"
        .Range("D1").Value = "行程内容"
        .Range("E1").Value = "开始时间"
        .Range("F1").Value = "结束时间"
        .Range("G1").Value = "行程状态"
        .Range("H1").Value = "创建时间"
        .Range("I1").Value = "完成时间"
        .Range("A1:I1").Font.Bold = True
        .Columns("A:I").AutoFit
    End With
End Sub

' ---------- 创建个人行程管理窗体 ----------

Private Sub CreateTripManagerForm()
    Dim oldForm As String
    oldForm = GetConfigProp("TripManagerFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "TripManagerFormName", actualName

    vbc.Properties("Caption") = "个人行程"
    vbc.Properties("Width") = 680
    vbc.Properties("Height") = 520
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' MultiPage控件 - 两个标签页
    Set ctl = dsg.Controls.Add("Forms.MultiPage.1")
    ctl.Name = "mpTrip"
    ctl.Left = 10: ctl.Top = 10: ctl.Width = 650: ctl.Height = 440
    ctl.Pages(0).Caption = "我的行程"
    ctl.Pages.Add
    ctl.Pages(1).Caption = "行程查询"

    ' --- 第一页: 我的行程 (待办/办理中) ---
    Dim pg As Object
    Set pg = ctl.Pages(0)

    Dim subCtl As Object
    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblMyTitle": subCtl.Caption = "待办行程列表"
    subCtl.Left = 10: subCtl.Top = 8: subCtl.Width = 200: subCtl.Height = 16
    subCtl.Font.Bold = True

    Set subCtl = pg.Controls.Add("Forms.ListBox.1")
    subCtl.Name = "lstMyTrips"
    subCtl.Left = 10: subCtl.Top = 28: subCtl.Width = 620: subCtl.Height = 300
    subCtl.ColumnCount = 6
    subCtl.ColumnWidths = "60;120;120;80;80;120"

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdNew": subCtl.Caption = "添加行程"
    subCtl.Left = 10: subCtl.Top = 340: subCtl.Width = 90: subCtl.Height = 28

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdView": subCtl.Caption = "查看"
    subCtl.Left = 110: subCtl.Top = 340: subCtl.Width = 70: subCtl.Height = 28

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdEdit": subCtl.Caption = "编辑"
    subCtl.Left = 190: subCtl.Top = 340: subCtl.Width = 70: subCtl.Height = 28

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdDelete": subCtl.Caption = "删除"
    subCtl.Left = 270: subCtl.Top = 340: subCtl.Width = 70: subCtl.Height = 28

    ' --- 第二页: 行程查询 (所有行程) ---
    Set pg = ctl.Pages(1)

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblAllTitle": subCtl.Caption = "全部行程记录（含已完成）"
    subCtl.Left = 10: subCtl.Top = 8: subCtl.Width = 300: subCtl.Height = 16
    subCtl.Font.Bold = True

    Set subCtl = pg.Controls.Add("Forms.ListBox.1")
    subCtl.Name = "lstAllTrips"
    subCtl.Left = 10: subCtl.Top = 28: subCtl.Width = 620: subCtl.Height = 340
    subCtl.ColumnCount = 6
    subCtl.ColumnWidths = "60;120;120;80;80;120"

    ' 关闭按钮（主窗体）
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 570: ctl.Top = 460: ctl.Width = 90: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    LoadMyTrips" & vbCrLf
    c = c & "    LoadAllTrips" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadMyTrips()" & vbCrLf
    c = c & "    lstMyTrips.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "个人行程" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 2).Value) = gCurrentUser Then" & vbCrLf
    c = c & "            Dim st As String" & vbCrLf
    c = c & "            st = CStr(ws.Cells(i, 7).Value)" & vbCrLf
    c = c & "            If st = " & q & "待办" & q & " Or st = " & q & "办理中" & q & " Then" & vbCrLf
    c = c & "                lstMyTrips.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "                lstMyTrips.List(lstMyTrips.ListCount - 1, 1) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "                lstMyTrips.List(lstMyTrips.ListCount - 1, 2) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "                lstMyTrips.List(lstMyTrips.ListCount - 1, 3) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "                lstMyTrips.List(lstMyTrips.ListCount - 1, 4) = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "                lstMyTrips.List(lstMyTrips.ListCount - 1, 5) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadAllTrips()" & vbCrLf
    c = c & "    lstAllTrips.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "个人行程" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 2).Value) = gCurrentUser Then" & vbCrLf
    c = c & "            lstAllTrips.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lstAllTrips.List(lstAllTrips.ListCount - 1, 1) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            lstAllTrips.List(lstAllTrips.ListCount - 1, 2) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lstAllTrips.List(lstAllTrips.ListCount - 1, 3) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            lstAllTrips.List(lstAllTrips.ListCount - 1, 4) = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            lstAllTrips.List(lstAllTrips.ListCount - 1, 5) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNew_Click()" & vbCrLf
    c = c & "    gEditTripID = " & q & q & vbCrLf
    c = c & "    ShowTripEditForm" & vbCrLf
    c = c & "    LoadMyTrips" & vbCrLf
    c = c & "    LoadAllTrips" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdView_Click()" & vbCrLf
    c = c & "    If lstMyTrips.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条行程！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewTripID = lstMyTrips.List(lstMyTrips.ListIndex, 0)" & vbCrLf
    c = c & "    ShowTripViewForm" & vbCrLf
    c = c & "    LoadMyTrips" & vbCrLf
    c = c & "    LoadAllTrips" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEdit_Click()" & vbCrLf
    c = c & "    If lstMyTrips.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条行程！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditTripID = lstMyTrips.List(lstMyTrips.ListIndex, 0)" & vbCrLf
    c = c & "    ShowTripEditForm" & vbCrLf
    c = c & "    LoadMyTrips" & vbCrLf
    c = c & "    LoadAllTrips" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelete_Click()" & vbCrLf
    c = c & "    If lstMyTrips.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条行程！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该行程？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        DeleteTripByID lstMyTrips.List(lstMyTrips.ListIndex, 0)" & vbCrLf
    c = c & "        LoadMyTrips" & vbCrLf
    c = c & "        LoadAllTrips" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建行程编辑窗体 ----------

Private Sub CreateTripEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("TripEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "TripEditFormName", actualName

    vbc.Properties("Caption") = "行程编辑"
    vbc.Properties("Width") = 420
    vbc.Properties("Height") = 320
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 20

    ' 行程标题
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTitle": ctl.Caption = "行程标题："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtTitle"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20

    yy = yy + 32
    ' 行程内容
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblContent": ctl.Caption = "行程内容："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtContent"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 60
    ctl.MultiLine = True: ctl.ScrollBars = 2

    yy = yy + 70
    ' 开始时间
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStart": ctl.Caption = "开始时间："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtStart"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20

    yy = yy + 32
    ' 结束时间
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblEnd": ctl.Caption = "结束时间："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtEnd2"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20

    yy = yy + 40
    ' 保存/取消
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    txtStart.Text = Format(Now, " & q & "yyyy-mm-dd hh:mm" & q & ")" & vbCrLf
    c = c & "    txtEnd2.Text = Format(Now + 1, " & q & "yyyy-mm-dd hh:mm" & q & ")" & vbCrLf
    c = c & "    If gEditTripID <> " & q & q & " Then" & vbCrLf
    c = c & "        Me.Caption = " & q & "编辑行程" & q & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Me.Caption = " & q & "添加行程" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "个人行程" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gEditTripID Then" & vbCrLf
    c = c & "            txtTitle.Text = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            txtContent.Text = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            txtStart.Text = CStr(ws.Cells(i, 5).Value)" & vbCrLf
    c = c & "            txtEnd2.Text = CStr(ws.Cells(i, 6).Value)" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(txtTitle.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写行程标题！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If gEditTripID <> " & q & q & " Then" & vbCrLf
    c = c & "        UpdateTripRecord gEditTripID, txtTitle.Text, txtContent.Text, txtStart.Text, txtEnd2.Text" & vbCrLf
    c = c & "        MsgBox " & q & "行程已更新！" & q & ", vbInformation" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        AddTripRecord gCurrentUser, txtTitle.Text, txtContent.Text, txtStart.Text, txtEnd2.Text" & vbCrLf
    c = c & "        MsgBox " & q & "行程已添加！" & q & ", vbInformation" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建行程查看窗体 ----------

Private Sub CreateTripViewForm()
    Dim oldForm As String
    oldForm = GetConfigProp("TripViewFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "TripViewFormName", actualName

    vbc.Properties("Caption") = "行程详情"
    vbc.Properties("Width") = 420
    vbc.Properties("Height") = 320
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 20

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTitleCap": ctl.Caption = "行程标题："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTitleVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblContentCap": ctl.Caption = "行程内容："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtContentV"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 55
    ctl.MultiLine = True: ctl.Locked = True: ctl.ScrollBars = 2

    yy = yy + 62
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTimeCap": ctl.Caption = "时间："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTimeVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 290: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStCap": ctl.Caption = "当前状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20
    ctl.Font.Size = 11: ctl.Font.Bold = True

    yy = yy + 40
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdInProgress"
    ctl.Caption = "办理中"
    ctl.Left = 30: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(255, 165, 0)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdComplete"
    ctl.Caption = "已完成"
    ctl.Left = 150: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30
    ctl.BackColor = RGB(60, 179, 113)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose"
    ctl.Caption = "关闭"
    ctl.Left = 270: ctl.Top = yy: ctl.Width = 100: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    If gViewTripID = " & q & q & " Then Unload Me: Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "个人行程" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gViewTripID Then" & vbCrLf
    c = c & "            lblTitleVal.Caption = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            txtContentV.Text = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lblTimeVal.Caption = ws.Cells(i, 5).Value & " & q & " ~ " & q & " & ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            lblStVal.Caption = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            Dim st As String" & vbCrLf
    c = c & "            st = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            If st = " & q & "已完成" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(60, 179, 113)" & vbCrLf
    c = c & "                cmdInProgress.Enabled = False" & vbCrLf
    c = c & "                cmdComplete.Enabled = False" & vbCrLf
    c = c & "            ElseIf st = " & q & "办理中" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(255, 140, 0)" & vbCrLf
    c = c & "                cmdInProgress.Enabled = False" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(0, 0, 0)" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdInProgress_Click()" & vbCrLf
    c = c & "    UpdateTripStatus gViewTripID, " & q & "办理中" & q & vbCrLf
    c = c & "    MsgBox " & q & "行程状态已更新为【办理中】！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdComplete_Click()" & vbCrLf
    c = c & "    UpdateTripStatus gViewTripID, " & q & "已完成" & q & vbCrLf
    c = c & "    MsgBox " & q & "行程已标记为【已完成】！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 个人行程窗体显示 ----------

Public Sub ShowTripManagerForm()
    Dim fName As String
    fName = GetConfigProp("TripManagerFormName")
    If fName = "" Or Not FormExists("TripManagerFormName") Then
        MsgBox "个人行程窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("TripManagerFormName")
    frm.Show
End Sub

Public Sub ShowTripEditForm()
    Dim fName As String
    fName = GetConfigProp("TripEditFormName")
    If fName = "" Or Not FormExists("TripEditFormName") Then
        MsgBox "行程编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("TripEditFormName")
    frm.Show
End Sub

Public Sub ShowTripViewForm()
    Dim fName As String
    fName = GetConfigProp("TripViewFormName")
    If fName = "" Or Not FormExists("TripViewFormName") Then
        MsgBox "行程查看窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("TripViewFormName")
    frm.Show
End Sub

' ---------- 个人行程数据操作函数 ----------

Private Function GenerateTripID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("个人行程")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "XC" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateTripID = "XC" & Format(maxID + 1, "0000")
End Function

Public Sub AddTripRecord(sUser As String, sTitle As String, sContent As String, sStart As String, sEnd2 As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("个人行程")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateTripID()
    ws.Cells(newRow, 2).Value = sUser
    ws.Cells(newRow, 3).Value = sTitle
    ws.Cells(newRow, 4).Value = sContent
    ws.Cells(newRow, 5).Value = sStart
    ws.Cells(newRow, 6).Value = sEnd2
    ws.Cells(newRow, 7).Value = "待办"
    ws.Cells(newRow, 8).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(newRow, 9).Value = ""
    ws.Columns("A:I").AutoFit
End Sub

Public Sub UpdateTripRecord(tID As String, sTitle As String, sContent As String, sStart As String, sEnd2 As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("个人行程")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = tID Then
            ws.Cells(i, 3).Value = sTitle
            ws.Cells(i, 4).Value = sContent
            ws.Cells(i, 5).Value = sStart
            ws.Cells(i, 6).Value = sEnd2
            Exit For
        End If
    Next i
    ws.Columns("A:I").AutoFit
End Sub

Public Sub UpdateTripStatus(tID As String, newStatus As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("个人行程")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = tID Then
            ws.Cells(i, 7).Value = newStatus
            If newStatus = "已完成" Then
                ws.Cells(i, 9).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub DeleteTripByID(tID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("个人行程")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = tID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

' ---------- 创建考勤管理表 ----------

Private Sub CreateAttendSheet()
    Dim ws As Worksheet
    If SheetExists("考勤管理") Then
        Set ws = ThisWorkbook.Sheets("考勤管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "考勤管理"
    End If
    With ws
        .Range("A1").Value = "记录编号"
        .Range("B1").Value = "用户"
        .Range("C1").Value = "日期"
        .Range("D1").Value = "签到时间"
        .Range("E1").Value = "签退时间"
        .Range("F1").Value = "考勤状态"
        .Range("G1").Value = "备注"
        .Range("A1:G1").Font.Bold = True
        .Columns("A:G").AutoFit
    End With
End Sub

' ---------- 创建考勤窗体 ----------

Private Sub CreateAttendForm()
    Dim oldForm As String
    oldForm = GetConfigProp("AttendFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "AttendFormName", actualName

    vbc.Properties("Caption") = "个人考勤"
    vbc.Properties("Width") = 650
    vbc.Properties("Height") = 500
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' MultiPage - 两个标签页
    Set ctl = dsg.Controls.Add("Forms.MultiPage.1")
    ctl.Name = "mpAttend"
    ctl.Left = 10: ctl.Top = 10: ctl.Width = 620: ctl.Height = 420
    ctl.Pages(0).Caption = "签到/签退"
    ctl.Pages.Add
    ctl.Pages(1).Caption = "考勤统计查询"

    ' --- 第一页: 签到/签退 ---
    Dim pg As Object
    Set pg = ctl.Pages(0)

    Dim subCtl As Object
    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblDate": subCtl.Caption = "日期加载中..."
    subCtl.Left = 20: subCtl.Top = 20: subCtl.Width = 400: subCtl.Height = 24
    subCtl.Font.Size = 14: subCtl.Font.Bold = True

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblTime": subCtl.Caption = ""
    subCtl.Left = 20: subCtl.Top = 50: subCtl.Width = 300: subCtl.Height = 22
    subCtl.Font.Size = 12

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblSignInTime": subCtl.Caption = "签到时间: --"
    subCtl.Left = 20: subCtl.Top = 90: subCtl.Width = 250: subCtl.Height = 20
    subCtl.Font.Size = 11

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblSignOutTime": subCtl.Caption = "签退时间: --"
    subCtl.Left = 20: subCtl.Top = 118: subCtl.Width = 250: subCtl.Height = 20
    subCtl.Font.Size = 11

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblStatus": subCtl.Caption = "今日状态: 未签到"
    subCtl.Left = 20: subCtl.Top = 150: subCtl.Width = 300: subCtl.Height = 22
    subCtl.Font.Size = 11: subCtl.Font.Bold = True

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdSignIn": subCtl.Caption = "签 到"
    subCtl.Left = 40: subCtl.Top = 200: subCtl.Width = 140: subCtl.Height = 50
    subCtl.BackColor = RGB(60, 179, 113)
    subCtl.Font.Size = 14: subCtl.Font.Bold = True

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdSignOut": subCtl.Caption = "签 退"
    subCtl.Left = 220: subCtl.Top = 200: subCtl.Width = 140: subCtl.Height = 50
    subCtl.BackColor = RGB(220, 80, 80)
    subCtl.Font.Size = 14: subCtl.Font.Bold = True

    ' --- 第二页: 考勤统计查询 ---
    Set pg = ctl.Pages(1)

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblQueryMonth": subCtl.Caption = "查询月份(yyyy-mm)："
    subCtl.Left = 10: subCtl.Top = 10: subCtl.Width = 130: subCtl.Height = 18

    Set subCtl = pg.Controls.Add("Forms.TextBox.1")
    subCtl.Name = "txtMonth"
    subCtl.Left = 145: subCtl.Top = 8: subCtl.Width = 80: subCtl.Height = 20

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdQuery": subCtl.Caption = "查询"
    subCtl.Left = 235: subCtl.Top = 8: subCtl.Width = 60: subCtl.Height = 22

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblSummary": subCtl.Caption = "统计: --"
    subCtl.Left = 310: subCtl.Top = 10: subCtl.Width = 280: subCtl.Height = 18
    subCtl.Font.Bold = True

    Set subCtl = pg.Controls.Add("Forms.ListBox.1")
    subCtl.Name = "lstAttend"
    subCtl.Left = 10: subCtl.Top = 38: subCtl.Width = 590: subCtl.Height = 300
    subCtl.ColumnCount = 6
    subCtl.ColumnWidths = "60;80;90;90;70;80"

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdEditRec": subCtl.Caption = "编辑"
    subCtl.Left = 10: subCtl.Top = 348: subCtl.Width = 70: subCtl.Height = 26

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdDeleteRec": subCtl.Caption = "删除"
    subCtl.Left = 90: subCtl.Top = 348: subCtl.Width = 70: subCtl.Height = 26

    ' 关闭按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 540: ctl.Top = 440: ctl.Width = 90: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    lblDate.Caption = " & q & "今天: " & q & " & Format(Date, " & q & "yyyy年mm月dd日 (ddd)" & q & ")" & vbCrLf
    c = c & "    lblTime.Caption = " & q & "当前时间: " & q & " & Format(Now, " & q & "hh:mm:ss" & q & ")" & vbCrLf
    c = c & "    txtMonth.Text = Format(Date, " & q & "yyyy-mm" & q & ")" & vbCrLf
    c = c & "    LoadTodayStatus" & vbCrLf
    c = c & "    QueryAttend" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadTodayStatus()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "考勤管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim todayStr As String" & vbCrLf
    c = c & "    todayStr = Format(Date, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    Dim found As Boolean" & vbCrLf
    c = c & "    found = False" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim cellDate As String" & vbCrLf
    c = c & "        If IsDate(ws.Cells(i, 3).Value) Then" & vbCrLf
    c = c & "            cellDate = Format(ws.Cells(i, 3).Value, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "        Else" & vbCrLf
    c = c & "            cellDate = CStr(ws.Cells(i, 3).Value)" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 2).Value) = gCurrentUser And cellDate = todayStr Then" & vbCrLf
    c = c & "            found = True" & vbCrLf
    c = c & "            lblSignInTime.Caption = " & q & "签到时间: " & q & " & Format(ws.Cells(i, 4).Value, " & q & "hh:mm:ss" & q & ")" & vbCrLf
    c = c & "            If Trim(CStr(ws.Cells(i, 5).Value)) <> " & q & q & " Then" & vbCrLf
    c = c & "                lblSignOutTime.Caption = " & q & "签退时间: " & q & " & Format(ws.Cells(i, 5).Value, " & q & "hh:mm:ss" & q & ")" & vbCrLf
    c = c & "                lblStatus.Caption = " & q & "今日状态: " & q & " & ws.Cells(i, 6).Value" & vbCrLf
    c = c & "                cmdSignIn.Enabled = False" & vbCrLf
    c = c & "                cmdSignOut.Enabled = False" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lblSignOutTime.Caption = " & q & "签退时间: --" & q & vbCrLf
    c = c & "                lblStatus.Caption = " & q & "今日状态: 已签到" & q & vbCrLf
    c = c & "                cmdSignIn.Enabled = False" & vbCrLf
    c = c & "                cmdSignOut.Enabled = True" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    If Not found Then" & vbCrLf
    c = c & "        lblSignInTime.Caption = " & q & "签到时间: --" & q & vbCrLf
    c = c & "        lblSignOutTime.Caption = " & q & "签退时间: --" & q & vbCrLf
    c = c & "        lblStatus.Caption = " & q & "今日状态: 未签到" & q & vbCrLf
    c = c & "        cmdSignIn.Enabled = True" & vbCrLf
    c = c & "        cmdSignOut.Enabled = False" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSignIn_Click()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "考勤管理" & q & ")" & vbCrLf
    c = c & "    Dim newRow As Long" & vbCrLf
    c = c & "    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    c = c & "    ws.Cells(newRow, 1).Value = GenerateAttendID()" & vbCrLf
    c = c & "    ws.Cells(newRow, 2).Value = gCurrentUser" & vbCrLf
    c = c & "    ws.Cells(newRow, 3).Value = Format(Date, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    ws.Cells(newRow, 4).Value = Format(Now, " & q & "hh:mm:ss" & q & ")" & vbCrLf
    c = c & "    ws.Cells(newRow, 5).Value = " & q & q & vbCrLf
    c = c & "    ws.Cells(newRow, 6).Value = " & q & "出勤" & q & vbCrLf
    c = c & "    ws.Cells(newRow, 7).Value = " & q & q & vbCrLf
    c = c & "    Dim signTime As Date" & vbCrLf
    c = c & "    signTime = Now" & vbCrLf
    c = c & "    If Hour(signTime) >= 9 And Minute(signTime) > 0 Then" & vbCrLf
    c = c & "        ws.Cells(newRow, 6).Value = " & q & "迟到" & q & vbCrLf
    c = c & "        ws.Cells(newRow, 7).Value = " & q & "迟到签到" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    ws.Columns(" & q & "A:G" & q & ").AutoFit" & vbCrLf
    c = c & "    MsgBox " & q & "签到成功！时间: " & q & " & Format(Now, " & q & "hh:mm:ss" & q & "), vbInformation" & vbCrLf
    c = c & "    LoadTodayStatus" & vbCrLf
    c = c & "    QueryAttend" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSignOut_Click()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "考勤管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim todayStr As String" & vbCrLf
    c = c & "    todayStr = Format(Date, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim cellDate2 As String" & vbCrLf
    c = c & "        If IsDate(ws.Cells(i, 3).Value) Then" & vbCrLf
    c = c & "            cellDate2 = Format(ws.Cells(i, 3).Value, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "        Else" & vbCrLf
    c = c & "            cellDate2 = CStr(ws.Cells(i, 3).Value)" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 2).Value) = gCurrentUser And cellDate2 = todayStr Then" & vbCrLf
    c = c & "            ws.Cells(i, 5).Value = Format(Now, " & q & "hh:mm:ss" & q & ")" & vbCrLf
    c = c & "            Dim signOutTime As Date" & vbCrLf
    c = c & "            signOutTime = Now" & vbCrLf
    c = c & "            If Hour(signOutTime) < 18 Then" & vbCrLf
    c = c & "                If ws.Cells(i, 6).Value = " & q & "迟到" & q & " Then" & vbCrLf
    c = c & "                    ws.Cells(i, 7).Value = " & q & "迟到+早退" & q & vbCrLf
    c = c & "                Else" & vbCrLf
    c = c & "                    ws.Cells(i, 6).Value = " & q & "早退" & q & vbCrLf
    c = c & "                    ws.Cells(i, 7).Value = " & q & "早退签退" & q & vbCrLf
    c = c & "                End If" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    ws.Columns(" & q & "A:G" & q & ").AutoFit" & vbCrLf
    c = c & "    MsgBox " & q & "签退成功！时间: " & q & " & Format(Now, " & q & "hh:mm:ss" & q & "), vbInformation" & vbCrLf
    c = c & "    LoadTodayStatus" & vbCrLf
    c = c & "    QueryAttend" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdQuery_Click()" & vbCrLf
    c = c & "    QueryAttend" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub QueryAttend()" & vbCrLf
    c = c & "    lstAttend.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "考勤管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim qMonth As String" & vbCrLf
    c = c & "    qMonth = Trim(txtMonth.Text)" & vbCrLf
    c = c & "    Dim cWork As Long, cLate As Long, cEarly As Long, cLeave As Long, cAbsent As Long" & vbCrLf
    c = c & "    cWork = 0: cLate = 0: cEarly = 0: cLeave = 0: cAbsent = 0" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 2).Value) = gCurrentUser Then" & vbCrLf
    c = c & "            Dim recDate As String" & vbCrLf
    c = c & "            If IsDate(ws.Cells(i, 3).Value) Then" & vbCrLf
    c = c & "                recDate = Format(ws.Cells(i, 3).Value, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                recDate = CStr(ws.Cells(i, 3).Value)" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            If Left(recDate, 7) = qMonth Then" & vbCrLf
    c = c & "                lstAttend.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "                lstAttend.List(lstAttend.ListCount - 1, 1) = recDate" & vbCrLf
    c = c & "                lstAttend.List(lstAttend.ListCount - 1, 2) = Format(ws.Cells(i, 4).Value, " & q & "hh:mm:ss" & q & ")" & vbCrLf
    c = c & "                If Trim(CStr(ws.Cells(i, 5).Value)) <> " & q & q & " Then" & vbCrLf
    c = c & "                    lstAttend.List(lstAttend.ListCount - 1, 3) = Format(ws.Cells(i, 5).Value, " & q & "hh:mm:ss" & q & ")" & vbCrLf
    c = c & "                Else" & vbCrLf
    c = c & "                    lstAttend.List(lstAttend.ListCount - 1, 3) = " & q & "--" & q & vbCrLf
    c = c & "                End If" & vbCrLf
    c = c & "                lstAttend.List(lstAttend.ListCount - 1, 4) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "                lstAttend.List(lstAttend.ListCount - 1, 5) = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "                Dim aSt As String" & vbCrLf
    c = c & "                aSt = CStr(ws.Cells(i, 6).Value)" & vbCrLf
    c = c & "                If aSt = " & q & "出勤" & q & " Then cWork = cWork + 1" & vbCrLf
    c = c & "                If aSt = " & q & "迟到" & q & " Then cLate = cLate + 1" & vbCrLf
    c = c & "                If aSt = " & q & "早退" & q & " Then cEarly = cEarly + 1" & vbCrLf
    c = c & "                If aSt = " & q & "请假" & q & " Then cLeave = cLeave + 1" & vbCrLf
    c = c & "                If aSt = " & q & "缺勤" & q & " Then cAbsent = cAbsent + 1" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    lblSummary.Caption = " & q & "出勤:" & q & " & cWork & " & q & " 迟到:" & q & " & cLate & " & q & " 早退:" & q & " & cEarly & " & q & " 请假:" & q & " & cLeave & " & q & " 缺勤:" & q & " & cAbsent" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEditRec_Click()" & vbCrLf
    c = c & "    If lstAttend.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条考勤记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "只有管理员或主管可以编辑考勤记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim recID As String" & vbCrLf
    c = c & "    recID = lstAttend.List(lstAttend.ListIndex, 0)" & vbCrLf
    c = c & "    Dim newSt As String" & vbCrLf
    c = c & "    newSt = InputBox(" & q & "请输入新的考勤状态(出勤/迟到/早退/请假/缺勤):" & q & ", " & q & "编辑考勤" & q & ")" & vbCrLf
    c = c & "    If newSt <> " & q & q & " Then" & vbCrLf
    c = c & "        UpdateAttendStatus recID, newSt" & vbCrLf
    c = c & "        QueryAttend" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDeleteRec_Click()" & vbCrLf
    c = c & "    If lstAttend.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条考勤记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        MsgBox " & q & "只有管理员或主管可以删除考勤记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该考勤记录？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        DeleteAttendByID lstAttend.List(lstAttend.ListIndex, 0)" & vbCrLf
    c = c & "        QueryAttend" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 考勤窗体显示 ----------

Public Sub ShowAttendForm()
    Dim fName As String
    fName = GetConfigProp("AttendFormName")
    If fName = "" Or Not FormExists("AttendFormName") Then
        MsgBox "考勤窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("AttendFormName")
    frm.Show
End Sub

' ---------- 考勤数据操作函数 ----------

Public Function GenerateAttendID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("考勤管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "KQ" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateAttendID = "KQ" & Format(maxID + 1, "0000")
End Function

Public Sub UpdateAttendStatus(aID As String, newStatus As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("考勤管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = aID Then
            ws.Cells(i, 6).Value = newStatus
            Exit For
        End If
    Next i
End Sub

Public Sub DeleteAttendByID(aID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("考勤管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = aID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

' ---------- 创建报修管理表 ----------

Private Sub CreateRepairSheet()
    Dim ws As Worksheet
    If SheetExists("报修管理") Then
        Set ws = ThisWorkbook.Sheets("报修管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "报修管理"
    End If
    With ws
        .Range("A1").Value = "工单编号"
        .Range("B1").Value = "楼栋"
        .Range("C1").Value = "房号"
        .Range("D1").Value = "报修类型"
        .Range("E1").Value = "问题描述"
        .Range("F1").Value = "紧急程度"
        .Range("G1").Value = "报修人"
        .Range("H1").Value = "维修人员"
        .Range("I1").Value = "工单状态"
        .Range("J1").Value = "图片路径"
        .Range("K1").Value = "创建时间"
        .Range("L1").Value = "完成时间"
        .Range("A1:L1").Font.Bold = True
        .Columns("A:L").AutoFit
    End With
End Sub

' ---------- 创建报修管理窗体 frmRepairs ----------

Private Sub CreateRepairsForm()
    Dim oldForm As String
    oldForm = GetConfigProp("RepairsFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "RepairsFormName", actualName

    vbc.Properties("Caption") = "报修管理"
    vbc.Properties("Width") = 510
    vbc.Properties("Height") = 400
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 报修工单列表
    Set ctl = dsg.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstRepairs"
    ctl.Left = 12: ctl.Top = 12: ctl.Width = 476: ctl.Height = 320
    ctl.ColumnCount = 5
    ctl.ColumnWidths = "60;90;70;70;140"
    ctl.ColumnHeads = False

    ' 按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNewRepair": ctl.Caption = "新建报修"
    ctl.Left = 12: ctl.Top = 340: ctl.Width = 100: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdProcess": ctl.Caption = "查看/处理"
    ctl.Left = 120: ctl.Top = 340: ctl.Width = 100: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDelete": ctl.Caption = "删除"
    ctl.Left = 228: ctl.Top = 340: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdRefresh": ctl.Caption = "刷新"
    ctl.Left = 306: ctl.Top = 340: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdBack": ctl.Caption = "返回"
    ctl.Left = 400: ctl.Top = 340: ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    lstRepairs.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "报修管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim loc As String" & vbCrLf
    c = c & "        loc = ws.Cells(i, 2).Value & " & q & "-" & q & " & ws.Cells(i, 3).Value" & vbCrLf
    c = c & "        lstRepairs.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "        lstRepairs.List(lstRepairs.ListCount - 1, 1) = loc" & vbCrLf
    c = c & "        lstRepairs.List(lstRepairs.ListCount - 1, 2) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "        lstRepairs.List(lstRepairs.ListCount - 1, 3) = ws.Cells(i, 9).Value" & vbCrLf
    c = c & "        lstRepairs.List(lstRepairs.ListCount - 1, 4) = ws.Cells(i, 11).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNewRepair_Click()" & vbCrLf
    c = c & "    gEditRepairID = " & q & q & vbCrLf
    c = c & "    ShowRepairEditForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdProcess_Click()" & vbCrLf
    c = c & "    If lstRepairs.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条报修工单！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewRepairID = lstRepairs.List(lstRepairs.ListIndex, 0)" & vbCrLf
    c = c & "    ShowRepairViewForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelete_Click()" & vbCrLf
    c = c & "    If lstRepairs.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条报修工单！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        Dim creator As String" & vbCrLf
    c = c & "        Dim selID As String" & vbCrLf
    c = c & "        selID = lstRepairs.List(lstRepairs.ListIndex, 0)" & vbCrLf
    c = c & "        Dim wsC As Worksheet" & vbCrLf
    c = c & "        Set wsC = ThisWorkbook.Sheets(" & q & "报修管理" & q & ")" & vbCrLf
    c = c & "        Dim rr As Long" & vbCrLf
    c = c & "        For rr = 2 To wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "            If CStr(wsC.Cells(rr, 1).Value) = selID Then" & vbCrLf
    c = c & "                creator = wsC.Cells(rr, 7).Value" & vbCrLf
    c = c & "                Exit For" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        Next rr" & vbCrLf
    c = c & "        If creator <> gCurrentUser Then" & vbCrLf
    c = c & "            MsgBox " & q & "只能删除自己创建的报修工单！" & q & ", vbExclamation" & vbCrLf
    c = c & "            Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该报修工单？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        DeleteRepairByID lstRepairs.List(lstRepairs.ListIndex, 0)" & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdRefresh_Click()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdBack_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建报修编辑窗体 frmRepairEdit ----------

Private Sub CreateRepairEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("RepairEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "RepairEditFormName", actualName

    vbc.Properties("Caption") = "报修编辑"
    vbc.Properties("Width") = 450
    vbc.Properties("Height") = 300
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 楼栋
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label1": ctl.Caption = "楼栋："
    ctl.Left = 24: ctl.Top = 24: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboBuilding"
    ctl.Left = 84: ctl.Top = 24: ctl.Width = 120: ctl.Height = 24
    ctl.Style = 2

    ' 房号
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label2": ctl.Caption = "房号："
    ctl.Left = 216: ctl.Top = 24: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboRoom"
    ctl.Left = 276: ctl.Top = 24: ctl.Width = 120: ctl.Height = 24
    ctl.Style = 0

    ' 报修类型
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label3": ctl.Caption = "报修类型："
    ctl.Left = 24: ctl.Top = 54: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboRepairType"
    ctl.Left = 84: ctl.Top = 54: ctl.Width = 150: ctl.Height = 24
    ctl.Style = 0

    ' 问题描述
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label4": ctl.Caption = "问题描述："
    ctl.Left = 24: ctl.Top = 84: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtDesc"
    ctl.Left = 84: ctl.Top = 84: ctl.Width = 300: ctl.Height = 80
    ctl.MultiLine = True: ctl.ScrollBars = 2

    ' 图片路径
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label5": ctl.Caption = "图片路径："
    ctl.Left = 24: ctl.Top = 174: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtPicPath"
    ctl.Left = 84: ctl.Top = 174: ctl.Width = 250: ctl.Height = 24
    ctl.Enabled = False
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdBrowsePic": ctl.Caption = "选择图片"
    ctl.Left = 344: ctl.Top = 174: ctl.Width = 80: ctl.Height = 24

    ' 紧急报修
    Set ctl = dsg.Controls.Add("Forms.CheckBox.1")
    ctl.Name = "chkUrgent": ctl.Caption = "紧急报修"
    ctl.Left = 24: ctl.Top = 204: ctl.Width = 100: ctl.Height = 18

    ' 保存/取消
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 100: ctl.Top = 240: ctl.Width = 80: ctl.Height = 28
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 200: ctl.Top = 240: ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Dim bk As Long" & vbCrLf
    c = c & "    For bk = 1 To 10" & vbCrLf
    c = c & "        cboBuilding.AddItem bk & " & q & "栋" & q & vbCrLf
    c = c & "    Next bk" & vbCrLf
    c = c & "    cboRepairType.AddItem " & q & "水管" & q & vbCrLf
    c = c & "    cboRepairType.AddItem " & q & "电路" & q & vbCrLf
    c = c & "    cboRepairType.AddItem " & q & "电梯" & q & vbCrLf
    c = c & "    cboRepairType.AddItem " & q & "门窗" & q & vbCrLf
    c = c & "    cboRepairType.AddItem " & q & "其他" & q & vbCrLf
    c = c & "    If gEditRepairID <> " & q & q & " Then" & vbCrLf
    c = c & "        Me.Caption = " & q & "编辑报修工单" & q & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Me.Caption = " & q & "新建报修工单" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cboBuilding_Change()" & vbCrLf
    c = c & "    cboRoom.Clear" & vbCrLf
    c = c & "    If cboBuilding.Text = " & q & q & " Then Exit Sub" & vbCrLf
    c = c & "    Dim fl As Long, rm As Long" & vbCrLf
    c = c & "    For fl = 1 To 12" & vbCrLf
    c = c & "        For rm = 1 To 4" & vbCrLf
    c = c & "            cboRoom.AddItem Format(fl, " & q & "00" & q & ") & " & q & "0" & q & " & rm" & vbCrLf
    c = c & "        Next rm" & vbCrLf
    c = c & "    Next fl" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "报修管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gEditRepairID Then" & vbCrLf
    c = c & "            cboBuilding.Text = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            cboRoom.Text = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            cboRepairType.Text = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            txtDesc.Text = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            chkUrgent.Value = (ws.Cells(i, 6).Value = " & q & "紧急" & q & ")" & vbCrLf
    c = c & "            txtPicPath.Text = ws.Cells(i, 10).Value" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdBrowsePic_Click()" & vbCrLf
    c = c & "    Dim fd As Object" & vbCrLf
    c = c & "    Set fd = Application.FileDialog(1)" & vbCrLf
    c = c & "    fd.Title = " & q & "选择报修图片" & q & vbCrLf
    c = c & "    fd.Filters.Clear" & vbCrLf
    c = c & "    fd.Filters.Add " & q & "图片文件" & q & ", " & q & "*.jpg;*.jpeg;*.png;*.bmp;*.gif" & q & vbCrLf
    c = c & "    If fd.Show = -1 Then" & vbCrLf
    c = c & "        txtPicPath.Text = fd.SelectedItems(1)" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(cboBuilding.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择楼栋！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(cboRoom.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择或输入房号！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(cboRepairType.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择报修类型！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(txtDesc.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写问题描述！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim urg As String" & vbCrLf
    c = c & "    If chkUrgent.Value Then urg = " & q & "紧急" & q & " Else urg = " & q & "普通" & q & vbCrLf
    c = c & "    If gEditRepairID <> " & q & q & " Then" & vbCrLf
    c = c & "        UpdateRepairRecord gEditRepairID, cboBuilding.Text, cboRoom.Text, cboRepairType.Text, txtDesc.Text, urg, txtPicPath.Text" & vbCrLf
    c = c & "        MsgBox " & q & "报修工单已更新！" & q & ", vbInformation" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        AddRepairRecord cboBuilding.Text, cboRoom.Text, cboRepairType.Text, txtDesc.Text, urg, txtPicPath.Text" & vbCrLf
    c = c & "        MsgBox " & q & "报修工单已创建！" & q & ", vbInformation" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建报修查看/处理窗体 ----------

Private Sub CreateRepairViewForm()
    Dim oldForm As String
    oldForm = GetConfigProp("RepairViewFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "RepairViewFormName", actualName

    vbc.Properties("Caption") = "工单详情"
    vbc.Properties("Width") = 460
    vbc.Properties("Height") = 440
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 18

    ' 工单编号
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblIDCap": ctl.Caption = "工单编号："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblIDVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 120: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblUrgCap": ctl.Caption = "紧急程度："
    ctl.Left = 240: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblUrgVal": ctl.Caption = ""
    ctl.Left = 320: ctl.Top = yy: ctl.Width = 100: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblLocCap": ctl.Caption = "报修位置："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblLocVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 200: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTypeCap": ctl.Caption = "报修类型："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTypeVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 120: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDescCap": ctl.Caption = "问题描述："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtDescV"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 330: ctl.Height = 60
    ctl.MultiLine = True: ctl.Locked = True: ctl.ScrollBars = 2

    yy = yy + 68
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblReporterCap": ctl.Caption = "报修人："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblReporterVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblWorkerCap": ctl.Caption = "维修人员："
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblWorkerVal": ctl.Caption = ""
    ctl.Left = 300: ctl.Top = yy: ctl.Width = 130: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStCap": ctl.Caption = "工单状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20
    ctl.Font.Size = 11: ctl.Font.Bold = True

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblImgCap": ctl.Caption = "图片路径："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblImgVal": ctl.Caption = ""
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 230: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdUploadPic": ctl.Caption = "上传图片"
    ctl.Left = 340: ctl.Top = yy - 2: ctl.Width = 80: ctl.Height = 22

    ' 指派维修人员区域（仅部门主管/管理员可用）
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblAssignCap": ctl.Caption = "指派人员："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboAssign"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 150: ctl.Height = 20
    ctl.Style = 0
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdAssign": ctl.Caption = "确认指派"
    ctl.Left = 260: ctl.Top = yy - 2: ctl.Width = 80: ctl.Height = 22

    ' 操作按钮
    yy = yy + 40
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdProcessing"
    ctl.Caption = "处理中"
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 90: ctl.Height = 30
    ctl.BackColor = RGB(255, 165, 0)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdComplete"
    ctl.Caption = "工单完成"
    ctl.Left = 125: ctl.Top = yy: ctl.Width = 90: ctl.Height = 30
    ctl.BackColor = RGB(60, 179, 113)

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdEdit"
    ctl.Caption = "编辑"
    ctl.Left = 230: ctl.Top = yy: ctl.Width = 80: ctl.Height = 30

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose"
    ctl.Caption = "关闭"
    ctl.Left = 340: ctl.Top = yy: ctl.Width = 80: ctl.Height = 30

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    If gViewRepairID = " & q & q & " Then Unload Me: Exit Sub" & vbCrLf
    c = c & "    LoadUsers" & vbCrLf
    c = c & "    LoadDetail" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadUsers()" & vbCrLf
    c = c & "    cboAssign.Clear" & vbCrLf
    c = c & "    Dim wsU As Worksheet" & vbCrLf
    c = c & "    Set wsU = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = wsU.Cells(wsU.Rows.Count, 2).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        cboAssign.AddItem wsU.Cells(i, 2).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadDetail()" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "报修管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gViewRepairID Then" & vbCrLf
    c = c & "            lblIDVal.Caption = ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lblLocVal.Caption = ws.Cells(i, 2).Value & " & q & " " & q & " & ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            lblTypeVal.Caption = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            txtDescV.Text = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lblUrgVal.Caption = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            If ws.Cells(i, 6).Value = " & q & "紧急" & q & " Then" & vbCrLf
    c = c & "                lblUrgVal.ForeColor = RGB(220, 0, 0)" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lblUrgVal.ForeColor = RGB(0, 128, 0)" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            lblReporterVal.Caption = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            lblWorkerVal.Caption = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            lblStVal.Caption = ws.Cells(i, 9).Value" & vbCrLf
    c = c & "            lblImgVal.Caption = ws.Cells(i, 10).Value" & vbCrLf
    c = c & "            Dim st As String" & vbCrLf
    c = c & "            st = ws.Cells(i, 9).Value" & vbCrLf
    c = c & "            If st = " & q & "待处理" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(180, 0, 0)" & vbCrLf
    c = c & "                cmdProcessing.Enabled = False" & vbCrLf
    c = c & "                cmdComplete.Enabled = False" & vbCrLf
    c = c & "            ElseIf st = " & q & "已指派" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(30, 144, 255)" & vbCrLf
    c = c & "                cmdProcessing.Enabled = True" & vbCrLf
    c = c & "                cmdComplete.Enabled = False" & vbCrLf
    c = c & "            ElseIf st = " & q & "处理中" & q & " Then" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(255, 140, 0)" & vbCrLf
    c = c & "                cmdProcessing.Enabled = False" & vbCrLf
    c = c & "                cmdComplete.Enabled = True" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lblStVal.ForeColor = RGB(60, 179, 113)" & vbCrLf
    c = c & "                cmdProcessing.Enabled = False" & vbCrLf
    c = c & "                cmdComplete.Enabled = False" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Dim isHigh As Boolean" & vbCrLf
    c = c & "            isHigh = IsHighPrivilege(gCurrentRole)" & vbCrLf
    c = c & "            cboAssign.Enabled = isHigh" & vbCrLf
    c = c & "            cmdAssign.Enabled = isHigh And (st = " & q & "待处理" & q & " Or st = " & q & "已指派" & q & ")" & vbCrLf
    c = c & "            If ws.Cells(i, 8).Value <> " & q & q & " Then" & vbCrLf
    c = c & "                cboAssign.Text = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdAssign_Click()" & vbCrLf
    c = c & "    If Trim(cboAssign.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择维修人员！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    AssignRepairWorker gViewRepairID, cboAssign.Text" & vbCrLf
    c = c & "    MsgBox " & q & "已指派维修人员: " & q & " & cboAssign.Text, vbInformation" & vbCrLf
    c = c & "    LoadDetail" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdProcessing_Click()" & vbCrLf
    c = c & "    UpdateRepairStatus gViewRepairID, " & q & "处理中" & q & vbCrLf
    c = c & "    MsgBox " & q & "工单状态已更新为【处理中】！" & q & ", vbInformation" & vbCrLf
    c = c & "    LoadDetail" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdComplete_Click()" & vbCrLf
    c = c & "    UpdateRepairStatus gViewRepairID, " & q & "已完成" & q & vbCrLf
    c = c & "    MsgBox " & q & "工单已完成！" & q & ", vbInformation" & vbCrLf
    c = c & "    LoadDetail" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdUploadPic_Click()" & vbCrLf
    c = c & "    Dim fd As Object" & vbCrLf
    c = c & "    Set fd = Application.FileDialog(1)" & vbCrLf
    c = c & "    fd.Title = " & q & "选择报修图片" & q & vbCrLf
    c = c & "    fd.Filters.Clear" & vbCrLf
    c = c & "    fd.Filters.Add " & q & "图片文件" & q & ", " & q & "*.jpg;*.jpeg;*.png;*.bmp;*.gif" & q & vbCrLf
    c = c & "    If fd.Show = -1 Then" & vbCrLf
    c = c & "        Dim picPath As String" & vbCrLf
    c = c & "        picPath = fd.SelectedItems(1)" & vbCrLf
    c = c & "        UpdateRepairPic gViewRepairID, picPath" & vbCrLf
    c = c & "        lblImgVal.Caption = picPath" & vbCrLf
    c = c & "        MsgBox " & q & "图片已上传！" & q & ", vbInformation" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEdit_Click()" & vbCrLf
    c = c & "    gEditRepairID = gViewRepairID" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "    ShowRepairEditForm" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 报修管理窗体显示 ----------

Public Sub ShowRepairsForm()
    Dim fName As String
    fName = GetConfigProp("RepairsFormName")
    If fName = "" Or Not FormExists("RepairsFormName") Then
        MsgBox "报修管理窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("RepairsFormName")
    frm.Show
End Sub

Public Sub ShowRepairEditForm()
    Dim fName As String
    fName = GetConfigProp("RepairEditFormName")
    If fName = "" Or Not FormExists("RepairEditFormName") Then
        MsgBox "报修编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("RepairEditFormName")
    frm.Show
End Sub

Public Sub ShowRepairViewForm()
    Dim fName As String
    fName = GetConfigProp("RepairViewFormName")
    If fName = "" Or Not FormExists("RepairViewFormName") Then
        MsgBox "报修查看窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("RepairViewFormName")
    frm.Show
End Sub

' ---------- 报修管理数据操作函数 ----------

Private Function GenerateRepairID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("报修管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "BX" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateRepairID = "BX" & Format(maxID + 1, "0000")
End Function

Public Sub AddRepairRecord(sBuilding As String, sRoom As String, sType As String, sDesc As String, sUrg As String, sPic As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("报修管理")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateRepairID()
    ws.Cells(newRow, 2).Value = sBuilding
    ws.Cells(newRow, 3).Value = sRoom
    ws.Cells(newRow, 4).Value = sType
    ws.Cells(newRow, 5).Value = sDesc
    ws.Cells(newRow, 6).Value = sUrg
    ws.Cells(newRow, 7).Value = gCurrentUser
    ws.Cells(newRow, 8).Value = ""
    ws.Cells(newRow, 9).Value = "待处理"
    ws.Cells(newRow, 10).Value = sPic
    ws.Cells(newRow, 11).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(newRow, 12).Value = ""
    ws.Columns("A:L").AutoFit
End Sub

Public Sub UpdateRepairRecord(rID As String, sBuilding As String, sRoom As String, sType As String, sDesc As String, sUrg As String, sPic As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("报修管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = rID Then
            ws.Cells(i, 2).Value = sBuilding
            ws.Cells(i, 3).Value = sRoom
            ws.Cells(i, 4).Value = sType
            ws.Cells(i, 5).Value = sDesc
            ws.Cells(i, 6).Value = sUrg
            ws.Cells(i, 10).Value = sPic
            Exit For
        End If
    Next i
    ws.Columns("A:L").AutoFit
End Sub

Public Sub UpdateRepairStatus(rID As String, newStatus As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("报修管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = rID Then
            ws.Cells(i, 9).Value = newStatus
            If newStatus = "已完成" Then
                ws.Cells(i, 12).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub AssignRepairWorker(rID As String, worker As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("报修管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = rID Then
            ws.Cells(i, 8).Value = worker
            If ws.Cells(i, 9).Value = "待处理" Then
                ws.Cells(i, 9).Value = "已指派"
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub UpdateRepairPic(rID As String, picPath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("报修管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = rID Then
            ws.Cells(i, 10).Value = picPath
            Exit For
        End If
    Next i
End Sub

Public Sub DeleteRepairByID(rID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("报修管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = rID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

' ---------- 创建投诉建议表 ----------

Private Sub CreateComplaintSheet()
    Dim ws As Worksheet
    If SheetExists("投诉建议") Then
        Set ws = ThisWorkbook.Sheets("投诉建议")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "投诉建议"
    End If
    With ws
        .Range("A1").Value = "记录编号"
        .Range("B1").Value = "类型"
        .Range("C1").Value = "子类型"
        .Range("D1").Value = "投诉内容"
        .Range("E1").Value = "提交人"
        .Range("F1").Value = "处理状态"
        .Range("G1").Value = "回复内容"
        .Range("H1").Value = "附件路径"
        .Range("I1").Value = "提交时间"
        .Range("J1").Value = "回复时间"
        .Range("A1:J1").Font.Bold = True
        .Columns("A:J").AutoFit
    End With
End Sub

' ---------- 创建投诉建议窗体 frmComplaints ----------

Private Sub CreateComplaintsForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ComplaintsFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ComplaintsFormName", actualName

    vbc.Properties("Caption") = "投诉建议"
    vbc.Properties("Width") = 510
    vbc.Properties("Height") = 400
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    Set ctl = dsg.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstComplaints"
    ctl.Left = 12: ctl.Top = 12: ctl.Width = 476: ctl.Height = 320
    ctl.ColumnCount = 5
    ctl.ColumnWidths = "60;60;60;80;180"
    ctl.ColumnHeads = False

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNewComplaint": ctl.Caption = "新建投诉"
    ctl.Left = 12: ctl.Top = 340: ctl.Width = 100: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdReply": ctl.Caption = "回复投诉"
    ctl.Left = 120: ctl.Top = 340: ctl.Width = 100: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDelete": ctl.Caption = "删除"
    ctl.Left = 228: ctl.Top = 340: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdRefresh": ctl.Caption = "刷新"
    ctl.Left = 306: ctl.Top = 340: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdBack": ctl.Caption = "返回"
    ctl.Left = 400: ctl.Top = 340: ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    lstComplaints.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "投诉建议" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        lstComplaints.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "        lstComplaints.List(lstComplaints.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "        lstComplaints.List(lstComplaints.ListCount - 1, 2) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "        lstComplaints.List(lstComplaints.ListCount - 1, 3) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "        lstComplaints.List(lstComplaints.ListCount - 1, 4) = ws.Cells(i, 9).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNewComplaint_Click()" & vbCrLf
    c = c & "    gEditComplaintID = " & q & q & vbCrLf
    c = c & "    ShowComplaintEditForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdReply_Click()" & vbCrLf
    c = c & "    If lstComplaints.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditComplaintID = lstComplaints.List(lstComplaints.ListIndex, 0)" & vbCrLf
    c = c & "    ShowComplaintEditForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelete_Click()" & vbCrLf
    c = c & "    If lstComplaints.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "        Dim selID As String" & vbCrLf
    c = c & "        selID = lstComplaints.List(lstComplaints.ListIndex, 0)" & vbCrLf
    c = c & "        Dim wsC As Worksheet" & vbCrLf
    c = c & "        Set wsC = ThisWorkbook.Sheets(" & q & "投诉建议" & q & ")" & vbCrLf
    c = c & "        Dim rr As Long" & vbCrLf
    c = c & "        For rr = 2 To wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "            If CStr(wsC.Cells(rr, 1).Value) = selID Then" & vbCrLf
    c = c & "                If wsC.Cells(rr, 5).Value <> gCurrentUser Then" & vbCrLf
    c = c & "                    MsgBox " & q & "只能删除自己提交的记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "                    Exit Sub" & vbCrLf
    c = c & "                End If" & vbCrLf
    c = c & "                Exit For" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        Next rr" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该记录？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        DeleteComplaintByID lstComplaints.List(lstComplaints.ListIndex, 0)" & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdRefresh_Click()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdBack_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建投诉编辑窗体 frmComplaintEdit ----------

Private Sub CreateComplaintEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ComplaintEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ComplaintEditFormName", actualName

    vbc.Properties("Caption") = "投诉编辑"
    vbc.Properties("Width") = 430
    vbc.Properties("Height") = 380
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' 投诉类型
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label1": ctl.Caption = "投诉类型："
    ctl.Left = 24: ctl.Top = 24: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboType"
    ctl.Left = 84: ctl.Top = 24: ctl.Width = 150: ctl.Height = 24
    ctl.Style = 2

    ' 投诉子类型（动态显示）
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblSubType": ctl.Caption = "投诉子类："
    ctl.Left = 246: ctl.Top = 24: ctl.Width = 60: ctl.Height = 18
    ctl.Visible = False
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboSubType"
    ctl.Left = 306: ctl.Top = 24: ctl.Width = 100: ctl.Height = 24
    ctl.Style = 0
    ctl.Visible = False

    ' 当前状态
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label2": ctl.Caption = "当前状态："
    ctl.Left = 24: ctl.Top = 54: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboStatus"
    ctl.Left = 84: ctl.Top = 54: ctl.Width = 120: ctl.Height = 24
    ctl.Style = 2

    ' 投诉内容
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label3": ctl.Caption = "投诉内容："
    ctl.Left = 24: ctl.Top = 84: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtContent"
    ctl.Left = 84: ctl.Top = 84: ctl.Width = 300: ctl.Height = 80
    ctl.MultiLine = True: ctl.ScrollBars = 2

    ' 回复内容
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "Label4": ctl.Caption = "回复内容："
    ctl.Left = 24: ctl.Top = 174: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtReply"
    ctl.Left = 84: ctl.Top = 174: ctl.Width = 300: ctl.Height = 60
    ctl.MultiLine = True: ctl.ScrollBars = 2

    ' 附件信息 Frame
    Set ctl = dsg.Controls.Add("Forms.Frame.1")
    ctl.Name = "fraAttachment": ctl.Caption = "附件信息（可选）"
    ctl.Left = 24: ctl.Top = 240: ctl.Width = 360: ctl.Height = 50

    Dim frm As Object
    Set frm = ctl
    Set ctl = frm.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtAttachPath"
    ctl.Left = 10: ctl.Top = 18: ctl.Width = 240: ctl.Height = 22
    ctl.Enabled = False
    Set ctl = frm.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdAttachFile": ctl.Caption = "选择文件"
    ctl.Left = 260: ctl.Top = 18: ctl.Width = 80: ctl.Height = 22

    ' 保存/取消
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 100: ctl.Top = 300: ctl.Width = 80: ctl.Height = 28
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 200: ctl.Top = 300: ctl.Width = 80: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim cc As String
    cc = "Option Explicit" & vbCrLf
    cc = cc & "" & vbCrLf
    cc = cc & "Private Sub UserForm_Initialize()" & vbCrLf
    cc = cc & "    cboType.AddItem " & q & "投诉" & q & vbCrLf
    cc = cc & "    cboType.AddItem " & q & "建议" & q & vbCrLf
    cc = cc & "    cboType.AddItem " & q & "咨询" & q & vbCrLf
    cc = cc & "    cboStatus.AddItem " & q & "未处理" & q & vbCrLf
    cc = cc & "    cboStatus.AddItem " & q & "处理中" & q & vbCrLf
    cc = cc & "    cboStatus.AddItem " & q & "已解决" & q & vbCrLf
    cc = cc & "    cboStatus.AddItem " & q & "已关闭" & q & vbCrLf
    cc = cc & "    cboStatus.ListIndex = 0" & vbCrLf
    cc = cc & "    If gEditComplaintID <> " & q & q & " Then" & vbCrLf
    cc = cc & "        Me.Caption = " & q & "查看/回复投诉" & q & vbCrLf
    cc = cc & "        LoadData" & vbCrLf
    cc = cc & "    Else" & vbCrLf
    cc = cc & "        Me.Caption = " & q & "新建投诉/建议" & q & vbCrLf
    cc = cc & "    End If" & vbCrLf
    cc = cc & "End Sub" & vbCrLf
    cc = cc & "" & vbCrLf
    cc = cc & "Private Sub cboType_Change()" & vbCrLf
    cc = cc & "    If cboType.Text = " & q & "投诉" & q & " Then" & vbCrLf
    cc = cc & "        lblSubType.Visible = True" & vbCrLf
    cc = cc & "        cboSubType.Visible = True" & vbCrLf
    cc = cc & "        cboSubType.Clear" & vbCrLf
    cc = cc & "        cboSubType.AddItem " & q & "设施问题" & q & vbCrLf
    cc = cc & "        cboSubType.AddItem " & q & "扰民问题" & q & vbCrLf
    cc = cc & "        cboSubType.AddItem " & q & "服务质量" & q & vbCrLf
    cc = cc & "        cboSubType.AddItem " & q & "其他" & q & vbCrLf
    cc = cc & "    Else" & vbCrLf
    cc = cc & "        lblSubType.Visible = False" & vbCrLf
    cc = cc & "        cboSubType.Visible = False" & vbCrLf
    cc = cc & "        cboSubType.Text = " & q & q & vbCrLf
    cc = cc & "    End If" & vbCrLf
    cc = cc & "End Sub" & vbCrLf
    cc = cc & "" & vbCrLf
    cc = cc & "Private Sub LoadData()" & vbCrLf
    cc = cc & "    Dim ws As Worksheet" & vbCrLf
    cc = cc & "    Set ws = ThisWorkbook.Sheets(" & q & "投诉建议" & q & ")" & vbCrLf
    cc = cc & "    Dim lr As Long, i As Long" & vbCrLf
    cc = cc & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    cc = cc & "    For i = 2 To lr" & vbCrLf
    cc = cc & "        If CStr(ws.Cells(i, 1).Value) = gEditComplaintID Then" & vbCrLf
    cc = cc & "            cboType.Text = ws.Cells(i, 2).Value" & vbCrLf
    cc = cc & "            If ws.Cells(i, 3).Value <> " & q & q & " Then" & vbCrLf
    cc = cc & "                cboSubType.Text = ws.Cells(i, 3).Value" & vbCrLf
    cc = cc & "            End If" & vbCrLf
    cc = cc & "            txtContent.Text = ws.Cells(i, 4).Value" & vbCrLf
    cc = cc & "            cboStatus.Text = ws.Cells(i, 6).Value" & vbCrLf
    cc = cc & "            txtReply.Text = ws.Cells(i, 7).Value" & vbCrLf
    cc = cc & "            txtAttachPath.Text = ws.Cells(i, 8).Value" & vbCrLf
    cc = cc & "            Dim isCreator As Boolean" & vbCrLf
    cc = cc & "            isCreator = (ws.Cells(i, 5).Value = gCurrentUser)" & vbCrLf
    cc = cc & "            Dim isHigh As Boolean" & vbCrLf
    cc = cc & "            isHigh = IsHighPrivilege(gCurrentRole)" & vbCrLf
    cc = cc & "            cboType.Enabled = (isCreator Or isHigh)" & vbCrLf
    cc = cc & "            txtContent.Locked = Not (isCreator Or isHigh)" & vbCrLf
    cc = cc & "            txtReply.Locked = Not isHigh" & vbCrLf
    cc = cc & "            cboStatus.Enabled = isHigh" & vbCrLf
    cc = cc & "            Exit For" & vbCrLf
    cc = cc & "        End If" & vbCrLf
    cc = cc & "    Next i" & vbCrLf
    cc = cc & "End Sub" & vbCrLf
    cc = cc & "" & vbCrLf
    cc = cc & "Private Sub cmdAttachFile_Click()" & vbCrLf
    cc = cc & "    Dim fd As Object" & vbCrLf
    cc = cc & "    Set fd = Application.FileDialog(1)" & vbCrLf
    cc = cc & "    fd.Title = " & q & "选择附件" & q & vbCrLf
    cc = cc & "    fd.Filters.Clear" & vbCrLf
    cc = cc & "    fd.Filters.Add " & q & "所有文件" & q & ", " & q & "*.*" & q & vbCrLf
    cc = cc & "    If fd.Show = -1 Then" & vbCrLf
    cc = cc & "        txtAttachPath.Text = fd.SelectedItems(1)" & vbCrLf
    cc = cc & "    End If" & vbCrLf
    cc = cc & "End Sub" & vbCrLf
    cc = cc & "" & vbCrLf
    cc = cc & "Private Sub cmdSave_Click()" & vbCrLf
    cc = cc & "    If Trim(cboType.Text) = " & q & q & " Then" & vbCrLf
    cc = cc & "        MsgBox " & q & "请选择投诉类型！" & q & ", vbExclamation" & vbCrLf
    cc = cc & "        Exit Sub" & vbCrLf
    cc = cc & "    End If" & vbCrLf
    cc = cc & "    If Trim(txtContent.Text) = " & q & q & " Then" & vbCrLf
    cc = cc & "        MsgBox " & q & "请填写投诉内容！" & q & ", vbExclamation" & vbCrLf
    cc = cc & "        Exit Sub" & vbCrLf
    cc = cc & "    End If" & vbCrLf
    cc = cc & "    If gEditComplaintID <> " & q & q & " Then" & vbCrLf
    cc = cc & "        Dim replyChanged As Boolean" & vbCrLf
    cc = cc & "        replyChanged = (Trim(txtReply.Text) <> " & q & q & ")" & vbCrLf
    cc = cc & "        Dim finalStatus As String" & vbCrLf
    cc = cc & "        finalStatus = cboStatus.Text" & vbCrLf
    cc = cc & "        If replyChanged And finalStatus <> " & q & "已解决" & q & " And finalStatus <> " & q & "已关闭" & q & " Then" & vbCrLf
    cc = cc & "            finalStatus = " & q & "已解决" & q & vbCrLf
    cc = cc & "        End If" & vbCrLf
    cc = cc & "        UpdateComplaintRecord gEditComplaintID, cboType.Text, cboSubType.Text, txtContent.Text, finalStatus, txtReply.Text, txtAttachPath.Text" & vbCrLf
    cc = cc & "        MsgBox " & q & "记录已更新！" & q & ", vbInformation" & vbCrLf
    cc = cc & "    Else" & vbCrLf
    cc = cc & "        AddComplaintRecord cboType.Text, cboSubType.Text, txtContent.Text, txtAttachPath.Text" & vbCrLf
    cc = cc & "        MsgBox " & q & "投诉/建议已提交！" & q & ", vbInformation" & vbCrLf
    cc = cc & "    End If" & vbCrLf
    cc = cc & "    Unload Me" & vbCrLf
    cc = cc & "End Sub" & vbCrLf
    cc = cc & "" & vbCrLf
    cc = cc & "Private Sub cmdCancel_Click()" & vbCrLf
    cc = cc & "    Unload Me" & vbCrLf
    cc = cc & "End Sub"

    cm.InsertLines 1, cc
End Sub

' ---------- 投诉建议窗体显示 ----------

Public Sub ShowComplaintsForm()
    Dim fName As String
    fName = GetConfigProp("ComplaintsFormName")
    If fName = "" Or Not FormExists("ComplaintsFormName") Then
        MsgBox "投诉建议窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ComplaintsFormName")
    frm.Show
End Sub

Public Sub ShowComplaintEditForm()
    Dim fName As String
    fName = GetConfigProp("ComplaintEditFormName")
    If fName = "" Or Not FormExists("ComplaintEditFormName") Then
        MsgBox "投诉编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ComplaintEditFormName")
    frm.Show
End Sub

' ---------- 投诉建议数据操作函数 ----------

Private Function GenerateComplaintID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("投诉建议")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "TS" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateComplaintID = "TS" & Format(maxID + 1, "0000")
End Function

Public Sub AddComplaintRecord(sType As String, sSub As String, sContent As String, sAttach As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("投诉建议")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateComplaintID()
    ws.Cells(newRow, 2).Value = sType
    ws.Cells(newRow, 3).Value = sSub
    ws.Cells(newRow, 4).Value = sContent
    ws.Cells(newRow, 5).Value = gCurrentUser
    ws.Cells(newRow, 6).Value = "未处理"
    ws.Cells(newRow, 7).Value = ""
    ws.Cells(newRow, 8).Value = sAttach
    ws.Cells(newRow, 9).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(newRow, 10).Value = ""
    ws.Columns("A:J").AutoFit
End Sub

Public Sub UpdateComplaintRecord(cID As String, sType As String, sSub As String, sContent As String, sStatus As String, sReply As String, sAttach As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("投诉建议")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = cID Then
            ws.Cells(i, 2).Value = sType
            ws.Cells(i, 3).Value = sSub
            ws.Cells(i, 4).Value = sContent
            ws.Cells(i, 6).Value = sStatus
            ws.Cells(i, 7).Value = sReply
            ws.Cells(i, 8).Value = sAttach
            If Trim(sReply) <> "" Then
                ws.Cells(i, 10).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
            End If
            Exit For
        End If
    Next i
    ws.Columns("A:J").AutoFit
End Sub

Public Sub DeleteComplaintByID(cID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("投诉建议")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = cID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

' ---------- 创建费用管理表 ----------

Private Sub CreateFeeSheet()
    Dim ws As Worksheet
    If SheetExists("费用管理") Then
        Set ws = ThisWorkbook.Sheets("费用管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "费用管理"
    End If
    With ws
        .Range("A1").Value = "楼号"
        .Range("B1").Value = "欠缴物业费"
        .Range("C1").Value = "应收物业费"
        .Range("D1").Value = "欠缴公摊费"
        .Range("E1").Value = "本月公摊费"
        .Range("F1").Value = "应收公摊费"
        .Range("G1").Value = "物业费欠缴率"
        .Range("H1").Value = "公摊费欠缴率"
        .Range("I1").Value = "缴费周期"
        .Range("A1:I1").Font.Bold = True
        .Columns("A:I").AutoFit
    End With
End Sub

' ---------- 创建费用管理窗体 ----------

Private Sub CreateFeeManagerForm()
    Dim oldForm As String
    oldForm = GetConfigProp("FeeManagerFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "FeeManagerFormName", actualName

    vbc.Properties("Caption") = "费用管理"
    vbc.Properties("Width") = 780
    vbc.Properties("Height") = 600
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' ---- 顶部统计卡片区域 ----
    ' 卡片1: 总应收物业费
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard1Title": ctl.Caption = "总应收物业费"
    ctl.Left = 15: ctl.Top = 8: ctl.Width = 170: ctl.Height = 16
    ctl.Font.Bold = True: ctl.Font.Size = 9: ctl.TextAlign = 2
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard1Val": ctl.Caption = "0.00"
    ctl.Left = 15: ctl.Top = 26: ctl.Width = 170: ctl.Height = 20
    ctl.Font.Size = 12: ctl.Font.Bold = True: ctl.TextAlign = 2
    ctl.ForeColor = RGB(30, 144, 255)

    ' 卡片2: 总欠缴物业费
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard2Title": ctl.Caption = "总欠缴物业费"
    ctl.Left = 200: ctl.Top = 8: ctl.Width = 170: ctl.Height = 16
    ctl.Font.Bold = True: ctl.Font.Size = 9: ctl.TextAlign = 2
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard2Val": ctl.Caption = "0.00"
    ctl.Left = 200: ctl.Top = 26: ctl.Width = 170: ctl.Height = 20
    ctl.Font.Size = 12: ctl.Font.Bold = True: ctl.TextAlign = 2
    ctl.ForeColor = RGB(220, 80, 80)

    ' 卡片3: 总应收公摊费
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard3Title": ctl.Caption = "总应收公摊费"
    ctl.Left = 385: ctl.Top = 8: ctl.Width = 170: ctl.Height = 16
    ctl.Font.Bold = True: ctl.Font.Size = 9: ctl.TextAlign = 2
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard3Val": ctl.Caption = "0.00"
    ctl.Left = 385: ctl.Top = 26: ctl.Width = 170: ctl.Height = 20
    ctl.Font.Size = 12: ctl.Font.Bold = True: ctl.TextAlign = 2
    ctl.ForeColor = RGB(30, 144, 255)

    ' 卡片4: 总欠缴公摊费
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard4Title": ctl.Caption = "总欠缴公摊费"
    ctl.Left = 570: ctl.Top = 8: ctl.Width = 170: ctl.Height = 16
    ctl.Font.Bold = True: ctl.Font.Size = 9: ctl.TextAlign = 2
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard4Val": ctl.Caption = "0.00"
    ctl.Left = 570: ctl.Top = 26: ctl.Width = 170: ctl.Height = 20
    ctl.Font.Size = 12: ctl.Font.Bold = True: ctl.TextAlign = 2
    ctl.ForeColor = RGB(220, 80, 80)

    ' 比率行
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRateInfo": ctl.Caption = "物业费欠缴率: 0%  |  公摊费欠缴率: 0%"
    ctl.Left = 15: ctl.Top = 50: ctl.Width = 740: ctl.Height = 16
    ctl.Font.Size = 9: ctl.Font.Bold = True: ctl.TextAlign = 2

    ' ---- 缴费周期筛选 ----
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPeriod": ctl.Caption = "缴费周期:"
    ctl.Left = 15: ctl.Top = 72: ctl.Width = 60: ctl.Height = 16
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboPeriod"
    ctl.Left = 78: ctl.Top = 70: ctl.Width = 90: ctl.Height = 20
    ctl.Style = 0
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdFilter": ctl.Caption = "查询"
    ctl.Left = 175: ctl.Top = 70: ctl.Width = 50: ctl.Height = 20
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdChart": ctl.Caption = "柱状图分析"
    ctl.Left = 232: ctl.Top = 70: ctl.Width = 80: ctl.Height = 20

    ' ---- 费用列表 ----
    Set ctl = dsg.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstFee"
    ctl.Left = 15: ctl.Top = 96: ctl.Width = 740: ctl.Height = 400
    ctl.ColumnCount = 9
    ctl.ColumnWidths = "45;80;80;80;80;80;70;70;70"
    ctl.ColumnHeads = False

    ' ---- 底部按钮 ----
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNew": ctl.Caption = "新建"
    ctl.Left = 15: ctl.Top = 505: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdEdit": ctl.Caption = "编辑"
    ctl.Left = 92: ctl.Top = 505: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDelete": ctl.Caption = "删除"
    ctl.Left = 169: ctl.Top = 505: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdImport": ctl.Caption = "导入"
    ctl.Left = 260: ctl.Top = 505: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdExport": ctl.Caption = "导出"
    ctl.Left = 337: ctl.Top = 505: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdRefresh": ctl.Caption = "刷新"
    ctl.Left = 414: ctl.Top = 505: ctl.Width = 70: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 685: ctl.Top = 505: ctl.Width = 70: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    LoadPeriods" & vbCrLf
    c = c & "    cboPeriod.Text = Format(Date, " & q & "yyyy-mm" & q & ")" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadPeriods()" & vbCrLf
    c = c & "    cboPeriod.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "费用管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim dict As Object" & vbCrLf
    c = c & "    Set dict = CreateObject(" & q & "Scripting.Dictionary" & q & ")" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim p As String" & vbCrLf
    c = c & "        p = CStr(ws.Cells(i, 9).Value)" & vbCrLf
    c = c & "        If p <> " & q & q & " And Not dict.Exists(p) Then" & vbCrLf
    c = c & "            dict.Add p, 1" & vbCrLf
    c = c & "            cboPeriod.AddItem p" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    lstFee.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "费用管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim filterP As String" & vbCrLf
    c = c & "    filterP = Trim(cboPeriod.Text)" & vbCrLf
    c = c & "    Dim sumRecv As Double, sumOweProp As Double" & vbCrLf
    c = c & "    Dim sumRecvPub As Double, sumOwePub As Double" & vbCrLf
    c = c & "    sumRecv = 0: sumOweProp = 0: sumRecvPub = 0: sumOwePub = 0" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim prd As String" & vbCrLf
    c = c & "        prd = CStr(ws.Cells(i, 9).Value)" & vbCrLf
    c = c & "        If filterP = " & q & q & " Or prd = filterP Then" & vbCrLf
    c = c & "            lstFee.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lstFee.List(lstFee.ListCount - 1, 1) = Format(Val(ws.Cells(i, 2).Value), " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "            lstFee.List(lstFee.ListCount - 1, 2) = Format(Val(ws.Cells(i, 3).Value), " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "            lstFee.List(lstFee.ListCount - 1, 3) = Format(Val(ws.Cells(i, 4).Value), " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "            lstFee.List(lstFee.ListCount - 1, 4) = Format(Val(ws.Cells(i, 5).Value), " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "            lstFee.List(lstFee.ListCount - 1, 5) = Format(Val(ws.Cells(i, 6).Value), " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "            lstFee.List(lstFee.ListCount - 1, 6) = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            lstFee.List(lstFee.ListCount - 1, 7) = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            lstFee.List(lstFee.ListCount - 1, 8) = prd" & vbCrLf
    c = c & "            sumRecv = sumRecv + Val(ws.Cells(i, 3).Value)" & vbCrLf
    c = c & "            sumOweProp = sumOweProp + Val(ws.Cells(i, 2).Value)" & vbCrLf
    c = c & "            sumRecvPub = sumRecvPub + Val(ws.Cells(i, 6).Value)" & vbCrLf
    c = c & "            sumOwePub = sumOwePub + Val(ws.Cells(i, 4).Value)" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    lblCard1Val.Caption = Format(sumRecv, " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "    lblCard2Val.Caption = Format(sumOweProp, " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "    lblCard3Val.Caption = Format(sumRecvPub, " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "    lblCard4Val.Caption = Format(sumOwePub, " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "    Dim rProp As String, rPub As String" & vbCrLf
    c = c & "    If sumRecv > 0 Then rProp = Format(sumOweProp / sumRecv * 100, " & q & "0.0" & q & ") & " & q & "%" & q & " Else rProp = " & q & "0%" & q & vbCrLf
    c = c & "    If sumRecvPub > 0 Then rPub = Format(sumOwePub / sumRecvPub * 100, " & q & "0.0" & q & ") & " & q & "%" & q & " Else rPub = " & q & "0%" & q & vbCrLf
    c = c & "    lblRateInfo.Caption = " & q & "物业费欠缴率: " & q & " & rProp & " & q & "  |  公摊费欠缴率: " & q & " & rPub" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdFilter_Click()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNew_Click()" & vbCrLf
    c = c & "    gEditFeeRow = 0" & vbCrLf
    c = c & "    ShowFeeEditForm" & vbCrLf
    c = c & "    LoadPeriods" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEdit_Click()" & vbCrLf
    c = c & "    If lstFee.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim selBld As String, selPrd As String" & vbCrLf
    c = c & "    selBld = lstFee.List(lstFee.ListIndex, 0)" & vbCrLf
    c = c & "    selPrd = lstFee.List(lstFee.ListIndex, 8)" & vbCrLf
    c = c & "    gEditFeeRow = FindFeeRow(selBld, selPrd)" & vbCrLf
    c = c & "    If gEditFeeRow > 0 Then" & vbCrLf
    c = c & "        ShowFeeEditForm" & vbCrLf
    c = c & "        LoadPeriods" & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelete_Click()" & vbCrLf
    c = c & "    If lstFee.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该费用记录？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        Dim selBld As String, selPrd As String" & vbCrLf
    c = c & "        selBld = lstFee.List(lstFee.ListIndex, 0)" & vbCrLf
    c = c & "        selPrd = lstFee.List(lstFee.ListIndex, 8)" & vbCrLf
    c = c & "        Dim rw As Long" & vbCrLf
    c = c & "        rw = FindFeeRow(selBld, selPrd)" & vbCrLf
    c = c & "        If rw > 0 Then" & vbCrLf
    c = c & "            ThisWorkbook.Sheets(" & q & "费用管理" & q & ").Rows(rw).Delete" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        LoadPeriods" & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdImport_Click()" & vbCrLf
    c = c & "    Dim fd As Object" & vbCrLf
    c = c & "    Set fd = Application.FileDialog(1)" & vbCrLf
    c = c & "    fd.Title = " & q & "选择导入文件" & q & vbCrLf
    c = c & "    fd.Filters.Clear" & vbCrLf
    c = c & "    fd.Filters.Add " & q & "Excel文件" & q & ", " & q & "*.xlsm;*.xlsx;*.xls" & q & vbCrLf
    c = c & "    If fd.Show <> -1 Then Exit Sub" & vbCrLf
    c = c & "    Dim fPath As String" & vbCrLf
    c = c & "    fPath = fd.SelectedItems(1)" & vbCrLf
    c = c & "    ImportFeeData fPath" & vbCrLf
    c = c & "    LoadPeriods" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "    MsgBox " & q & "导入完成！" & q & ", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdExport_Click()" & vbCrLf
    c = c & "    ExportFeeData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdChart_Click()" & vbCrLf
    c = c & "    GenerateFeeChart" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdRefresh_Click()" & vbCrLf
    c = c & "    LoadPeriods" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建费用编辑窗体 ----------

Private Sub CreateFeeEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("FeeEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "FeeEditFormName", actualName

    vbc.Properties("Caption") = "费用编辑"
    vbc.Properties("Width") = 380
    vbc.Properties("Height") = 300
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 18

    ' 楼号
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl1": ctl.Caption = "楼号："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboBuilding"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20
    ctl.Style = 0

    ' 缴费周期
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl9": ctl.Caption = "缴费周期："
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtPeriod"
    ctl.Left = 285: ctl.Top = yy: ctl.Width = 70: ctl.Height = 20

    yy = yy + 30
    ' 应收物业费
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl3": ctl.Caption = "应收物业费："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtRecvProp"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20

    yy = yy + 30
    ' 欠缴物业费
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl2": ctl.Caption = "欠缴物业费："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtOweProp"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20

    yy = yy + 30
    ' 本月公摊费
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl5": ctl.Caption = "本月公摊费："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtMonthPub"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20

    yy = yy + 30
    ' 欠缴公摊费
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl4": ctl.Caption = "欠缴公摊费："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtOwePub"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20

    yy = yy + 30
    ' 自动计算提示
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCalcInfo"
    ctl.Caption = "* 应收公摊费、物业欠缴率、公摊欠缴率由系统自动计算"
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 340: ctl.Height = 16
    ctl.ForeColor = RGB(128, 128, 128): ctl.Font.Size = 8

    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 80: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 190: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Dim bk As Long" & vbCrLf
    c = c & "    For bk = 1 To 10" & vbCrLf
    c = c & "        cboBuilding.AddItem bk & " & q & "栋" & q & vbCrLf
    c = c & "    Next bk" & vbCrLf
    c = c & "    txtPeriod.Text = Format(Date, " & q & "yyyy-mm" & q & ")" & vbCrLf
    c = c & "    If gEditFeeRow > 0 Then" & vbCrLf
    c = c & "        Me.Caption = " & q & "编辑费用记录" & q & vbCrLf
    c = c & "        Dim ws As Worksheet" & vbCrLf
    c = c & "        Set ws = ThisWorkbook.Sheets(" & q & "费用管理" & q & ")" & vbCrLf
    c = c & "        cboBuilding.Text = ws.Cells(gEditFeeRow, 1).Value" & vbCrLf
    c = c & "        txtOweProp.Text = ws.Cells(gEditFeeRow, 2).Value" & vbCrLf
    c = c & "        txtRecvProp.Text = ws.Cells(gEditFeeRow, 3).Value" & vbCrLf
    c = c & "        txtOwePub.Text = ws.Cells(gEditFeeRow, 4).Value" & vbCrLf
    c = c & "        txtMonthPub.Text = ws.Cells(gEditFeeRow, 5).Value" & vbCrLf
    c = c & "        txtPeriod.Text = ws.Cells(gEditFeeRow, 9).Value" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Me.Caption = " & q & "新建费用记录" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(cboBuilding.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择楼号！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(txtPeriod.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写缴费周期！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    SaveFeeRecord gEditFeeRow, cboBuilding.Text, Val(txtOweProp.Text), Val(txtRecvProp.Text), Val(txtOwePub.Text), Val(txtMonthPub.Text), txtPeriod.Text" & vbCrLf
    c = c & "    MsgBox " & q & "费用记录已保存！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 费用管理窗体显示 ----------

Public Sub ShowFeeManagerForm()
    Dim fName As String
    fName = GetConfigProp("FeeManagerFormName")
    If fName = "" Or Not FormExists("FeeManagerFormName") Then
        MsgBox "费用管理窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("FeeManagerFormName")
    frm.Show
End Sub

Public Sub ShowFeeEditForm()
    Dim fName As String
    fName = GetConfigProp("FeeEditFormName")
    If fName = "" Or Not FormExists("FeeEditFormName") Then
        MsgBox "费用编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("FeeEditFormName")
    frm.Show
End Sub

' ---------- 费用管理数据操作函数 ----------

Public Function FindFeeRow(bld As String, prd As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("费用管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    FindFeeRow = 0
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = bld And CStr(ws.Cells(i, 9).Value) = prd Then
            FindFeeRow = i
            Exit Function
        End If
    Next i
End Function

Public Sub SaveFeeRecord(rowNum As Long, sBld As String, oweProp As Double, recvProp As Double, owePub As Double, monthPub As Double, sPeriod As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("费用管理")
    Dim r As Long
    If rowNum > 0 Then
        r = rowNum
    Else
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If
    ws.Cells(r, 1).Value = sBld
    ws.Cells(r, 2).Value = oweProp
    ws.Cells(r, 3).Value = recvProp
    ws.Cells(r, 4).Value = owePub
    ws.Cells(r, 5).Value = monthPub
    ' 自动计算应收公摊费 = 上期本月公摊费 + 本期本月公摊费
    Dim prevPub As Double
    prevPub = GetPrevMonthPub(sBld, sPeriod)
    Dim recvPub As Double
    recvPub = prevPub + monthPub
    ws.Cells(r, 6).Value = recvPub
    ' 自动计算欠缴率
    If recvProp > 0 Then
        ws.Cells(r, 7).Value = Format(oweProp / recvProp * 100, "0.0") & "%"
    Else
        ws.Cells(r, 7).Value = "0%"
    End If
    If recvPub > 0 Then
        ws.Cells(r, 8).Value = Format(owePub / recvPub * 100, "0.0") & "%"
    Else
        ws.Cells(r, 8).Value = "0%"
    End If
    ws.Cells(r, 9).Value = sPeriod
    ws.Columns("A:I").AutoFit
End Sub

Private Function GetPrevMonthPub(bld As String, curPeriod As String) As Double
    GetPrevMonthPub = 0
    On Error GoTo exitFunc
    Dim yr As Long, mo As Long
    yr = Val(Left(curPeriod, 4))
    mo = Val(Mid(curPeriod, 6, 2))
    mo = mo - 1
    If mo < 1 Then mo = 12: yr = yr - 1
    Dim prevP As String
    prevP = Format(yr, "0000") & "-" & Format(mo, "00")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("费用管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = bld And CStr(ws.Cells(i, 9).Value) = prevP Then
            GetPrevMonthPub = Val(ws.Cells(i, 5).Value)
            Exit Function
        End If
    Next i
exitFunc:
End Function

Public Sub ImportFeeData(fPath As String)
    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(fPath, ReadOnly:=True)
    Dim srcWs As Worksheet
    Set srcWs = srcWb.Sheets(1)
    Dim dstWs As Worksheet
    Set dstWs = ThisWorkbook.Sheets("费用管理")
    Dim srcLr As Long, i As Long
    srcLr = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    Dim startRow As Long
    startRow = 2
    If CStr(srcWs.Cells(1, 1).Value) = "楼号" Then startRow = 2 Else startRow = 1
    For i = startRow To srcLr
        Dim bld As String
        bld = CStr(srcWs.Cells(i, 1).Value)
        If bld = "" Then GoTo NextRow
        Dim prd As String
        prd = CStr(srcWs.Cells(i, 9).Value)
        Dim oweProp As Double, recvProp As Double
        Dim owePub As Double, monthPub As Double
        oweProp = Val(srcWs.Cells(i, 2).Value)
        recvProp = Val(srcWs.Cells(i, 3).Value)
        owePub = Val(srcWs.Cells(i, 4).Value)
        monthPub = Val(srcWs.Cells(i, 5).Value)
        Dim existRow As Long
        existRow = FindFeeRow(bld, prd)
        SaveFeeRecord existRow, bld, oweProp, recvProp, owePub, monthPub, prd
NextRow:
    Next i
    srcWb.Close False
End Sub

Public Sub ExportFeeData()
    Dim fd As Object
    Set fd = Application.FileDialog(2)
    fd.Title = "选择导出路径"
    fd.InitialFileName = "费用管理导出_" & Format(Now, "yyyymmdd") & ".xlsx"
    fd.FilterIndex = 1
    If fd.Show <> -1 Then Exit Sub
    Dim savePath As String
    savePath = fd.SelectedItems(1)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("费用管理")
    ws.Copy
    Dim newWb As Workbook
    Set newWb = ActiveWorkbook
    On Error Resume Next
    If Right(LCase(savePath), 5) = ".xlsm" Then
        newWb.SaveAs savePath, xlOpenXMLWorkbookMacroEnabled
    ElseIf Right(LCase(savePath), 4) = ".xls" Then
        newWb.SaveAs savePath, xlExcel8
    Else
        If InStr(savePath, ".") = 0 Then savePath = savePath & ".xlsx"
        newWb.SaveAs savePath, xlOpenXMLWorkbook
    End If
    On Error GoTo 0
    newWb.Close False
    MsgBox "导出完成！" & vbCrLf & savePath, vbInformation
End Sub

Public Sub GenerateFeeChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("费用管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then
        MsgBox "没有费用数据，无法生成图表！", vbExclamation
        Exit Sub
    End If
    ' 删除旧图表
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    ' 创建新柱状图
    Dim rng As Range
    Set rng = ws.Range("A1:I" & lr)
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=20, Top:=lr * 18 + 40, Width:=600, Height:=350)
    With cht.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "各栋楼欠缴费用分析"
        ' 手动构建数据系列
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "欠缴物业费"
        .SeriesCollection(1).Values = ws.Range("B2:B" & lr)
        .SeriesCollection(1).XValues = ws.Range("A2:A" & lr)
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "欠缴公摊费"
        .SeriesCollection(2).Values = ws.Range("D2:D" & lr)
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "楼栋/周期"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "金额(元)"
        .HasLegend = True
    End With
    ws.Activate
    MsgBox "柱状图已生成，请查看费用管理工作表！", vbInformation
End Sub

' ========== 停车管理模块 ==========

' ---------- 创建停车管理工作表 ----------

Private Sub CreateParkingSheet()
    Dim ws As Worksheet
    If SheetExists("停车管理") Then
        Set ws = ThisWorkbook.Sheets("停车管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "停车管理"
    End If
    With ws
        .Range("A1").Value = "车位编号"
        .Range("B1").Value = "车位位置"
        .Range("C1").Value = "状态"
        .Range("D1").Value = "楼号"
        .Range("E1").Value = "房号"
        .Range("F1").Value = "车主姓名"
        .Range("G1").Value = "车牌号"
        .Range("H1").Value = "联系电话"
        .Range("I1").Value = "月租费"
        .Range("J1").Value = "创建时间"
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.Color = RGB(70, 130, 180)
        .Range("A1:J1").Font.Color = RGB(255, 255, 255)
        .Columns("A:J").AutoFit
    End With
End Sub

' ---------- 创建停车管理主窗体 ----------

Private Sub CreateParkingManagerForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ParkingManagerFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ParkingManagerFormName", actualName

    vbc.Properties("Caption") = "停车管理"
    vbc.Properties("Width") = 780
    vbc.Properties("Height") = 560
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' ---- 顶部统计卡片 ----
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard1Title": ctl.Caption = "总车位数"
    ctl.Left = 15: ctl.Top = 8: ctl.Width = 170: ctl.Height = 16
    ctl.Font.Bold = True: ctl.Font.Size = 9: ctl.TextAlign = 2
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard1Val": ctl.Caption = "0"
    ctl.Left = 15: ctl.Top = 26: ctl.Width = 170: ctl.Height = 20
    ctl.Font.Size = 14: ctl.Font.Bold = True: ctl.TextAlign = 2
    ctl.ForeColor = RGB(30, 144, 255)

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard2Title": ctl.Caption = "空闲车位"
    ctl.Left = 200: ctl.Top = 8: ctl.Width = 170: ctl.Height = 16
    ctl.Font.Bold = True: ctl.Font.Size = 9: ctl.TextAlign = 2
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard2Val": ctl.Caption = "0"
    ctl.Left = 200: ctl.Top = 26: ctl.Width = 170: ctl.Height = 20
    ctl.Font.Size = 14: ctl.Font.Bold = True: ctl.TextAlign = 2
    ctl.ForeColor = RGB(46, 139, 87)

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard3Title": ctl.Caption = "已占用车位"
    ctl.Left = 385: ctl.Top = 8: ctl.Width = 170: ctl.Height = 16
    ctl.Font.Bold = True: ctl.Font.Size = 9: ctl.TextAlign = 2
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard3Val": ctl.Caption = "0"
    ctl.Left = 385: ctl.Top = 26: ctl.Width = 170: ctl.Height = 20
    ctl.Font.Size = 14: ctl.Font.Bold = True: ctl.TextAlign = 2
    ctl.ForeColor = RGB(220, 80, 80)

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard4Title": ctl.Caption = "月租金总额"
    ctl.Left = 570: ctl.Top = 8: ctl.Width = 170: ctl.Height = 16
    ctl.Font.Bold = True: ctl.Font.Size = 9: ctl.TextAlign = 2
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblCard4Val": ctl.Caption = "0.00"
    ctl.Left = 570: ctl.Top = 26: ctl.Width = 170: ctl.Height = 20
    ctl.Font.Size = 14: ctl.Font.Bold = True: ctl.TextAlign = 2
    ctl.ForeColor = RGB(255, 140, 0)

    ' ---- 筛选区域 ----
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblFilterStatus": ctl.Caption = "状态筛选:"
    ctl.Left = 15: ctl.Top = 55: ctl.Width = 60: ctl.Height = 16
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboFilterStatus"
    ctl.Left = 78: ctl.Top = 53: ctl.Width = 80: ctl.Height = 20
    ctl.Style = 2

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblFilterKey": ctl.Caption = "关键字:"
    ctl.Left = 170: ctl.Top = 55: ctl.Width = 50: ctl.Height = 16
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtFilterKey"
    ctl.Left = 222: ctl.Top = 53: ctl.Width = 120: ctl.Height = 20

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdFilter": ctl.Caption = "查询"
    ctl.Left = 350: ctl.Top = 53: ctl.Width = 50: ctl.Height = 20

    ' ---- 车位列表 ----
    Set ctl = dsg.Controls.Add("Forms.ListBox.1")
    ctl.Name = "lstParking"
    ctl.Left = 15: ctl.Top = 80: ctl.Width = 740: ctl.Height = 380
    ctl.ColumnCount = 9
    ctl.ColumnWidths = "65;80;50;50;50;70;80;90;70"
    ctl.ColumnHeads = False

    ' ---- 底部按钮 ----
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdNew": ctl.Caption = "添加车位"
    ctl.Left = 15: ctl.Top = 470: ctl.Width = 80: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdView": ctl.Caption = "查看"
    ctl.Left = 102: ctl.Top = 470: ctl.Width = 60: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdEdit": ctl.Caption = "编辑"
    ctl.Left = 169: ctl.Top = 470: ctl.Width = 60: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdDelete": ctl.Caption = "删除"
    ctl.Left = 236: ctl.Top = 470: ctl.Width = 60: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdImport": ctl.Caption = "导入"
    ctl.Left = 330: ctl.Top = 470: ctl.Width = 60: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdExport": ctl.Caption = "导出"
    ctl.Left = 397: ctl.Top = 470: ctl.Width = 60: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdRefresh": ctl.Caption = "刷新"
    ctl.Left = 490: ctl.Top = 470: ctl.Width = 60: ctl.Height = 28

    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 685: ctl.Top = 470: ctl.Width = 70: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboFilterStatus.AddItem " & q & "全部" & q & vbCrLf
    c = c & "    cboFilterStatus.AddItem " & q & "空闲" & q & vbCrLf
    c = c & "    cboFilterStatus.AddItem " & q & "占用" & q & vbCrLf
    c = c & "    cboFilterStatus.Text = " & q & "全部" & q & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub LoadData()" & vbCrLf
    c = c & "    lstParking.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "停车管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim filterSt As String" & vbCrLf
    c = c & "    filterSt = Trim(cboFilterStatus.Text)" & vbCrLf
    c = c & "    Dim filterKey As String" & vbCrLf
    c = c & "    filterKey = LCase(Trim(txtFilterKey.Text))" & vbCrLf
    c = c & "    Dim totalCnt As Long, freeCnt As Long, usedCnt As Long" & vbCrLf
    c = c & "    Dim totalRent As Double" & vbCrLf
    c = c & "    totalCnt = 0: freeCnt = 0: usedCnt = 0: totalRent = 0" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim st As String" & vbCrLf
    c = c & "        st = CStr(ws.Cells(i, 3).Value)" & vbCrLf
    c = c & "        totalCnt = totalCnt + 1" & vbCrLf
    c = c & "        If st = " & q & "空闲" & q & " Then freeCnt = freeCnt + 1" & vbCrLf
    c = c & "        If st = " & q & "占用" & q & " Then" & vbCrLf
    c = c & "            usedCnt = usedCnt + 1" & vbCrLf
    c = c & "            totalRent = totalRent + Val(ws.Cells(i, 9).Value)" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        Dim showRow As Boolean" & vbCrLf
    c = c & "        showRow = True" & vbCrLf
    c = c & "        If filterSt <> " & q & "全部" & q & " And filterSt <> " & q & q & " Then" & vbCrLf
    c = c & "            If st <> filterSt Then showRow = False" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        If showRow And filterKey <> " & q & q & " Then" & vbCrLf
    c = c & "            Dim rowText As String" & vbCrLf
    c = c & "            rowText = LCase(CStr(ws.Cells(i, 1).Value) & CStr(ws.Cells(i, 2).Value) & CStr(ws.Cells(i, 4).Value) & CStr(ws.Cells(i, 6).Value) & CStr(ws.Cells(i, 7).Value))" & vbCrLf
    c = c & "            If InStr(rowText, filterKey) = 0 Then showRow = False" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        If showRow Then" & vbCrLf
    c = c & "            lstParking.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lstParking.List(lstParking.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lstParking.List(lstParking.ListCount - 1, 2) = st" & vbCrLf
    c = c & "            lstParking.List(lstParking.ListCount - 1, 3) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lstParking.List(lstParking.ListCount - 1, 4) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lstParking.List(lstParking.ListCount - 1, 5) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            lstParking.List(lstParking.ListCount - 1, 6) = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            lstParking.List(lstParking.ListCount - 1, 7) = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            lstParking.List(lstParking.ListCount - 1, 8) = Format(Val(ws.Cells(i, 9).Value), " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    lblCard1Val.Caption = totalCnt" & vbCrLf
    c = c & "    lblCard2Val.Caption = freeCnt" & vbCrLf
    c = c & "    lblCard3Val.Caption = usedCnt" & vbCrLf
    c = c & "    lblCard4Val.Caption = Format(totalRent, " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdFilter_Click()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNew_Click()" & vbCrLf
    c = c & "    gEditParkingRow = 0" & vbCrLf
    c = c & "    ShowParkingEditForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdView_Click()" & vbCrLf
    c = c & "    If lstParking.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条车位记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewParkingID = lstParking.List(lstParking.ListIndex, 0)" & vbCrLf
    c = c & "    ShowParkingViewForm" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEdit_Click()" & vbCrLf
    c = c & "    If lstParking.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条车位记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim selID As String" & vbCrLf
    c = c & "    selID = lstParking.List(lstParking.ListIndex, 0)" & vbCrLf
    c = c & "    gEditParkingRow = FindParkingRow(selID)" & vbCrLf
    c = c & "    If gEditParkingRow > 0 Then" & vbCrLf
    c = c & "        ShowParkingEditForm" & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelete_Click()" & vbCrLf
    c = c & "    If lstParking.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条车位记录！" & q & ", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该车位记录？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        DeleteParkingByID lstParking.List(lstParking.ListIndex, 0)" & vbCrLf
    c = c & "        LoadData" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdImport_Click()" & vbCrLf
    c = c & "    Dim fd As Object" & vbCrLf
    c = c & "    Set fd = Application.FileDialog(1)" & vbCrLf
    c = c & "    fd.Title = " & q & "选择导入文件" & q & vbCrLf
    c = c & "    fd.Filters.Clear" & vbCrLf
    c = c & "    fd.Filters.Add " & q & "Excel文件" & q & ", " & q & "*.xlsm;*.xlsx;*.xls" & q & vbCrLf
    c = c & "    If fd.Show <> -1 Then Exit Sub" & vbCrLf
    c = c & "    Dim fPath As String" & vbCrLf
    c = c & "    fPath = fd.SelectedItems(1)" & vbCrLf
    c = c & "    ImportParkingData fPath" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "    MsgBox " & q & "导入完成！" & q & ", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdExport_Click()" & vbCrLf
    c = c & "    ExportParkingData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdRefresh_Click()" & vbCrLf
    c = c & "    LoadData" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建停车编辑窗体 ----------

Private Sub CreateParkingEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ParkingEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ParkingEditFormName", actualName

    vbc.Properties("Caption") = "车位信息编辑"
    vbc.Properties("Width") = 420
    vbc.Properties("Height") = 400
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 18

    ' 车位编号
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblParkID": ctl.Caption = "车位编号："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtParkID"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 130: ctl.Height = 20

    ' 车位位置
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblLocation": ctl.Caption = "车位位置："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtLocation"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 280: ctl.Height = 20

    ' 状态
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblStatus": ctl.Caption = "状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboStatus"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20
    ctl.Style = 2

    ' 楼号
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblBuilding": ctl.Caption = "楼号："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboBuilding"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20
    ctl.Style = 0

    ' 房号
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRoom": ctl.Caption = "房号："
    ctl.Left = 230: ctl.Top = yy: ctl.Width = 50: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtRoom"
    ctl.Left = 285: ctl.Top = yy: ctl.Width = 105: ctl.Height = 20

    ' 车主姓名
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblOwner": ctl.Caption = "车主姓名："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtOwner"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 130: ctl.Height = 20

    ' 车牌号
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPlate": ctl.Caption = "车牌号："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtPlate"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 130: ctl.Height = 20

    ' 联系电话
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblPhone": ctl.Caption = "联系电话："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtPhone"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 130: ctl.Height = 20

    ' 月租费
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblRent": ctl.Caption = "月租费(元)："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 80: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtRent"
    ctl.Left = 110: ctl.Top = yy: ctl.Width = 130: ctl.Height = 20

    ' 提示信息
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblTip"
    ctl.Caption = "* 状态为空闲时，楼号/房号/车主/车牌/电话/月租费可不填"
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 380: ctl.Height = 16
    ctl.ForeColor = RGB(128, 128, 128): ctl.Font.Size = 8

    ' 按钮
    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 210: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "空闲" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "占用" & q & vbCrLf
    c = c & "    cboStatus.Text = " & q & "空闲" & q & vbCrLf
    c = c & "    Dim bk As Long" & vbCrLf
    c = c & "    For bk = 1 To 10" & vbCrLf
    c = c & "        cboBuilding.AddItem bk & " & q & "栋" & q & vbCrLf
    c = c & "    Next bk" & vbCrLf
    c = c & "    If gEditParkingRow > 0 Then" & vbCrLf
    c = c & "        Me.Caption = " & q & "编辑车位信息" & q & vbCrLf
    c = c & "        Dim ws As Worksheet" & vbCrLf
    c = c & "        Set ws = ThisWorkbook.Sheets(" & q & "停车管理" & q & ")" & vbCrLf
    c = c & "        txtParkID.Text = ws.Cells(gEditParkingRow, 1).Value" & vbCrLf
    c = c & "        txtParkID.Enabled = False" & vbCrLf
    c = c & "        txtLocation.Text = ws.Cells(gEditParkingRow, 2).Value" & vbCrLf
    c = c & "        cboStatus.Text = ws.Cells(gEditParkingRow, 3).Value" & vbCrLf
    c = c & "        cboBuilding.Text = ws.Cells(gEditParkingRow, 4).Value" & vbCrLf
    c = c & "        txtRoom.Text = ws.Cells(gEditParkingRow, 5).Value" & vbCrLf
    c = c & "        txtOwner.Text = ws.Cells(gEditParkingRow, 6).Value" & vbCrLf
    c = c & "        txtPlate.Text = ws.Cells(gEditParkingRow, 7).Value" & vbCrLf
    c = c & "        txtPhone.Text = ws.Cells(gEditParkingRow, 8).Value" & vbCrLf
    c = c & "        txtRent.Text = ws.Cells(gEditParkingRow, 9).Value" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Me.Caption = " & q & "添加车位" & q & vbCrLf
    c = c & "        txtParkID.Text = GenerateParkingID()" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cboStatus_Change()" & vbCrLf
    c = c & "    Dim isFree As Boolean" & vbCrLf
    c = c & "    isFree = (cboStatus.Text = " & q & "空闲" & q & ")" & vbCrLf
    c = c & "    If isFree Then" & vbCrLf
    c = c & "        cboBuilding.Text = " & q & q & vbCrLf
    c = c & "        txtRoom.Text = " & q & q & vbCrLf
    c = c & "        txtOwner.Text = " & q & q & vbCrLf
    c = c & "        txtPlate.Text = " & q & q & vbCrLf
    c = c & "        txtPhone.Text = " & q & q & vbCrLf
    c = c & "        txtRent.Text = " & q & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(txtParkID.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写车位编号！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Trim(txtLocation.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写车位位置！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If cboStatus.Text = " & q & "占用" & q & " Then" & vbCrLf
    c = c & "        If Trim(txtOwner.Text) = " & q & q & " Then" & vbCrLf
    c = c & "            MsgBox " & q & "占用状态请填写车主姓名！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        If Trim(txtPlate.Text) = " & q & q & " Then" & vbCrLf
    c = c & "            MsgBox " & q & "占用状态请填写车牌号！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If gEditParkingRow = 0 Then" & vbCrLf
    c = c & "        Dim chkRow As Long" & vbCrLf
    c = c & "        chkRow = FindParkingRow(Trim(txtParkID.Text))" & vbCrLf
    c = c & "        If chkRow > 0 Then" & vbCrLf
    c = c & "            MsgBox " & q & "车位编号已存在！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    SaveParkingRecord gEditParkingRow, txtParkID.Text, txtLocation.Text, cboStatus.Text, cboBuilding.Text, txtRoom.Text, txtOwner.Text, txtPlate.Text, txtPhone.Text, Val(txtRent.Text)" & vbCrLf
    c = c & "    MsgBox " & q & "车位信息已保存！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建停车查看窗体 ----------

Private Sub CreateParkingViewForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ParkingViewFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ParkingViewFormName", actualName

    vbc.Properties("Caption") = "车位详情"
    vbc.Properties("Width") = 400
    vbc.Properties("Height") = 380
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 18

    Dim labels As Variant
    labels = Array("车位编号：", "车位位置：", "状态：", "楼号：", "房号：", "车主姓名：", "车牌号：", "联系电话：", "月租费(元)：", "创建时间：")
    Dim valNames As Variant
    valNames = Array("lblValID", "lblValLoc", "lblValSt", "lblValBld", "lblValRoom", "lblValOwner", "lblValPlate", "lblValPhone", "lblValRent", "lblValTime")

    Dim idx As Long
    For idx = 0 To 9
        Set ctl = dsg.Controls.Add("Forms.Label.1")
        ctl.Name = "lblTitle" & idx: ctl.Caption = labels(idx)
        ctl.Left = 20: ctl.Top = yy: ctl.Width = 90: ctl.Height = 18
        ctl.Font.Bold = True

        Set ctl = dsg.Controls.Add("Forms.Label.1")
        ctl.Name = valNames(idx): ctl.Caption = ""
        ctl.Left = 120: ctl.Top = yy: ctl.Width = 250: ctl.Height = 18
        ctl.ForeColor = RGB(50, 50, 120)

        yy = yy + 28
    Next idx

    yy = yy + 10
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 150: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    If gViewParkingID = " & q & q & " Then Unload Me: Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "停车管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gViewParkingID Then" & vbCrLf
    c = c & "            lblValID.Caption = ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lblValLoc.Caption = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lblValSt.Caption = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            If ws.Cells(i, 3).Value = " & q & "空闲" & q & " Then" & vbCrLf
    c = c & "                lblValSt.ForeColor = RGB(46, 139, 87)" & vbCrLf
    c = c & "            Else" & vbCrLf
    c = c & "                lblValSt.ForeColor = RGB(220, 80, 80)" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            lblValBld.Caption = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lblValRoom.Caption = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lblValOwner.Caption = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            lblValPlate.Caption = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            lblValPhone.Caption = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "            lblValRent.Caption = Format(Val(ws.Cells(i, 9).Value), " & q & "#,##0.00" & q & ")" & vbCrLf
    c = c & "            lblValTime.Caption = ws.Cells(i, 10).Value" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 停车管理窗体显示 ----------

Public Sub ShowParkingManagerForm()
    Dim fName As String
    fName = GetConfigProp("ParkingManagerFormName")
    If fName = "" Or Not FormExists("ParkingManagerFormName") Then
        MsgBox "停车管理窗体不存在！请重新运行 SetupOASystem", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ParkingManagerFormName")
    frm.Show
End Sub

Public Sub ShowParkingEditForm()
    Dim fName As String
    fName = GetConfigProp("ParkingEditFormName")
    If fName = "" Or Not FormExists("ParkingEditFormName") Then
        MsgBox "车位编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ParkingEditFormName")
    frm.Show
End Sub

Public Sub ShowParkingViewForm()
    Dim fName As String
    fName = GetConfigProp("ParkingViewFormName")
    If fName = "" Or Not FormExists("ParkingViewFormName") Then
        MsgBox "车位查看窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ParkingViewFormName")
    frm.Show
End Sub

' ---------- 停车管理数据操作函数 ----------

Public Function GenerateParkingID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("停车管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "CW" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateParkingID = "CW" & Format(maxID + 1, "0000")
End Function

Public Function FindParkingRow(parkID As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("停车管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    FindParkingRow = 0
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = parkID Then
            FindParkingRow = i
            Exit Function
        End If
    Next i
End Function

Public Sub SaveParkingRecord(rowNum As Long, sID As String, sLoc As String, sSt As String, sBld As String, sRoom As String, sOwner As String, sPlate As String, sPhone As String, dRent As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("停车管理")
    Dim r As Long
    If rowNum > 0 Then
        r = rowNum
    Else
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If
    ws.Cells(r, 1).Value = sID
    ws.Cells(r, 2).Value = sLoc
    ws.Cells(r, 3).Value = sSt
    ws.Cells(r, 4).Value = sBld
    ws.Cells(r, 5).Value = sRoom
    ws.Cells(r, 6).Value = sOwner
    ws.Cells(r, 7).Value = sPlate
    ws.Cells(r, 8).Value = sPhone
    ws.Cells(r, 9).Value = dRent
    If rowNum = 0 Then
        ws.Cells(r, 10).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    End If
    ws.Columns("A:J").AutoFit
End Sub

Public Sub DeleteParkingByID(parkID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("停车管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = parkID Then
            ws.Rows(i).Delete
            Exit Sub
        End If
    Next i
End Sub

Public Sub ImportParkingData(fPath As String)
    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(fPath, ReadOnly:=True)
    Dim srcWs As Worksheet
    Set srcWs = srcWb.Sheets(1)
    Dim dstWs As Worksheet
    Set dstWs = ThisWorkbook.Sheets("停车管理")
    Dim srcLr As Long, i As Long
    srcLr = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    Dim startRow As Long
    If CStr(srcWs.Cells(1, 1).Value) = "车位编号" Then startRow = 2 Else startRow = 1
    For i = startRow To srcLr
        Dim sID As String
        sID = CStr(srcWs.Cells(i, 1).Value)
        If sID = "" Then GoTo nextParkRow
        Dim existRow As Long
        existRow = FindParkingRow(sID)
        SaveParkingRecord existRow, _
            sID, _
            CStr(srcWs.Cells(i, 2).Value), _
            CStr(srcWs.Cells(i, 3).Value), _
            CStr(srcWs.Cells(i, 4).Value), _
            CStr(srcWs.Cells(i, 5).Value), _
            CStr(srcWs.Cells(i, 6).Value), _
            CStr(srcWs.Cells(i, 7).Value), _
            CStr(srcWs.Cells(i, 8).Value), _
            Val(srcWs.Cells(i, 9).Value)
nextParkRow:
    Next i
    srcWb.Close False
End Sub

Public Sub ExportParkingData()
    Dim fd As Object
    Set fd = Application.FileDialog(2)
    fd.Title = "选择导出路径"
    fd.InitialFileName = "停车管理导出_" & Format(Now, "yyyymmdd") & ".xlsx"
    fd.FilterIndex = 1
    If fd.Show <> -1 Then Exit Sub
    Dim savePath As String
    savePath = fd.SelectedItems(1)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("停车管理")
    ws.Copy
    Dim newWb As Workbook
    Set newWb = ActiveWorkbook
    On Error Resume Next
    If Right(LCase(savePath), 5) = ".xlsm" Then
        newWb.SaveAs savePath, xlOpenXMLWorkbookMacroEnabled
    ElseIf Right(LCase(savePath), 4) = ".xls" Then
        newWb.SaveAs savePath, xlExcel8
    Else
        If InStr(savePath, ".") = 0 Then savePath = savePath & ".xlsx"
        newWb.SaveAs savePath, xlOpenXMLWorkbook
    End If
    On Error GoTo 0
    newWb.Close False
    MsgBox "导出完成！" & vbCrLf & savePath, vbInformation
End Sub

' ========== 人力资源模块 ==========

' ---------- 创建考勤审批表 ----------

Private Sub CreateAttendApplySheet()
    Dim ws As Worksheet
    If SheetExists("考勤审批") Then
        Set ws = ThisWorkbook.Sheets("考勤审批")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "考勤审批"
    End If
    With ws
        .Range("A1").Value = "申请编号"
        .Range("B1").Value = "申请人"
        .Range("C1").Value = "考勤日期"
        .Range("D1").Value = "原考勤状态"
        .Range("E1").Value = "申请理由"
        .Range("F1").Value = "审批状态"
        .Range("G1").Value = "审批回复"
        .Range("H1").Value = "申请时间"
        .Range("A1:H1").Font.Bold = True
        .Columns("A:H").AutoFit
    End With
End Sub

' ---------- 创建排班管理表 ----------

Private Sub CreateScheduleSheet()
    Dim ws As Worksheet
    If SheetExists("排班管理") Then
        Set ws = ThisWorkbook.Sheets("排班管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "排班管理"
    End If
    With ws
        .Range("A1").Value = "姓名"
        .Range("B1").Value = "排班时间"
        .Range("C1").Value = "工作地点"
        .Range("D1").Value = "备注"
        .Range("E1").Value = "创建时间"
        .Range("A1:E1").Font.Bold = True
        .Columns("A:E").AutoFit
    End With
End Sub

' ---------- 创建人事管理表 ----------

Private Sub CreatePersonnelSheet()
    Dim ws As Worksheet
    If SheetExists("人事管理") Then
        Set ws = ThisWorkbook.Sheets("人事管理")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "人事管理"
    End If
    With ws
        .Range("A1").Value = "姓名"
        .Range("B1").Value = "性别"
        .Range("C1").Value = "年龄"
        .Range("D1").Value = "部门"
        .Range("E1").Value = "职务"
        .Range("F1").Value = "入职时间"
        .Range("G1").Value = "在职状态"
        .Range("H1").Value = "备注"
        .Range("A1:H1").Font.Bold = True
        .Columns("A:H").AutoFit
    End With
End Sub

' ---------- 创建考勤统计表 ----------

Private Sub CreateAttendStatsSheet()
    Dim ws As Worksheet
    If SheetExists("考勤统计") Then
        Set ws = ThisWorkbook.Sheets("考勤统计")
    Else
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "考勤统计"
    End If
    With ws
        .Range("A1").Value = "日期"
        .Range("B1").Value = "姓名"
        .Range("C1").Value = "角色"
        .Range("D1").Value = "出勤天数"
        .Range("E1").Value = "迟到"
        .Range("F1").Value = "早退"
        .Range("G1").Value = "请假"
        .Range("H1").Value = "缺勤"
        .Range("I1").Value = "备注"
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(70, 130, 180)
        .Range("A1:I1").Font.Color = RGB(255, 255, 255)
        .Columns("A:I").AutoFit
    End With
End Sub

' ---------- 刷新考勤统计数据 ----------

Public Sub RefreshAttendStats()
    If Not SheetExists("考勤管理") Then
        MsgBox "未找到考勤管理表!", vbExclamation
        Exit Sub
    End If
    If Not SheetExists("考勤统计") Then CreateAttendStatsSheet

    Dim wsKQ As Worksheet, wsStat As Worksheet, wsUser As Worksheet
    Set wsKQ = ThisWorkbook.Sheets("考勤管理")
    Set wsStat = ThisWorkbook.Sheets("考勤统计")

    ' 清除旧数据(保留表头)
    Dim lastClear As Long
    lastClear = wsStat.Cells(wsStat.Rows.Count, 1).End(xlUp).Row
    If lastClear >= 2 Then wsStat.Range("A2:I" & lastClear).ClearContents

    ' 构建用户角色字典
    Dim roleDict As Object
    Set roleDict = CreateObject("Scripting.Dictionary")
    If SheetExists("用户管理") Then
        Set wsUser = ThisWorkbook.Sheets("用户管理")
        Dim ru As Long
        For ru = 2 To wsUser.Cells(wsUser.Rows.Count, 1).End(xlUp).Row
            Dim uKey As String
            uKey = Trim(CStr(wsUser.Cells(ru, 1).Value))
            If uKey <> "" And Not roleDict.Exists(uKey) Then
                roleDict.Add uKey, CStr(wsUser.Cells(ru, 4).Value)
            End If
        Next ru
    End If

    ' 聚合考勤数据: key = "月份|姓名"
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim lr As Long, i As Long
    lr = wsKQ.Cells(wsKQ.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        Dim recDate As String
        If IsDate(wsKQ.Cells(i, 3).Value) Then
            recDate = Format(wsKQ.Cells(i, 3).Value, "yyyy-mm")
        Else
            recDate = Left(CStr(wsKQ.Cells(i, 3).Value), 7)
        End If
        Dim uName As String
        uName = Trim(CStr(wsKQ.Cells(i, 2).Value))
        If uName = "" Then GoTo NextRow
        Dim dKey As String
        dKey = recDate & "|" & uName
        If Not dict.Exists(dKey) Then
            ' Array: 出勤, 迟到, 早退, 请假, 缺勤
            dict.Add dKey, Array(0, 0, 0, 0, 0)
        End If
        Dim arr As Variant
        arr = dict(dKey)
        Dim aSt As String
        aSt = Trim(CStr(wsKQ.Cells(i, 6).Value))
        If aSt = "出勤" Then arr(0) = arr(0) + 1
        If aSt = "迟到" Then arr(1) = arr(1) + 1
        If aSt = "早退" Then arr(2) = arr(2) + 1
        If aSt = "请假" Then arr(3) = arr(3) + 1
        If aSt = "缺勤" Then arr(4) = arr(4) + 1
        dict(dKey) = arr
NextRow:
    Next i

    ' 写入统计表
    Dim outRow As Long
    outRow = 2
    Dim k As Variant
    For Each k In dict.Keys
        arr = dict(k)
        Dim parts As Variant
        parts = Split(CStr(k), "|")
        wsStat.Cells(outRow, 1).Value = parts(0)  ' 日期(月份)
        wsStat.Cells(outRow, 2).Value = parts(1)  ' 姓名
        If roleDict.Exists(parts(1)) Then
            wsStat.Cells(outRow, 3).Value = roleDict(parts(1))
        Else
            wsStat.Cells(outRow, 3).Value = ""
        End If
        wsStat.Cells(outRow, 4).Value = arr(0)    ' 出勤天数
        wsStat.Cells(outRow, 5).Value = arr(1)    ' 迟到
        wsStat.Cells(outRow, 6).Value = arr(2)    ' 早退
        wsStat.Cells(outRow, 7).Value = arr(3)    ' 请假
        wsStat.Cells(outRow, 8).Value = arr(4)    ' 缺勤
        wsStat.Cells(outRow, 9).Value = ""         ' 备注
        outRow = outRow + 1
    Next k
    wsStat.Columns("A:I").AutoFit
End Sub

' ---------- 创建人力资源主窗体 ----------

Private Sub CreateHRMainForm()
    Dim oldForm As String
    oldForm = GetConfigProp("HRMainFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "HRMainFormName", actualName

    vbc.Properties("Caption") = "人力资源"
    vbc.Properties("Width") = 780
    vbc.Properties("Height") = 560
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object

    ' MultiPage: 4个标签页
    Set ctl = dsg.Controls.Add("Forms.MultiPage.1")
    ctl.Name = "mpHR"
    ctl.Left = 5: ctl.Top = 5: ctl.Width = 760: ctl.Height = 490
    ctl.Pages(0).Caption = "考勤管理"
    ctl.Pages.Add: ctl.Pages(1).Caption = "考勤审批"
    ctl.Pages.Add: ctl.Pages(2).Caption = "排班管理"
    ctl.Pages.Add: ctl.Pages(3).Caption = "人事管理"

    Dim pg As Object, subCtl As Object

    ' ==== 第0页: 考勤管理 ====
    Set pg = ctl.Pages(0)

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblAttMonth": subCtl.Caption = "查询月份:"
    subCtl.Left = 10: subCtl.Top = 8: subCtl.Width = 65: subCtl.Height = 16
    Set subCtl = pg.Controls.Add("Forms.TextBox.1")
    subCtl.Name = "txtAttMonth"
    subCtl.Left = 78: subCtl.Top = 6: subCtl.Width = 80: subCtl.Height = 20
    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdAttQuery": subCtl.Caption = "查询"
    subCtl.Left = 165: subCtl.Top = 6: subCtl.Width = 50: subCtl.Height = 20

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdRefreshStats": subCtl.Caption = "刷新统计表"
    subCtl.Left = 225: subCtl.Top = 6: subCtl.Width = 90: subCtl.Height = 20

    Set subCtl = pg.Controls.Add("Forms.ListBox.1")
    subCtl.Name = "lstAttMgr"
    subCtl.Left = 10: subCtl.Top = 32: subCtl.Width = 730: subCtl.Height = 400
    subCtl.ColumnCount = 7
    subCtl.ColumnWidths = "80;80;80;80;60;60;60"

    ' ==== 第1页: 考勤审批 ====
    Set pg = ctl.Pages(1)

    Set subCtl = pg.Controls.Add("Forms.ListBox.1")
    subCtl.Name = "lstAttApply"
    subCtl.Left = 10: subCtl.Top = 8: subCtl.Width = 730: subCtl.Height = 360
    subCtl.ColumnCount = 6
    subCtl.ColumnWidths = "60;70;80;70;60;120"

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdNewApply": subCtl.Caption = "新建申请"
    subCtl.Left = 10: subCtl.Top = 378: subCtl.Width = 90: subCtl.Height = 26

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdViewApply": subCtl.Caption = "查看/审批"
    subCtl.Left = 110: subCtl.Top = 378: subCtl.Width = 90: subCtl.Height = 26

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdDelApply": subCtl.Caption = "删除"
    subCtl.Left = 210: subCtl.Top = 378: subCtl.Width = 70: subCtl.Height = 26

    ' ==== 第2页: 排班管理 ====
    Set pg = ctl.Pages(2)

    Set subCtl = pg.Controls.Add("Forms.Label.1")
    subCtl.Name = "lblSchQuery": subCtl.Caption = "查询:"
    subCtl.Left = 10: subCtl.Top = 8: subCtl.Width = 35: subCtl.Height = 16
    Set subCtl = pg.Controls.Add("Forms.TextBox.1")
    subCtl.Name = "txtSchQuery"
    subCtl.Left = 48: subCtl.Top = 6: subCtl.Width = 100: subCtl.Height = 20
    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdSchSearch": subCtl.Caption = "搜索"
    subCtl.Left = 155: subCtl.Top = 6: subCtl.Width = 50: subCtl.Height = 20

    Set subCtl = pg.Controls.Add("Forms.ListBox.1")
    subCtl.Name = "lstSchedule"
    subCtl.Left = 10: subCtl.Top = 32: subCtl.Width = 730: subCtl.Height = 360
    subCtl.ColumnCount = 5
    subCtl.ColumnWidths = "100;120;150;150;120"

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdNewSch": subCtl.Caption = "新建排班"
    subCtl.Left = 10: subCtl.Top = 400: subCtl.Width = 90: subCtl.Height = 26
    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdEditSch": subCtl.Caption = "编辑"
    subCtl.Left = 110: subCtl.Top = 400: subCtl.Width = 70: subCtl.Height = 26
    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdDelSch": subCtl.Caption = "删除"
    subCtl.Left = 190: subCtl.Top = 400: subCtl.Width = 70: subCtl.Height = 26

    ' ==== 第3页: 人事管理 ====
    Set pg = ctl.Pages(3)

    Set subCtl = pg.Controls.Add("Forms.ListBox.1")
    subCtl.Name = "lstPersonnel"
    subCtl.Left = 10: subCtl.Top = 8: subCtl.Width = 730: subCtl.Height = 370
    subCtl.ColumnCount = 8
    subCtl.ColumnWidths = "70;40;40;80;70;80;60;120"

    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdNewPers": subCtl.Caption = "新建档案"
    subCtl.Left = 10: subCtl.Top = 388: subCtl.Width = 90: subCtl.Height = 26
    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdEditPers": subCtl.Caption = "编辑"
    subCtl.Left = 110: subCtl.Top = 388: subCtl.Width = 70: subCtl.Height = 26
    Set subCtl = pg.Controls.Add("Forms.CommandButton.1")
    subCtl.Name = "cmdDelPers": subCtl.Caption = "删除"
    subCtl.Left = 190: subCtl.Top = 388: subCtl.Width = 70: subCtl.Height = 26

    ' 关闭按钮
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 685: ctl.Top = 505: ctl.Width = 70: ctl.Height = 26

    ' 注入事件代码
    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    txtAttMonth.Text = Format(Date, " & q & "yyyy-mm" & q & ")" & vbCrLf
    c = c & "    LoadAttendMgr" & vbCrLf
    c = c & "    LoadAttendApply" & vbCrLf
    c = c & "    LoadSchedule " & q & q & vbCrLf
    c = c & "    LoadPersonnel" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    ' ---- 考勤管理 ----
    c = c & "Private Sub LoadAttendMgr()" & vbCrLf
    c = c & "    lstAttMgr.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "考勤管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    Dim qm As String" & vbCrLf
    c = c & "    qm = Trim(txtAttMonth.Text)" & vbCrLf
    c = c & "    Dim dict As Object" & vbCrLf
    c = c & "    Set dict = CreateObject(" & q & "Scripting.Dictionary" & q & ")" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim recDate As String" & vbCrLf
    c = c & "        If IsDate(ws.Cells(i, 3).Value) Then" & vbCrLf
    c = c & "            recDate = Format(ws.Cells(i, 3).Value, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "        Else" & vbCrLf
    c = c & "            recDate = CStr(ws.Cells(i, 3).Value)" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        If qm = " & q & q & " Or Left(recDate, 7) = qm Then" & vbCrLf
    c = c & "            Dim uName As String" & vbCrLf
    c = c & "            uName = CStr(ws.Cells(i, 2).Value)" & vbCrLf
    c = c & "            If Not dict.Exists(uName) Then" & vbCrLf
    c = c & "                dict.Add uName, Array(0, 0, 0, 0, 0)" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Dim arr As Variant" & vbCrLf
    c = c & "            arr = dict(uName)" & vbCrLf
    c = c & "            Dim aSt As String" & vbCrLf
    c = c & "            aSt = CStr(ws.Cells(i, 6).Value)" & vbCrLf
    c = c & "            If aSt = " & q & "出勤" & q & " Then arr(0) = arr(0) + 1" & vbCrLf
    c = c & "            If aSt = " & q & "迟到" & q & " Then arr(1) = arr(1) + 1" & vbCrLf
    c = c & "            If aSt = " & q & "早退" & q & " Then arr(2) = arr(2) + 1" & vbCrLf
    c = c & "            If aSt = " & q & "请假" & q & " Then arr(3) = arr(3) + 1" & vbCrLf
    c = c & "            If aSt = " & q & "缺勤" & q & " Then arr(4) = arr(4) + 1" & vbCrLf
    c = c & "            dict(uName) = arr" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    Dim k As Variant" & vbCrLf
    c = c & "    For Each k In dict.Keys" & vbCrLf
    c = c & "        arr = dict(k)" & vbCrLf
    c = c & "        Dim total As Long" & vbCrLf
    c = c & "        total = arr(0) + arr(1) + arr(2) + arr(3) + arr(4)" & vbCrLf
    c = c & "        lstAttMgr.AddItem CStr(k)" & vbCrLf
    c = c & "        lstAttMgr.List(lstAttMgr.ListCount - 1, 1) = qm" & vbCrLf
    c = c & "        lstAttMgr.List(lstAttMgr.ListCount - 1, 2) = arr(0)" & vbCrLf
    c = c & "        lstAttMgr.List(lstAttMgr.ListCount - 1, 3) = arr(1)" & vbCrLf
    c = c & "        lstAttMgr.List(lstAttMgr.ListCount - 1, 4) = arr(2)" & vbCrLf
    c = c & "        lstAttMgr.List(lstAttMgr.ListCount - 1, 5) = arr(3)" & vbCrLf
    c = c & "        lstAttMgr.List(lstAttMgr.ListCount - 1, 6) = arr(4)" & vbCrLf
    c = c & "    Next k" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdAttQuery_Click()" & vbCrLf
    c = c & "    LoadAttendMgr" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdRefreshStats_Click()" & vbCrLf
    c = c & "    RefreshAttendStats" & vbCrLf
    c = c & "    MsgBox " & q & "考勤统计表已刷新！" & q & ", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    ' ---- 考勤审批 ----
    c = c & "Private Sub LoadAttendApply()" & vbCrLf
    c = c & "    lstAttApply.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "考勤审批" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        lstAttApply.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "        lstAttApply.List(lstAttApply.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "        lstAttApply.List(lstAttApply.ListCount - 1, 2) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "        lstAttApply.List(lstAttApply.ListCount - 1, 3) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "        lstAttApply.List(lstAttApply.ListCount - 1, 4) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "        lstAttApply.List(lstAttApply.ListCount - 1, 5) = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNewApply_Click()" & vbCrLf
    c = c & "    gEditAttendApplyID = " & q & q & vbCrLf
    c = c & "    ShowAttendApplyEditForm" & vbCrLf
    c = c & "    LoadAttendApply" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdViewApply_Click()" & vbCrLf
    c = c & "    If lstAttApply.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条申请！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gViewAttendApplyID = lstAttApply.List(lstAttApply.ListIndex, 0)" & vbCrLf
    c = c & "    ShowAttendApplyViewForm" & vbCrLf
    c = c & "    LoadAttendApply" & vbCrLf
    c = c & "    LoadAttendMgr" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelApply_Click()" & vbCrLf
    c = c & "    If lstAttApply.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条申请！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        DeleteAttendApplyByID lstAttApply.List(lstAttApply.ListIndex, 0)" & vbCrLf
    c = c & "        LoadAttendApply" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    ' ---- 排班管理 ----
    c = c & "Private Sub LoadSchedule(keyword As String)" & vbCrLf
    c = c & "    lstSchedule.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "排班管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        Dim matched As Boolean" & vbCrLf
    c = c & "        matched = True" & vbCrLf
    c = c & "        If keyword <> " & q & q & " Then" & vbCrLf
    c = c & "            Dim rowText As String" & vbCrLf
    c = c & "            rowText = ws.Cells(i, 1).Value & ws.Cells(i, 2).Value & ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            If InStr(1, rowText, keyword, vbTextCompare) = 0 Then matched = False" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "        If matched Then" & vbCrLf
    c = c & "            lstSchedule.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "            lstSchedule.List(lstSchedule.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lstSchedule.List(lstSchedule.ListCount - 1, 2) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            lstSchedule.List(lstSchedule.ListCount - 1, 3) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            lstSchedule.List(lstSchedule.ListCount - 1, 4) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSchSearch_Click()" & vbCrLf
    c = c & "    LoadSchedule txtSchQuery.Text" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNewSch_Click()" & vbCrLf
    c = c & "    gEditScheduleRow = 0" & vbCrLf
    c = c & "    ShowScheduleEditForm" & vbCrLf
    c = c & "    LoadSchedule " & q & q & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEditSch_Click()" & vbCrLf
    c = c & "    If lstSchedule.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条排班！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditScheduleRow = FindScheduleRow(lstSchedule.List(lstSchedule.ListIndex, 0), lstSchedule.List(lstSchedule.ListIndex, 1))" & vbCrLf
    c = c & "    If gEditScheduleRow > 0 Then ShowScheduleEditForm" & vbCrLf
    c = c & "    LoadSchedule " & q & q & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelSch_Click()" & vbCrLf
    c = c & "    If lstSchedule.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条排班！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        Dim rw As Long" & vbCrLf
    c = c & "        rw = FindScheduleRow(lstSchedule.List(lstSchedule.ListIndex, 0), lstSchedule.List(lstSchedule.ListIndex, 1))" & vbCrLf
    c = c & "        If rw > 0 Then ThisWorkbook.Sheets(" & q & "排班管理" & q & ").Rows(rw).Delete" & vbCrLf
    c = c & "        LoadSchedule " & q & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    ' ---- 人事管理 ----
    c = c & "Private Sub LoadPersonnel()" & vbCrLf
    c = c & "    lstPersonnel.Clear" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "人事管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        lstPersonnel.AddItem ws.Cells(i, 1).Value" & vbCrLf
    c = c & "        lstPersonnel.List(lstPersonnel.ListCount - 1, 1) = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "        lstPersonnel.List(lstPersonnel.ListCount - 1, 2) = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "        lstPersonnel.List(lstPersonnel.ListCount - 1, 3) = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "        lstPersonnel.List(lstPersonnel.ListCount - 1, 4) = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "        lstPersonnel.List(lstPersonnel.ListCount - 1, 5) = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "        lstPersonnel.List(lstPersonnel.ListCount - 1, 6) = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "        lstPersonnel.List(lstPersonnel.ListCount - 1, 7) = ws.Cells(i, 8).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdNewPers_Click()" & vbCrLf
    c = c & "    gEditPersonnelRow = 0" & vbCrLf
    c = c & "    ShowPersonnelEditForm" & vbCrLf
    c = c & "    LoadPersonnel" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdEditPers_Click()" & vbCrLf
    c = c & "    If lstPersonnel.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条记录！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    gEditPersonnelRow = lstPersonnel.ListIndex + 2" & vbCrLf
    c = c & "    ShowPersonnelEditForm" & vbCrLf
    c = c & "    LoadPersonnel" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdDelPers_Click()" & vbCrLf
    c = c & "    If lstPersonnel.ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox " & q & "请先选择一条记录！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If MsgBox(" & q & "确定删除该人事档案？" & q & ", vbYesNo + vbQuestion) = vbYes Then" & vbCrLf
    c = c & "        ThisWorkbook.Sheets(" & q & "人事管理" & q & ").Rows(lstPersonnel.ListIndex + 2).Delete" & vbCrLf
    c = c & "        LoadPersonnel" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建考勤审批编辑窗体 ----------

Private Sub CreateAttendApplyEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("AttendApplyEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "AttendApplyEditFormName", actualName

    vbc.Properties("Caption") = "考勤异常申请"
    vbc.Properties("Width") = 400
    vbc.Properties("Height") = 260
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 18

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl1": ctl.Caption = "考勤日期："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtDate"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 120: ctl.Height = 20

    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl2": ctl.Caption = "考勤状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboStatus"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 120: ctl.Height = 20
    ctl.Style = 2

    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl3": ctl.Caption = "申请理由："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtReason"
    ctl.Left = 100: ctl.Top = yy: ctl.Width = 260: ctl.Height = 70
    ctl.MultiLine = True: ctl.ScrollBars = 2

    yy = yy + 80
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "提交申请"
    ctl.Left = 80: ctl.Top = yy: ctl.Width = 100: ctl.Height = 28
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 200: ctl.Top = yy: ctl.Width = 100: ctl.Height = 28

    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    txtDate.Text = Format(Date, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "迟到" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "早退" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "缺勤" & q & vbCrLf
    c = c & "    cboStatus.AddItem " & q & "请假" & q & vbCrLf
    c = c & "    cboStatus.ListIndex = 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(txtReason.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写申请理由！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    AddAttendApply gCurrentUser, txtDate.Text, cboStatus.Text, txtReason.Text" & vbCrLf
    c = c & "    MsgBox " & q & "申请已提交！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建考勤审批查看窗体 ----------

Private Sub CreateAttendApplyViewForm()
    Dim oldForm As String
    oldForm = GetConfigProp("AttendApplyViewFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "AttendApplyViewFormName", actualName

    vbc.Properties("Caption") = "考勤审批详情"
    vbc.Properties("Width") = 420
    vbc.Properties("Height") = 310
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 18

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lc1": ctl.Caption = "申请人："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18: ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblApplicant": ctl.Caption = ""
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 100: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lc2": ctl.Caption = "考勤日期："
    ctl.Left = 200: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18: ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblDate": ctl.Caption = ""
    ctl.Left = 275: ctl.Top = yy: ctl.Width = 100: ctl.Height = 18

    yy = yy + 26
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lc3": ctl.Caption = "考勤状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18: ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblOldSt": ctl.Caption = ""
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 100: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lc4": ctl.Caption = "审批状态："
    ctl.Left = 200: ctl.Top = yy: ctl.Width = 70: ctl.Height = 18: ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lblApprSt": ctl.Caption = ""
    ctl.Left = 275: ctl.Top = yy: ctl.Width = 100: ctl.Height = 20
    ctl.Font.Bold = True

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lc5": ctl.Caption = "申请理由："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18: ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtReasonV"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 300: ctl.Height = 50
    ctl.MultiLine = True: ctl.Locked = True: ctl.ScrollBars = 2

    yy = yy + 58
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lc6": ctl.Caption = "审批回复："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18: ctl.Font.Bold = True
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtReply"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 300: ctl.Height = 50
    ctl.MultiLine = True: ctl.ScrollBars = 2

    yy = yy + 60
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdApprove": ctl.Caption = "同意"
    ctl.Left = 30: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28
    ctl.BackColor = RGB(60, 179, 113)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdReject": ctl.Caption = "拒绝"
    ctl.Left = 140: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28
    ctl.BackColor = RGB(220, 80, 80)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdClose": ctl.Caption = "关闭"
    ctl.Left = 280: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28

    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    If gViewAttendApplyID = " & q & q & " Then Unload Me: Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet" & vbCrLf
    c = c & "    Set ws = ThisWorkbook.Sheets(" & q & "考勤审批" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        If CStr(ws.Cells(i, 1).Value) = gViewAttendApplyID Then" & vbCrLf
    c = c & "            lblApplicant.Caption = ws.Cells(i, 2).Value" & vbCrLf
    c = c & "            lblDate.Caption = ws.Cells(i, 3).Value" & vbCrLf
    c = c & "            lblOldSt.Caption = ws.Cells(i, 4).Value" & vbCrLf
    c = c & "            txtReasonV.Text = ws.Cells(i, 5).Value" & vbCrLf
    c = c & "            lblApprSt.Caption = ws.Cells(i, 6).Value" & vbCrLf
    c = c & "            txtReply.Text = ws.Cells(i, 7).Value" & vbCrLf
    c = c & "            If ws.Cells(i, 6).Value <> " & q & "待审批" & q & " Then" & vbCrLf
    c = c & "                cmdApprove.Enabled = False" & vbCrLf
    c = c & "                cmdReject.Enabled = False" & vbCrLf
    c = c & "                txtReply.Locked = True" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            If Not IsHighPrivilege(gCurrentRole) Then" & vbCrLf
    c = c & "                cmdApprove.Enabled = False" & vbCrLf
    c = c & "                cmdReject.Enabled = False" & vbCrLf
    c = c & "                txtReply.Locked = True" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "            Exit For" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdApprove_Click()" & vbCrLf
    c = c & "    ApproveAttendApply gViewAttendApplyID, " & q & "已同意" & q & ", txtReply.Text" & vbCrLf
    c = c & "    MsgBox " & q & "已同意该申请，考勤记录已修改！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdReject_Click()" & vbCrLf
    c = c & "    ApproveAttendApply gViewAttendApplyID, " & q & "已拒绝" & q & ", txtReply.Text" & vbCrLf
    c = c & "    MsgBox " & q & "已拒绝该申请！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdClose_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建排班编辑窗体 ----------

Private Sub CreateScheduleEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("ScheduleEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "ScheduleEditFormName", actualName

    vbc.Properties("Caption") = "排班编辑"
    vbc.Properties("Width") = 380
    vbc.Properties("Height") = 260
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 18

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl1": ctl.Caption = "姓名："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboName"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 150: ctl.Height = 20
    ctl.Style = 0

    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl2": ctl.Caption = "排班时间："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtTime"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 250: ctl.Height = 20

    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl3": ctl.Caption = "工作地点："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtPlace"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 250: ctl.Height = 20

    yy = yy + 30
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl4": ctl.Caption = "备注："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtRemark"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 250: ctl.Height = 50
    ctl.MultiLine = True: ctl.ScrollBars = 2

    yy = yy + 60
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 80: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 190: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28

    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Dim wsU As Worksheet" & vbCrLf
    c = c & "    Set wsU = ThisWorkbook.Sheets(" & q & "用户管理" & q & ")" & vbCrLf
    c = c & "    Dim lr As Long, i As Long" & vbCrLf
    c = c & "    lr = wsU.Cells(wsU.Rows.Count, 2).End(xlUp).Row" & vbCrLf
    c = c & "    For i = 2 To lr" & vbCrLf
    c = c & "        cboName.AddItem wsU.Cells(i, 2).Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    txtTime.Text = Format(Date, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    If gEditScheduleRow > 0 Then" & vbCrLf
    c = c & "        Me.Caption = " & q & "编辑排班" & q & vbCrLf
    c = c & "        Dim ws As Worksheet" & vbCrLf
    c = c & "        Set ws = ThisWorkbook.Sheets(" & q & "排班管理" & q & ")" & vbCrLf
    c = c & "        cboName.Text = ws.Cells(gEditScheduleRow, 1).Value" & vbCrLf
    c = c & "        txtTime.Text = ws.Cells(gEditScheduleRow, 2).Value" & vbCrLf
    c = c & "        txtPlace.Text = ws.Cells(gEditScheduleRow, 3).Value" & vbCrLf
    c = c & "        txtRemark.Text = ws.Cells(gEditScheduleRow, 4).Value" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Me.Caption = " & q & "新建排班" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(cboName.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请选择姓名！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    SaveScheduleRecord gEditScheduleRow, cboName.Text, txtTime.Text, txtPlace.Text, txtRemark.Text" & vbCrLf
    c = c & "    MsgBox " & q & "排班已保存！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 创建人事管理编辑窗体 ----------

Private Sub CreatePersonnelEditForm()
    Dim oldForm As String
    oldForm = GetConfigProp("PersonnelEditFormName")
    If oldForm <> "" And FormExists(oldForm) Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(oldForm)
    End If

    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    Dim actualName As String
    actualName = vbc.Name
    SetConfigProp "PersonnelEditFormName", actualName

    vbc.Properties("Caption") = "人事档案编辑"
    vbc.Properties("Width") = 380
    vbc.Properties("Height") = 340
    vbc.Properties("StartUpPosition") = 2

    Dim dsg As Object
    Set dsg = vbc.Designer
    Dim ctl As Object
    Dim yy As Long
    yy = 18

    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl1": ctl.Caption = "姓名："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtName"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 120: ctl.Height = 20
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl2": ctl.Caption = "性别："
    ctl.Left = 220: ctl.Top = yy: ctl.Width = 40: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboGender"
    ctl.Left = 265: ctl.Top = yy: ctl.Width = 80: ctl.Height = 20
    ctl.Style = 2

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl3": ctl.Caption = "年龄："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtAge"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 60: ctl.Height = 20
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl4": ctl.Caption = "部门："
    ctl.Left = 160: ctl.Top = yy: ctl.Width = 40: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboDept"
    ctl.Left = 205: ctl.Top = yy: ctl.Width = 140: ctl.Height = 20
    ctl.Style = 0

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl5": ctl.Caption = "职务："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboPosition"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 140: ctl.Height = 20
    ctl.Style = 0

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl6": ctl.Caption = "入职时间："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtJoinDate"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 120: ctl.Height = 20

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl7": ctl.Caption = "在职状态："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.ComboBox.1")
    ctl.Name = "cboJobStatus"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 120: ctl.Height = 20
    ctl.Style = 2

    yy = yy + 28
    Set ctl = dsg.Controls.Add("Forms.Label.1")
    ctl.Name = "lbl8": ctl.Caption = "备注："
    ctl.Left = 20: ctl.Top = yy: ctl.Width = 60: ctl.Height = 18
    Set ctl = dsg.Controls.Add("Forms.TextBox.1")
    ctl.Name = "txtRemark"
    ctl.Left = 90: ctl.Top = yy: ctl.Width = 250: ctl.Height = 50
    ctl.MultiLine = True: ctl.ScrollBars = 2

    yy = yy + 58
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdSave": ctl.Caption = "保存"
    ctl.Left = 80: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28
    ctl.BackColor = RGB(60, 120, 216)
    Set ctl = dsg.Controls.Add("Forms.CommandButton.1")
    ctl.Name = "cmdCancel": ctl.Caption = "取消"
    ctl.Left = 190: ctl.Top = yy: ctl.Width = 90: ctl.Height = 28

    Dim cm As Object
    Set cm = vbc.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim q As String
    q = Chr(34)
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    cboGender.AddItem " & q & "男" & q & vbCrLf
    c = c & "    cboGender.AddItem " & q & "女" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "物业管理部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "工程维修部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "保洁绿化部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "安保部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "客服部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "财务部" & q & vbCrLf
    c = c & "    cboDept.AddItem " & q & "行政部" & q & vbCrLf
    c = c & "    cboPosition.AddItem " & q & "经理" & q & vbCrLf
    c = c & "    cboPosition.AddItem " & q & "主管" & q & vbCrLf
    c = c & "    cboPosition.AddItem " & q & "主任" & q & vbCrLf
    c = c & "    cboPosition.AddItem " & q & "员工" & q & vbCrLf
    c = c & "    cboPosition.AddItem " & q & "实习生" & q & vbCrLf
    c = c & "    cboJobStatus.AddItem " & q & "在职" & q & vbCrLf
    c = c & "    cboJobStatus.AddItem " & q & "离职" & q & vbCrLf
    c = c & "    cboJobStatus.AddItem " & q & "试用期" & q & vbCrLf
    c = c & "    cboJobStatus.AddItem " & q & "休假" & q & vbCrLf
    c = c & "    txtJoinDate.Text = Format(Date, " & q & "yyyy-mm-dd" & q & ")" & vbCrLf
    c = c & "    cboJobStatus.ListIndex = 0" & vbCrLf
    c = c & "    If gEditPersonnelRow > 0 Then" & vbCrLf
    c = c & "        Me.Caption = " & q & "编辑人事档案" & q & vbCrLf
    c = c & "        Dim ws As Worksheet" & vbCrLf
    c = c & "        Set ws = ThisWorkbook.Sheets(" & q & "人事管理" & q & ")" & vbCrLf
    c = c & "        txtName.Text = ws.Cells(gEditPersonnelRow, 1).Value" & vbCrLf
    c = c & "        cboGender.Text = ws.Cells(gEditPersonnelRow, 2).Value" & vbCrLf
    c = c & "        txtAge.Text = ws.Cells(gEditPersonnelRow, 3).Value" & vbCrLf
    c = c & "        cboDept.Text = ws.Cells(gEditPersonnelRow, 4).Value" & vbCrLf
    c = c & "        cboPosition.Text = ws.Cells(gEditPersonnelRow, 5).Value" & vbCrLf
    c = c & "        txtJoinDate.Text = ws.Cells(gEditPersonnelRow, 6).Value" & vbCrLf
    c = c & "        cboJobStatus.Text = ws.Cells(gEditPersonnelRow, 7).Value" & vbCrLf
    c = c & "        txtRemark.Text = ws.Cells(gEditPersonnelRow, 8).Value" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        Me.Caption = " & q & "新建人事档案" & q & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdSave_Click()" & vbCrLf
    c = c & "    If Trim(txtName.Text) = " & q & q & " Then" & vbCrLf
    c = c & "        MsgBox " & q & "请填写姓名！" & q & ", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    SavePersonnelRecord gEditPersonnelRow, txtName.Text, cboGender.Text, txtAge.Text, cboDept.Text, cboPosition.Text, txtJoinDate.Text, cboJobStatus.Text, txtRemark.Text" & vbCrLf
    c = c & "    MsgBox " & q & "人事档案已保存！" & q & ", vbInformation" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub cmdCancel_Click()" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub"

    cm.InsertLines 1, c
End Sub

' ---------- 人力资源窗体显示 ----------

Public Sub ShowHRMainForm()
    Dim fName As String
    fName = GetConfigProp("HRMainFormName")
    If fName = "" Or Not FormExists("HRMainFormName") Then
        MsgBox "人力资源窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("HRMainFormName")
    frm.Show
End Sub

Public Sub ShowAttendApplyEditForm()
    Dim fName As String
    fName = GetConfigProp("AttendApplyEditFormName")
    If fName = "" Or Not FormExists("AttendApplyEditFormName") Then
        MsgBox "考勤申请窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("AttendApplyEditFormName")
    frm.Show
End Sub

Public Sub ShowAttendApplyViewForm()
    Dim fName As String
    fName = GetConfigProp("AttendApplyViewFormName")
    If fName = "" Or Not FormExists("AttendApplyViewFormName") Then
        MsgBox "考勤审批窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("AttendApplyViewFormName")
    frm.Show
End Sub

Public Sub ShowScheduleEditForm()
    Dim fName As String
    fName = GetConfigProp("ScheduleEditFormName")
    If fName = "" Or Not FormExists("ScheduleEditFormName") Then
        MsgBox "排班编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ScheduleEditFormName")
    frm.Show
End Sub

Public Sub ShowPersonnelEditForm()
    Dim fName As String
    fName = GetConfigProp("PersonnelEditFormName")
    If fName = "" Or Not FormExists("PersonnelEditFormName") Then
        MsgBox "人事编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("PersonnelEditFormName")
    frm.Show
End Sub

' ---------- 人力资源数据操作函数 ----------

Private Function GenerateAttendApplyID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("考勤审批")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "KS" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateAttendApplyID = "KS" & Format(maxID + 1, "0000")
End Function

Public Sub AddAttendApply(sUser As String, sDate As String, sStatus As String, sReason As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("考勤审批")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateAttendApplyID()
    ws.Cells(newRow, 2).Value = sUser
    ws.Cells(newRow, 3).Value = sDate
    ws.Cells(newRow, 4).Value = sStatus
    ws.Cells(newRow, 5).Value = sReason
    ws.Cells(newRow, 6).Value = "待审批"
    ws.Cells(newRow, 7).Value = ""
    ws.Cells(newRow, 8).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Columns("A:H").AutoFit
End Sub

Public Sub ApproveAttendApply(aID As String, result As String, reply As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("考勤审批")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = aID Then
            ws.Cells(i, 6).Value = result
            ws.Cells(i, 7).Value = reply
            ' 如果同意，修改考勤管理表
            If result = "已同意" Then
                Dim applyUser As String
                applyUser = ws.Cells(i, 2).Value
                Dim applyDate As String
                applyDate = ws.Cells(i, 3).Value
                Dim wsA As Worksheet
                Set wsA = ThisWorkbook.Sheets("考勤管理")
                Dim lr2 As Long, j As Long
                lr2 = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
                For j = 2 To lr2
                    Dim recDate As String
                    If IsDate(wsA.Cells(j, 3).Value) Then
                        recDate = Format(wsA.Cells(j, 3).Value, "yyyy-mm-dd")
                    Else
                        recDate = CStr(wsA.Cells(j, 3).Value)
                    End If
                    If CStr(wsA.Cells(j, 2).Value) = applyUser And recDate = applyDate Then
                        wsA.Cells(j, 6).Value = "出勤"
                        wsA.Cells(j, 7).Value = "考勤审批修正"
                        Exit For
                    End If
                Next j
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub DeleteAttendApplyByID(aID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("考勤审批")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = aID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

Public Function FindScheduleRow(sName As String, sTime As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("排班管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    FindScheduleRow = 0
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = sName And CStr(ws.Cells(i, 2).Value) = sTime Then
            FindScheduleRow = i
            Exit Function
        End If
    Next i
End Function

Public Sub SaveScheduleRecord(rowNum As Long, sName As String, sTime As String, sPlace As String, sRemark As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("排班管理")
    Dim r As Long
    If rowNum > 0 Then
        r = rowNum
    Else
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If
    ws.Cells(r, 1).Value = sName
    ws.Cells(r, 2).Value = sTime
    ws.Cells(r, 3).Value = sPlace
    ws.Cells(r, 4).Value = sRemark
    If rowNum = 0 Then
        ws.Cells(r, 5).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    End If
    ws.Columns("A:E").AutoFit
End Sub

Public Sub SavePersonnelRecord(rowNum As Long, sName As String, sGender As String, sAge As String, sDept As String, sPos As String, sJoin As String, sJobSt As String, sRemark As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("人事管理")
    Dim r As Long
    If rowNum > 0 Then
        r = rowNum
    Else
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If
    ws.Cells(r, 1).Value = sName
    ws.Cells(r, 2).Value = sGender
    ws.Cells(r, 3).Value = sAge
    ws.Cells(r, 4).Value = sDept
    ws.Cells(r, 5).Value = sPos
    ws.Cells(r, 6).Value = sJoin
    ws.Cells(r, 7).Value = sJobSt
    ws.Cells(r, 8).Value = sRemark
    ws.Columns("A:H").AutoFit
End Sub

' ---------- 认证函数 ----------

Public Function AuthenticateUser(uName As String, uPass As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    AuthenticateUser = False
    If Not SheetExists("用户管理") Then Exit Function
    Set ws = ThisWorkbook.Sheets("用户管理")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If LCase(ws.Cells(i, 2).Value) = LCase(uName) Then
            If ws.Cells(i, 3).Value = uPass Then
                AuthenticateUser = True
                gCurrentUser = ws.Cells(i, 2).Value
                gCurrentRole = ws.Cells(i, 4).Value
                Exit Function
            End If
        End If
    Next i
End Function

Public Function GetUserRole(uName As String) As String
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    GetUserRole = ""
    If Not SheetExists("用户管理") Then Exit Function
    Set ws = ThisWorkbook.Sheets("用户管理")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        If LCase(ws.Cells(i, 2).Value) = LCase(uName) Then
            GetUserRole = ws.Cells(i, 4).Value
            Exit Function
        End If
    Next i
End Function

Public Function IsHighPrivilege(role As String) As Boolean
    IsHighPrivilege = (role = "管理员" Or role = "部门主管")
End Function

Public Function AdminExists() As Boolean
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    AdminExists = False
    If Not SheetExists("用户管理") Then Exit Function
    Set ws = ThisWorkbook.Sheets("用户管理")
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    For i = 2 To lastRow
        If ws.Cells(i, 4).Value = "管理员" Then
            AdminExists = True
            Exit Function
        End If
    Next i
End Function

Public Function AccountExists(acct As String) As Boolean
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    AccountExists = False
    If Not SheetExists("用户管理") Then Exit Function
    Set ws = ThisWorkbook.Sheets("用户管理")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        If LCase(ws.Cells(i, 2).Value) = LCase(acct) Then
            AccountExists = True
            Exit Function
        End If
    Next i
End Function

Public Sub RegisterNewUser(uName As String, uAcct As String, uPwd As String, uRole As String)
    Dim ws As Worksheet
    Dim newRow As Long
    Set ws = ThisWorkbook.Sheets("用户管理")
    newRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = uName
    ws.Cells(newRow, 2).Value = uAcct
    ws.Cells(newRow, 3).Value = uPwd
    ws.Cells(newRow, 4).Value = uRole
    ws.Cells(newRow, 5).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Columns("A:E").AutoFit
End Sub

' ---------- 启动系统 ----------

Public Sub LaunchSystem()
    Dim fName As String
    fName = GetConfigProp("LoginFormName")
    If fName = "" Then
        MsgBox "系统未安装，请先运行 SetupOASystem！", vbExclamation
        Exit Sub
    End If
    If Not FormExists("LoginFormName") Then
        MsgBox "登录窗体不存在，请重新运行 SetupOASystem！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("LoginFormName")
    frm.Show
End Sub

Public Sub ShowRegForm()
    Dim fName As String
    fName = GetConfigProp("RegisterFormName")
    If fName = "" Or Not FormExists("RegisterFormName") Then
        MsgBox "注册窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("RegisterFormName")
    frm.Show
End Sub

Public Sub ShowDashForm()
    Dim fName As String
    fName = GetConfigProp("DashboardFormName")
    If fName = "" Or Not FormExists("DashboardFormName") Then
        MsgBox "仪表盘窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("DashboardFormName")
    frm.Show
End Sub

Public Sub ShowWorkManagerForm()
    Dim fName As String
    fName = GetConfigProp("WorkManagerFormName")
    If fName = "" Or Not FormExists("WorkManagerFormName") Then
        MsgBox "工作管理窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("WorkManagerFormName")
    frm.Show
End Sub

Public Sub ShowWorkEditForm()
    Dim fName As String
    fName = GetConfigProp("WorkEditFormName")
    If fName = "" Or Not FormExists("WorkEditFormName") Then
        MsgBox "工作编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("WorkEditFormName")
    frm.Show
End Sub

Public Sub ShowWorkViewForm()
    Dim fName As String
    fName = GetConfigProp("WorkViewFormName")
    If fName = "" Or Not FormExists("WorkViewFormName") Then
        MsgBox "工作查看窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("WorkViewFormName")
    frm.Show
End Sub

' ---------- 工作管理数据操作函数 ----------

Private Function GenerateWorkID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("工作管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "WK" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateWorkID = "WK" & Format(maxID + 1, "0000")
End Function

Public Sub AddWorkRecord(sTitle As String, sDesc As String, sType As String, sPri As String, sAssignee As String, sDue As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("工作管理")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateWorkID()
    ws.Cells(newRow, 2).Value = sTitle
    ws.Cells(newRow, 3).Value = sDesc
    ws.Cells(newRow, 4).Value = sType
    ws.Cells(newRow, 5).Value = sPri
    ws.Cells(newRow, 6).Value = sAssignee
    ws.Cells(newRow, 7).Value = gCurrentUser
    ws.Cells(newRow, 8).Value = "待办"
    ws.Cells(newRow, 9).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(newRow, 10).Value = sDue
    ws.Cells(newRow, 11).Value = ""
    ws.Columns("A:K").AutoFit
End Sub

Public Sub UpdateWorkRecord(wID As String, sTitle As String, sDesc As String, sType As String, sPri As String, sAssignee As String, sDue As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("工作管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = wID Then
            ws.Cells(i, 2).Value = sTitle
            ws.Cells(i, 3).Value = sDesc
            ws.Cells(i, 4).Value = sType
            ws.Cells(i, 5).Value = sPri
            ws.Cells(i, 6).Value = sAssignee
            ws.Cells(i, 10).Value = sDue
            Exit For
        End If
    Next i
    ws.Columns("A:K").AutoFit
End Sub

Public Sub UpdateWorkStatus(wID As String, newStatus As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("工作管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = wID Then
            ws.Cells(i, 8).Value = newStatus
            If newStatus = "已完成" Then
                ws.Cells(i, 11).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub DeleteWorkByID(wID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("工作管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = wID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

' ---------- 信息管理窗体显示 ----------

Public Sub ShowNoticesForm()
    Dim fName As String
    fName = GetConfigProp("NoticesFormName")
    If fName = "" Or Not FormExists("NoticesFormName") Then
        MsgBox "公告列表窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("NoticesFormName")
    frm.Show
End Sub

Public Sub ShowNoticeEditForm()
    Dim fName As String
    fName = GetConfigProp("NoticeEditFormName")
    If fName = "" Or Not FormExists("NoticeEditFormName") Then
        MsgBox "公告编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("NoticeEditFormName")
    frm.Show
End Sub

' ---------- 信息管理数据操作函数 ----------

Private Function GenerateNoticeID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("信息管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "NT" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateNoticeID = "NT" & Format(maxID + 1, "0000")
End Function

Public Sub AddNoticeRecord(sTitle As String, sContent As String, sType As String, sDept As String, sTarget As String, stopp As String, sUrg As String, sStart As String, sEnd2 As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("信息管理")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateNoticeID()
    ws.Cells(newRow, 2).Value = sTitle
    ws.Cells(newRow, 3).Value = sContent
    ws.Cells(newRow, 4).Value = sType
    ws.Cells(newRow, 5).Value = sDept
    ws.Cells(newRow, 6).Value = gCurrentUser
    ws.Cells(newRow, 7).Value = sTarget
    ws.Cells(newRow, 8).Value = stopp
    ws.Cells(newRow, 9).Value = sUrg
    ws.Cells(newRow, 10).Value = sStart
    ws.Cells(newRow, 11).Value = sEnd2
    ws.Cells(newRow, 12).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Columns("A:L").AutoFit
End Sub

Public Sub UpdateNoticeRecord(nID As String, sTitle As String, sContent As String, sType As String, sDept As String, sTarget As String, stopp As String, sUrg As String, sStart As String, sEnd2 As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("信息管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = nID Then
            ws.Cells(i, 2).Value = sTitle
            ws.Cells(i, 3).Value = sContent
            ws.Cells(i, 4).Value = sType
            ws.Cells(i, 5).Value = sDept
            ws.Cells(i, 7).Value = sTarget
            ws.Cells(i, 8).Value = stopp
            ws.Cells(i, 9).Value = sUrg
            ws.Cells(i, 10).Value = sStart
            ws.Cells(i, 11).Value = sEnd2
            Exit For
        End If
    Next i
    ws.Columns("A:L").AutoFit
End Sub

Public Sub DeleteNoticeByID(nID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("信息管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = nID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

' ---------- 审批管理窗体显示 ----------

Public Sub ShowApprovalsForm()
    Dim fName As String
    fName = GetConfigProp("ApprovalsFormName")
    If fName = "" Or Not FormExists("ApprovalsFormName") Then
        MsgBox "审批中心窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ApprovalsFormName")
    frm.Show
End Sub

Public Sub ShowApprovalEditForm()
    Dim fName As String
    fName = GetConfigProp("ApprovalEditFormName")
    If fName = "" Or Not FormExists("ApprovalEditFormName") Then
        MsgBox "审批编辑窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ApprovalEditFormName")
    frm.Show
End Sub

Public Sub ShowApprovalViewForm()
    Dim fName As String
    fName = GetConfigProp("ApprovalViewFormName")
    If fName = "" Or Not FormExists("ApprovalViewFormName") Then
        MsgBox "审批查看窗体不存在！", vbExclamation
        Exit Sub
    End If
    Dim frm As Object
    Set frm = VBA.UserForms.Add("ApprovalViewFormName")
    frm.Show
End Sub

' ---------- 审批管理数据操作函数 ----------

Private Function GenerateApprovalID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("审批管理")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxID As Long
    maxID = 0
    Dim i As Long
    For i = 2 To lr
        Dim idStr As String
        idStr = CStr(ws.Cells(i, 1).Value)
        If Left(idStr, 2) = "SP" Then
            Dim numPart As Long
            numPart = Val(Mid(idStr, 3))
            If numPart > maxID Then maxID = numPart
        End If
    Next i
    GenerateApprovalID = "SP" & Format(maxID + 1, "0000")
End Function

Public Function GetDefaultApprover() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("用户管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    GetDefaultApprover = ""
    For i = 2 To lr
        Dim r As String
        r = ws.Cells(i, 4).Value
        If r = "管理员" Or r = "部门主管" Then
            GetDefaultApprover = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
End Function

Public Sub AddApprovalRecord(sType As String, sSub As String, sApplicant As String, sApprover As String, sReason As String, sStart As String, sEnd2 As String, sDays As String, sAmt As String, sItem As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("审批管理")
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = GenerateApprovalID()
    ws.Cells(newRow, 2).Value = sType
    ws.Cells(newRow, 3).Value = sSub
    ws.Cells(newRow, 4).Value = sApplicant
    ws.Cells(newRow, 5).Value = sApprover
    ws.Cells(newRow, 6).Value = sReason
    ws.Cells(newRow, 7).Value = sStart
    ws.Cells(newRow, 8).Value = sEnd2
    ws.Cells(newRow, 9).Value = sDays
    ws.Cells(newRow, 10).Value = sAmt
    ws.Cells(newRow, 11).Value = sItem
    ws.Cells(newRow, 12).Value = "待审批"
    ws.Cells(newRow, 13).Value = ""
    ws.Cells(newRow, 14).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Cells(newRow, 15).Value = ""
    ws.Columns("A:O").AutoFit
End Sub

Public Sub UpdateApprovalRecord(aID As String, sType As String, sSub As String, sReason As String, sStart As String, sEnd2 As String, sDays As String, sAmt As String, sItem As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("审批管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = aID Then
            ws.Cells(i, 2).Value = sType
            ws.Cells(i, 3).Value = sSub
            ws.Cells(i, 6).Value = sReason
            ws.Cells(i, 7).Value = sStart
            ws.Cells(i, 8).Value = sEnd2
            ws.Cells(i, 9).Value = sDays
            ws.Cells(i, 10).Value = sAmt
            ws.Cells(i, 11).Value = sItem
            Exit For
        End If
    Next i
    ws.Columns("A:O").AutoFit
End Sub

Public Sub UpdateApprovalStatus(aID As String, newStatus As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("审批管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = aID Then
            ws.Cells(i, 12).Value = newStatus
            ws.Cells(i, 15).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
            Exit For
        End If
    Next i
End Sub

Public Sub DeleteApprovalByID(aID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("审批管理")
    Dim lr As Long, i As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If CStr(ws.Cells(i, 1).Value) = aID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

Sub main()
    LoginFormName.Show
End Sub


