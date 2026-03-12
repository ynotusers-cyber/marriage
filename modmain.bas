Attribute VB_Name = "modmain"
Public con As New ADODB.Connection
Public pricecon As New ADODB.Connection
Public LoginSQL  As Integer
Public DefaultPrinter As Integer
Public InvPrinterName As String
Public InvPrinterPort As String
Public PrintTimes As Integer
Public WHATSAPPMSGSEND1 As String
Public WHATSAPPMSGvenue1 As String
Public gCstrDatabase As String
Public pwd1 As String
Public bkuppath As String

Public Sub Main()
    
    setconnection
    frmlogin.Show
End Sub
Public Sub showListView(lvwlist As ListView, sSQL As String, intColoumns As Integer)
    Dim litNode As ListItem
    Dim recRs As Recordset
    Dim ocmdCommand As Command
    Set recRs = CreateObject("adodb.recordset")
    Set ocmdCommand = CreateObject("adodb.Command")
    ocmdCommand.CommandText = sSQL
    ocmdCommand.ActiveConnection = con
        Set recRs = ocmdCommand.Execute
    Dim i As Integer
    lvwlist.ListItems.clear
    While Not recRs.EOF
        Set litNode = lvwlist.ListItems.Add(, , CStr(recRs.Fields(0).Value))
        For i = 1 To intColoumns - 1
            If Not IsNull(recRs.Fields(i).Value) Then
                If recRs.Fields(i).Type = adCurrency Then
                    litNode.SubItems(i) = Format(CStr(recRs.Fields(i).Value), "0.00")

                Else
                    If recRs.Fields(i).Value = "120:00:00 AM" Then
                        litNode.SubItems(i) = ""
                    Else
                        litNode.SubItems(i) = CStr(recRs.Fields(i).Value)
                    End If
                End If
            End If
        Next i
        recRs.MoveNext
    Wend
    Set ocmdCommand = Nothing
    Set recRs = Nothing
    End Sub

Public Function SetBackColor(frm As Form)
On Error GoTo ErrHnd
    Dim ctrl As Control
    Dim mintred As Long
    Dim mintgreen As Long
    Dim mintblue As Long
    
    mintred = 188
    mintgreen = 208
    mintblue = 190
'    mintred = 147
'    mintgreen = 155
'    mintblue = 193

    frm.BackColor = RGB(mintred, mintgreen, mintblue)
 '   frm.Icon = LoadResPicture(101, vbResIcon)
    
    For Each ctrl In frm
'        Debug.Print ctrl.Name
        If TypeOf ctrl Is Label Or TypeOf ctrl Is Frame Or TypeOf _
               ctrl Is OptionButton Or TypeOf ctrl Is CheckBox Or TypeOf _
               ctrl Is PictureBox Or TypeOf ctrl Is SSTab _
               Or TypeOf ctrl Is CommandButton Then
            ctrl.BackColor = RGB(mintred, mintgreen, mintblue)
        ElseIf TypeOf ctrl Is MSHFlexGrid Then
            ctrl.BackColorBkg = RGB(mintred, mintgreen, mintblue)
            ctrl.BackColorFixed = RGB(mintred, mintgreen, mintblue)
        End If
        If TypeOf ctrl Is Label Then
            ctrl.Font = "Arial"
        End If
        If TypeOf ctrl Is CommandButton Then
            ctrl.Font = "calibri"
        End If
       If TypeOf ctrl Is DTPicker Then
         '   ctrl.Value = Date
       End If
    Next
    
Exit Function
ErrHnd:
    Err.Raise Err.Number, , Err.Description
End Function


Public Sub Centerform(frm As Form)
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
End Sub
Public Sub setconnection()
    Dim str As String
    Dim database As String
    Dim server As String
    Dim pwd As String
    Dim pricecheckerdata As String
    server = GetSetting("POS", "server", "ServerName", "")
    database = GetSetting("POS", "server", "DataBase", "")
    gCstrDatabase = database
   ' pricecheckerdata = GetSetting("POS", "Server", "Pricecheckerdatabase", "")
  '  SaveSetting "POS", "Server", "Pricecheckerdatabase", YNOT
      ' Save value

'  gvcmdenable = GetSetting("POS", "Settings", "gvcmdenable", 0)
'    SaveSetting "POS", "Settings", "gvcmdenable", gvcmdenable
' Read value
WHATSAPPMSGSEND = GetSetting("POS", "Settings", "WHATSAPPMSGSEND", "")
'SaveSetting "POS", "Settings", "WHATSAPPMSGSEND", ""
WHATSAPPMSGSEND1 = WHATSAPPMSGSEND

WHATSAPPVENUE = GetSetting("POS", "Settings", "WHATSAPPvenue", "")
WHATSAPPMSGvenue1 = WHATSAPPVENUE
'SaveSetting "POS", "Settings", "WHATSAPPvenue", ""

'MsgLine2 = GetSetting("POS", "Settings", "MsgLine2", "")
'SaveSetting "POS", "Settings", "MsgLine2", THANKS

    bkuppath = GetSetting("POS", "Settings", "bkuppath", "D:")
    SaveSetting "POS", "Settings", "bkuppath", bkuppath
 


    pwd = GetSetting("POS", "Settings", "pwd", "")
    pwd1 = pwd
    LoginSQL = GetSetting("POS", "Settings", "LoginSql", 1)
       
        If LoginSQL = 1 Then
            str = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & database & ";Data Source=" & server
        Else
            str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=" & pwd & ";Initial Catalog=" & database & ";Data Source=" & server
        End If
        con.Open str
'
'     If LoginSQL = 1 Then
'            str = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & database & ";Data Source=" & server
'        Else
'            str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=" & pwd & ";Initial Catalog=" & database & ";Data Source=" & server
'        End If
'        con.Open str
    InvPrinterName = GetSetting("POS", "Settings", "InvPrinterName", "")
    SaveSetting "POS", "Settings", "InvPrinterName", InvPrinterName
   ' Public PrintTimes As Integer
    InvPrinterPort = GetSetting("POS", "Settings", "InvPrinterPort", "")
    SaveSetting "POS", "Settings", "InvPrinterPort", InvPrinterPort
        If LoginSQL = 1 Then
            str = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & database & ";Data Source=" & server
        Else
            str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password=" & pwd & ";Initial Catalog=" & database & ";Data Source=" & server
        End If
        pricecon.Open str
     
End Sub

Public Function getbarcode(barcode As String)
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
    With cmd
        .ActiveConnection = con
        .CommandText = "alter proc items as select a.*,b.qty from itemsearchview a inner join pstocktable b on a.id=b.ItemMasterId where barcode='" & barcode & "'"
        .CommandType = adCmdText
        Set rs = .Execute
    End With
'    If Not rs.EOF Then
'    End If
End Function
Public Function CheckForFloat(KeyAscii As Integer) As Integer
    If KeyAscii = 46 Then
        Dim k As Long
        k = 0
        k = InStr(1, Screen.ActiveControl.Text, ".")
            If k <> 0 Then
                CheckForFloat = 0
            Else
                CheckForFloat = KeyAscii
            End If
        Exit Function
    End If
    If KeyAscii = 8 Then
        CheckForFloat = KeyAscii
    ElseIf KeyAscii <= 47 Or KeyAscii >= 58 Then
        CheckForFloat = 0
    Else
        CheckForFloat = KeyAscii
    End If
End Function
Public Function NzString(val As Variant) As String
    If IsNull(val) Then
        NzString = ""
    Else
        NzString = CStr(val)
    End If
End Function
