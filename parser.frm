VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17475
   LinkTopic       =   "Form1"
   ScaleHeight     =   12165
   ScaleWidth      =   17475
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   11415
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   20135
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "View "
      TabPicture(0)   =   "parser.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DBGrid2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DBGrid1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Data1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Data2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "List2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Convert / Search"
      TabPicture(1)   =   "parser.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "text1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "List1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "File1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "parser.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command8 
         Caption         =   "Learn"
         Height          =   540
         Left            =   11475
         TabIndex        =   20
         Top             =   4275
         Width           =   5040
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">"
         Height          =   390
         Left            =   16125
         TabIndex        =   19
         Top             =   1500
         Width           =   390
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<"
         Height          =   390
         Left            =   11475
         TabIndex        =   18
         Top             =   1500
         Width           =   390
      End
      Begin VB.TextBox Text4 
         Height          =   390
         Left            =   14100
         TabIndex        =   17
         Text            =   "Text4"
         Top             =   1500
         Width           =   1890
      End
      Begin VB.TextBox Text3 
         Height          =   390
         Left            =   12000
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   1500
         Width           =   1890
      End
      Begin RichTextLib.RichTextBox text1 
         Height          =   9765
         Left            =   -74175
         TabIndex        =   7
         Top             =   1500
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   17224
         _Version        =   393217
         TextRTF         =   $"parser.frx":0054
      End
      Begin VB.TextBox Text2 
         Height          =   1890
         HideSelection   =   0   'False
         Left            =   11400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2175
         Width           =   5130
      End
      Begin VB.ListBox List2 
         Height          =   4335
         ItemData        =   "parser.frx":00DF
         Left            =   11475
         List            =   "parser.frx":00E1
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   6300
         Width           =   4965
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6375
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   10500
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   450
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   10575
         Visible         =   0   'False
         Width           =   1440
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "parser.frx":00E3
         Height          =   10140
         Left            =   225
         OleObjectBlob   =   "parser.frx":00F7
         TabIndex        =   9
         Top             =   975
         Width           =   5490
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Converte CSV"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74100
         TabIndex        =   8
         Top             =   750
         Width           =   1740
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search Abstract"
         Height          =   315
         Left            =   -66075
         TabIndex        =   6
         Top             =   750
         Width           =   1740
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Get PDF/TXT"
         Height          =   315
         Left            =   -72075
         TabIndex        =   5
         Top             =   750
         Width           =   1740
      End
      Begin VB.ListBox List1 
         Height          =   9420
         Left            =   -66525
         TabIndex        =   4
         Top             =   1500
         Width           =   8040
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Verify TXT"
         Height          =   315
         Left            =   -70050
         TabIndex        =   3
         Top             =   750
         Width           =   1740
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   -69675
         Pattern         =   "*.PDF"
         TabIndex        =   2
         Top             =   1575
         Width           =   2265
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Search TXT"
         Height          =   315
         Left            =   -64125
         TabIndex        =   1
         Top             =   750
         Width           =   1740
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "parser.frx":0AA9
         Height          =   10140
         Left            =   6000
         OleObjectBlob   =   "parser.frx":0ABD
         TabIndex        =   10
         Top             =   975
         Width           =   4965
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13875
         TabIndex        =   22
         Top             =   5700
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13800
         TabIndex        =   21
         Top             =   4950
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Counting 0/0"
         Height          =   315
         Left            =   11550
         TabIndex        =   14
         Top             =   10800
         Width           =   4965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "9999"
         Height          =   195
         Left            =   -64125
         TabIndex        =   13
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13500
         TabIndex        =   12
         Top             =   975
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'--------------------------------------------------------------
'Declarações de varáveis globais
'--------------------------------------------------------------
Dim x_db As Database

Dim x_ont As Recordset
Dim x_up As Recordset
Dim x_upd As Recordset

Dim x_totcasos As Recordset
Dim x_selcasos As Recordset


Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
  ByVal szFileName As String, ByVal dwReserved As Long, _
  ByVal lpfnCB As Long) As Long


Dim x_file_content  As String
Dim x_nome As String

Dim x_rule As Integer
'--------------------------------------------------------------
'Busca e retorna o conteúdo de uma tag dentro de
'Assume que as tags possam ser diferentes para um mesmo marcador/grupo
'--------------------------------------------------------------
Function acha_tag(x_xml As String, x_tagi As String, x_tagf As String) As String

acha_tag = ""
x_poi = InStr(1, x_xml, x_tagi)
If x_poi > 0 Then
   x_pof = InStr(x_poi, x_xml, x_tagf)
Else
   Exit Function
End If



If x_pof > x_poi Then
   acha_tag = Mid(x_xml, x_poi + Len(x_tagi), x_pof - x_poi - Len(x_tagf) + 1)
End If


End Function


Function neigbors(x_txt, x_desc, x_wich) As String

If Len(x_txt) = 0 Then
   Exit Function
End If

neigbors = ""

Const onespace As String = " "
tmpWords = Split(x_txt, onespace)


x_pos = InStr(1, x_desc, onespace)
If x_pos > 0 Then
   If x_wich = 0 Then
      x_tmp = Trim(Mid(x_desc, 1, x_pos))
      neigbors = neigbors(x_txt, x_tmp, 0)
   ElseIf x_wich = 1 Then
      x_tmp = Trim(Mid(x_desc, x_pos, 255))
      neigbors = neigbors(x_txt, x_tmp, 1)
   End If

   Exit Function
End If


For i = 0 To UBound(tmpWords)

    If tmpWords(i) = Trim(x_desc) Then
    
       If x_wich = 0 Then
            If i > 0 Then
               neigbors = tmpWords(i - 1)
            End If
       ElseIf x_wich = 1 Then
            If i < UBound(tmpWords) Then
               neigbors = tmpWords(i + 1)
            End If
       End If
       Exit For

    End If
 
Next

End Function

Sub converte(x_file As String)
x_cont = 0
x_asd = Chr(34)
Dim x_tag As String
Dim x_lic As String
Dim x_data As String
   
   Open App.Path + "\XML\convertjson_" + x_file + ".xml" For Input As #1    ' Open file.
     
   x_tag = ""
   text1.Text = ""
    
   Do While Not EOF(1)
      Line Input #1, x_line ' le uma linha
      
       
       If Trim(x_line) = "<items>" Then
          Do While Not EOF(1)
             Line Input #1, x_line ' le uma linha
             x_tag = x_tag + x_line
             If Trim(x_line) = "</items>" Then
                Exit Do
             End If
          Loop
       End If
       '"<reference-count>"
       
       If x_tag <> "" Then
          x_lin = x_lin + x_asd + Me.acha_tag(x_tag, "<reference-count>", "</reference-count>") + x_asd + ";"
          x_lin = x_lin + x_asd + Me.acha_tag(x_tag, "<short-container-title>", "</short-container-title>") + x_asd + ";"
          
          
          
          x_data = Me.acha_tag(x_tag, "<published-print>", "</published-print>")
          x_data = Trim(Me.acha_tag(x_data, "<date-parts>", "</date-parts>"))
          x_dat = ""
          If x_data <> "" Then
                x_poi = InStr(1, x_data, "<key>") + 5
                x_pof = InStr(1, x_data, "</key>")
                x_dat = Mid(x_data, x_poi, x_pof - x_poi) + "/"
                
                x_poi = InStr(x_pof, x_data, "<key>") + 5
                x_pof = InStr(x_pof + 5, x_data, "</key>")
                x_dat = x_dat + Mid(x_data, x_poi, x_pof - x_poi) + "/"
                
                x_poi = InStr(x_pof, x_data, "<key>") + 5
                x_pof = InStr(x_pof + 5, x_data, "</key>")
                If x_pof > x_poi Then
                   x_dat = x_dat + Mid(x_data, x_poi, x_pof - x_poi)
                Else
                   x_dat = Mid(x_dat, 1, Len(x_dat) - 1)
                End If
          
          End If
          
          x_lin = x_lin + x_asd + Format(x_dat, "dd/mm/yyyy") + x_asd + ";"
          
          x_lin = x_lin + x_asd + Me.acha_tag(x_tag, "<title>", "</title>") + x_asd + ";"
         
          x_lic = Me.acha_tag(x_tag, "<license>", "</license>")
          If Trim(x_lic) <> "" Then
             x_lin = x_lin + "Sim;"
          Else
             x_lin = x_lin + "Não;"
          End If
          
          x_lic = Me.acha_tag(x_tag, "<link>", "</link>")
          x_lic = Me.acha_tag(x_lic, "<URL>", "</URL>")
          x_lin = x_lin + x_asd + x_lic + x_asd + ";"
          
         
          x_lin = x_lin + x_asd + Me.acha_tag(x_tag, "<language>", "</language>") + x_asd + ";"
          
          x_lin = x_lin + x_asd + monta_subject(x_tag, 1) + x_asd + ";"
          
          x_lin = x_lin + x_asd + Me.acha_tag(x_tag, "<type>", "</type>") + x_asd + ";"
          
          x_lin = x_lin + x_asd + Me.acha_tag(x_tag, "<abstract>", "</abstract>") + x_asd + ";"
         
          'text1.Text = text1.Text + x_lin + vbNewLine
          Print #2, x_lin
          x_lin = ""
          x_tag = ""
      
       End If
       
       x_cont = x_cont + 1
       Me.Caption = x_cont
   Loop
   Close #1


'text1.SaveFile App.Path + "\XML\convertjson_" + x_file + ".csv", rtfText

End Sub

Function monta_descritores() As String
x_ret = ""
If x_ont.Fields("desc1") <> "" Then x_ret = x_ont.Fields("desc1")
If x_ont.Fields("desc2") <> "" Then x_ret = x_ret + ", " + x_ont.Fields("desc2")
If x_ont.Fields("desc3") <> "" Then x_ret = x_ret + ", " + x_ont.Fields("desc3")
If x_ont.Fields("desc4") <> "" Then x_ret = x_ret + ", " + x_ont.Fields("desc4")
If x_ont.Fields("desc5") <> "" Then x_ret = x_ret + ", " + x_ont.Fields("desc5")

monta_descritores = x_ret

End Function

Function monta_ontologia() As String

x_ret = ""
If x_ont.Fields("desc1") <> "" Then x_ret = " Todos.abstract LIKE '*" + x_ont.Fields("desc1") + "*' "
If x_ont.Fields("desc2") <> "" Then x_ret = x_ret + " OR Todos.abstract LIKE '*" + x_ont.Fields("desc2") + "*' "
If x_ont.Fields("desc3") <> "" Then x_ret = x_ret + " OR Todos.abstract LIKE '*" + x_ont.Fields("desc3") + "*' "
If x_ont.Fields("desc4") <> "" Then x_ret = x_ret + " OR Todos.abstract LIKE '*" + x_ont.Fields("desc4") + "*' "
If x_ont.Fields("desc5") <> "" Then x_ret = x_ret + " OR Todos.abstract LIKE '*" + x_ont.Fields("desc5") + "*' "


monta_ontologia = x_ret

End Function

Function monta_subject(x_xml, x_pi) As String


monta_subject = ""
x_poi = InStr(x_pi, x_xml, "<subject>")
If x_poi > 0 Then
   x_pof = InStr(x_poi, x_xml, "</subject>")
   
   monta_subject = monta_subject + monta_subject(x_xml, x_pof)
   
Else
   Exit Function
End If




If x_pof > x_poi Then
   monta_subject = monta_subject + Mid(x_xml, x_poi + Len("<subject>"), x_pof - x_poi - Len("</subject>") + 1) + ", "
End If





End Function

Function percent_total() As String
    
    Dim x_per As Recordset
    Set x_per = x_db.OpenRecordset("temp_totais")
    x_temp = (x_per.Fields(0) * 100) / (x_per.Fields(0) + x_per.Fields(1))
    percent_total = "Global " + Format(x_temp, "##0") + "% Classified"
    x_per.Close

End Function

Function save_learning_table() As String


    save_learning_table = ""
    x_cont = 1
    x_save = ""
    x_file = "learning_table.txt"
    If Dir(App.Path + "\" + x_file) = "" Then
       Exit Function
    End If
    
    x_desc = Trim(Mid(Data2.Recordset.Fields(1), 4, 255))
    x_ndes = Trim(Mid(Data2.Recordset.Fields(1), 1, 2))
    x_aux = "[" + Data1.Recordset.Fields(0) + "]" + _
                   "[" + x_ndes + "]" + _
                   "[" + Text3.Text + "]" + _
                   "[" + x_desc + "]" + _
                   "[" + Text4.Text + "]" + _
                   ""

         
    Open App.Path + "\" + x_file For Input As #1
    
    Do While Not EOF(1)
       Line Input #1, x_line
       
       If InStr(1, x_line, x_aux) > 0 Then
          save_learning_table = "[E_" + Format(x_cont, "0000") + "]"
          Close #1
          Exit Function
       End If
       
       x_save = x_save + x_line + vbNewLine
       x_cont = x_cont + 1
    Loop
    
    save_learning_table = "[E_" + Format(x_cont, "0000") + "]"
    x_rule = x_cont
    
    x_save = x_save + "[E_" + Format(x_cont, "0000") + "]" + _
                        x_aux
                        
    Close #1
    
    Open App.Path + "\" + x_file For Output As #1
    Print #1, x_save
    Close #1

End Function

Function search_descriptor(x_descriptor As String, x_descord As String) As String

search_descriptor = ""
x_cont = 0
Dim i As Long
Dim x_ctrl As Integer
Dim x_par As String

If x_descriptor <> "" Then
   x_poi = InStr(1, LCase(x_file_content), " " + LCase(x_descriptor) + " ")
   While x_poi > 0
   
         x_par = ""
         If x_poi <= 255 Then
            x_par = Mid(x_file_content, x_poi, 510)
            x_i = 1
         Else
            x_par = Mid(x_file_content, x_poi - 255, 510)
            x_i = x_poi - 255
         End If
         
         x_curchar = Mid(x_file_content, x_poi, 1)
         x_par = ""
         i = x_i
         
         
         'calcs if the number of char to get will past the end of file. Its to avoid a dead loop.
         x_numchar = Len(x_file_content) - x_poi
         If x_numchar >= 512 Then
            x_numchar = 512
         Else
            x_numchar = x_numchar - 1
         End If
         
         ' this variable is for dead loop control
         x_ctrl = 0

         
         Do While Len(x_par) <= x_numchar And x_ctrl < 1024

             x_ctrl = x_ctrl + 1
             
             If x_curchar <> " " Then
                x_par = x_par + x_curchar
             Else
                If Mid(x_file_content, i, 1) <> " " Then
                   x_par = x_par + x_curchar
                End If
             End If
             x_curchar = Mid(x_file_content, i, 1)
             i = i + 1
             Label2.Caption = File1.List(File1.ListIndex) + Str(x_poi) + " - " + Str(Len(x_file_content)) + " - " + Str(Len(x_par))
             
            
         Loop
         
         Me.Refresh
         
         Print #3, x_nome + ";" + x_ont.Fields(0) + ";" + x_descord + " - " + x_descriptor + ";" + Str(x_poi) + ";'" + rmv_special(x_par) + "'"
   
         x_cont = x_cont + 1
         x_poi = InStr(x_poi + Len(x_descriptor), LCase(x_file_content), " " + LCase(x_descriptor) + " ")
   Wend



   search_descriptor = x_descriptor + ";" + Trim(Str(x_cont))
   
End If



End Function

Sub ToText(x_filename As String)
    
    
    
    ' Hanlde Error
    On Error GoTo ErrorHandler:
    
    ' Create Bytescout.PDFExtractor.TextExtractor object
    Set extractor = CreateObject("Bytescout.PDFExtractor.TextExtractor")
    
    ' Set Registration name and key
    extractor.RegistrationName = "demo"
    extractor.RegistrationKey = "demo"
  
    ' Load sample PDF document
    extractor.LoadDocumentFromFile x_filename + ".pdf"
    
    ' Peform Save to Text file
    extractor.SaveTextToFile x_filename + ".txt"
    
    ' Show Success Message
    'MsgBox "Extracted data saved to 'output.text' file.", vbInformation, "Success"
    
    ' Close form
    'Unload Me
    
ErrorHandler:
If Err.Number <> 0 Then
    'MsgBox Err.Description, vbInformation, "Error"
End If

End Sub

Function verifica_arquivo(x_path As String) As Boolean

    verifica_arquivo = False

    If Dir(x_path) = "" Then
       Exit Function
    End If
    
    Open x_path For Input As #1    ' Open file.
     
    Line Input #1, x_line ' le uma linha
    Close #1

    x_poi = -1
    x_inicio = Mid(x_line, 1, 20)
    
    x_poi = InStr(1, UCase(x_inicio), "PDF")
    
    If x_poi > 0 Then
       verifica_arquivo = True
    Else
       verifica_arquivo = False
       Kill x_path
    End If
  

End Function

Private Sub Command1_Click()

Open App.Path + "\todos.CSV" For Output As #2    ' Open file.

For i = 1 To 7
   converte Trim(Str(i))
Next

Close #2
MsgBox "ok"


End Sub




Private Sub Command2_Click()

          text1.Text = ""

          'Set x_db = OpenDatabase(App.Path + "\base_acm_crossref.accdb", False, False, "")
          Set x_ont = x_db.OpenRecordset("select * from ontologia order by categoria")
          Set x_upd = x_db.OpenRecordset("todos")
          x_upd.Index = "id"

          While Not x_ont.EOF
                x_aux = "SELECT * From Todos WHERE " + monta_ontologia() + ";"
                Set x_up = x_db.OpenRecordset(x_aux)
                x_cont = 0
                While Not x_up.EOF
                      x_cont = x_cont + 1
                      
                      'x_upd.Seek "=", x_up.Fields("id")
                      
                      x_up.Edit
                      x_up.Fields("classif") = IIf(IsNull(x_up.Fields("classif")) Or Trim(x_up.Fields("classif")) = "", Me.monta_descritores(), x_up.Fields("classif") + ", " + Me.monta_descritores())
                      x_up.Update
                      
                      
                      x_up.MoveNext
                Wend
                text1.Text = text1.Text + monta_descritores() + " " + Trim(Str(x_cont)) + vbNewLine
                x_up.Close
                x_ont.MoveNext
          Wend
          
          MsgBox ok
End Sub


Private Sub Command3_Click()
        Set extractor = CreateObject("Bytescout.PDFExtractor.TextExtractor")
        
        Dim x_file As String
        Dim x_status As String
        
        'Set x_db = OpenDatabase(App.Path + "\base_acm_crossref.accdb", False, False, "")
        Set x_upd = x_db.OpenRecordset("todos")
        x_upd.Index = "id"

        While Not x_upd.EOF
              x_url = x_upd.Fields("link")
                                        
              x_s = Format(x_upd.Fields("id"), "000000")
              x_file = App.Path + "\PDF\ret_" + x_s '+ ".pdf"
              x_status = ""
              
              
              If IsNull(x_upd.Fields("status_pdf")) Or x_upd.Fields("status_pdf") = "" Then
              
                      If Dir(x_file + ".pdf") = "" Then
                      
                            If Not IsNull(x_url) And x_url <> "" Then
                                  returnValue = URLDownloadToFile(0, x_url, x_file + ".pdf", 0, 0)
                                  
                                  x_pdf = verifica_arquivo(x_file + ".pdf")
                                  
                                  If returnValue <> 0 Then
                                     x_status = x_s + " " + x_upd.Fields("restrito") + " -> ERRO " + x_url
                                  Else
                                     If x_pdf Then
                                        ToText App.Path + "\PDF\ret_" + x_s
                                        x_status = x_s + " " + x_upd.Fields("restrito") + " -> ACESSO ABERTO"
                                     Else
                                        x_status = x_s + " " + x_upd.Fields("restrito") + " -> RESTRITO"
                                     End If
                                     Me.Refresh
                                     List1.ListIndex = List1.ListCount - 1
                                  End If
                            Else
                               
                               x_status = x_s + " " + x_upd.Fields("restrito") + " -> ERRO URL"
                            End If
                      Else
                          If Dir(x_file + ".txt") = "" Then
                             ToText App.Path + "\PDF\ret_" + x_s
                          End If
                          x_status = x_s + " " + x_upd.Fields("restrito") + " -> ACESSO ABERTO NO DISCO"
                      End If
              
                    x_upd.Edit
                    x_upd.Fields("status_pdf") = x_status
                    x_upd.Update
              
              
              Else
                  x_status = x_upd.Fields("status_pdf")
              End If
              
              List1.AddItem x_status
              
              
              
              
              x_upd.MoveNext
              Form1.Refresh
        Wend


MsgBox "terminou"

End Sub


Private Sub Command4_Click()
Dim x_file As String
File1.Path = App.Path + "\PDF"

For i = 0 To File1.ListCount - 1

    x_file = App.Path + "\PDF\" + Mid(File1.List(i), 1, 10)

    If Dir(x_file + ".txt") = "" Then
       ToText x_file
       List1.AddItem File1.List(i)
    End If


Next





End Sub


Private Sub Command5_Click()

          
Dim x_file As String
Dim x_ret As String

'Set x_db = OpenDatabase(App.Path + "\base_acm_crossref.accdb", False, False, "")
Set x_ont = x_db.OpenRecordset("select * from ontologia order by categoria")
    
Open App.Path + "\search_descriptor_sum.csv" For Output As #1   ' Open file.
Open App.Path + "\search_descriptor_cases.csv" For Output As #3   ' Open file.


File1.Path = App.Path + "\PDF"
File1.Pattern = "*.txt"
File1.Refresh


text1.Text = ""
x_ont.MoveFirst
x_ret = ""

For i = 0 To File1.ListCount - 1


    File1.ListIndex = i
    x_file = App.Path + "\PDF\" + Mid(File1.List(i), 1, 10) + ".txt"
    
    x_file_content = ""
    
    Open x_file For Input As #2      ' Open file.
    
    
    Do While Not EOF(2)
       Line Input #2, x ' le uma linha
       x_file_content = x_file_content + x
    Loop
    Close #2
    
    Me.Refresh
    
    x_nome = Trim(Str(Val(Mid(File1.List(i), 5, 6))))
    
    'for debug propouses
    If x_nome = "120" Then
       x = 0
    End If
    
    While Not x_ont.EOF
     
          If x_ont.Fields("desc1") = "x - controle" Then
             x = 0
          End If
        
          If Not IsNull(x_ont.Fields("desc1")) Then Print #1, x_nome + ";" + x_ont.Fields(0) + ";" + Me.search_descriptor(x_ont.Fields("desc1"), "1")
          
          If Not IsNull(x_ont.Fields("desc2")) Then Print #1, x_nome + ";" + x_ont.Fields(0) + ";" + Me.search_descriptor(x_ont.Fields("desc2"), "2")
          
          If Not IsNull(x_ont.Fields("desc3")) Then Print #1, x_nome + ";" + x_ont.Fields(0) + ";" + Me.search_descriptor(x_ont.Fields("desc3"), "3")
          
          If Not IsNull(x_ont.Fields("desc4")) Then Print #1, x_nome + ";" + x_ont.Fields(0) + ";" + Me.search_descriptor(x_ont.Fields("desc4"), "4")
          
          If Not IsNull(x_ont.Fields("desc5")) Then Print #1, x_nome + ";" + x_ont.Fields(0) + ";" + Me.search_descriptor(x_ont.Fields("desc5"), "5")
           
          x_ont.MoveNext
    
    Wend
  
    x_ont.MoveFirst
    Me.Caption = File1.List(i)
    Me.Refresh

Next
          
          
Close #1
Close #2
Close #3

text1.LoadFile App.Path + "\search_descriptor_cases.csv"

MsgBox ok

End Sub


Function rmv_special(x As String) As String
Dim t, i As Long
Dim s, Y, z As String
s = ""
t = Len(x)
For i = 1 To t
    
    If Mid(UCase(x), i, 6) = "*DEMO*" Then
       i = i + 6
    End If

    Y = UCase(Mid(x, i, 1))
    x_pos = InStr(1, UCase(" 01234567890abcdefghijklmnopqrstuvxywz().|&-:/\%"), Y)
    If x_pos = 0 Then
       Y = " "
    End If
    s = s + Y
Next
rmv_special = s
End Function


Private Sub Command6_Click()
    
    x_desc = Text3.Text
    Text3.Text = Me.neigbors(Text2.Text, x_desc, 0)

End Sub

Private Sub Command7_Click()
    x_desc = Text4.Text
    Text4.Text = Me.neigbors(Text2.Text, x_desc, 1)

End Sub


Private Sub Command8_Click()

x = Me.save_learning_table()


' analisa vizinhos anterior e posterior com k=0
x_pre = Trim(Text3.Text)
x_suc = Trim(Text4.Text)
x_update = x
x_desc = Trim(Mid(Data2.Recordset.Fields(1), 4, 255))

'AND (aceite is null or aceite='')
x_query = "Select artigo, descritor, fragmento, aceite from busca_descritores_casos " + _
"WHERE descritor = '" + Data2.Recordset.Fields(1) + "' AND (" + _
IIf(x_pre <> "", "fragmento Like '* " + x_pre + " " + x_desc + " *' ", "") + _
IIf(x_pre <> "" And x_suc <> "", " OR ", "") + _
IIf(x_suc <> "", "fragmento Like '* " + x_desc + " " + x_suc + " *'", "") + _
")"


Set x_selcasos = x_db.OpenRecordset(x_query)
'Set Data2.Recordset = x_selcasos
'Data2.Refresh
'DBGrid2.Columns(0).Width = 500




x_a = x_selcasos.RecordCount
x_reg = Data2.Recordset.Fields("artigo")
With x_selcasos
     Do While Not .EOF
           .Edit
           .Fields("aceite") = IIf(Not IsNull(.Fields("aceite")), .Fields("aceite"), "") + x_update
           .Update
           .MoveNext
     Loop
End With

Data2.Refresh
DBGrid2.Columns(0).Width = 500
    
'Data2.Recordset.FindFirst "artigo = '" + x_reg + "'"
Data2.Recordset.FindNext "aceite IS NULL"

    
DBGrid2.SetFocus
    
End Sub

Private Sub Data1_Reposition()
'On Error Resume Next

Set x_selcasos = x_db.OpenRecordset("Select artigo, descritor, fragmento, aceite from busca_descritores_casos where descritor='" + _
                                    Data1.Recordset.Fields("descritor") + "' order by val(artigo)")
Set Data2.Recordset = x_selcasos
DBGrid2.Columns(0).Width = 500
    Label5.Caption = percent_total()

End Sub

Private Sub Data2_Reposition()
    On Error Resume Next
    
    Dim x_txt As String
    
    x_txt = LCase(Data2.Recordset.Fields(2))
    
    'x = single_w_count(x_txt)
    
    x_desc = Trim(Mid(Data2.Recordset.Fields(1), 4, 255))
    
    x_pos = InStr(1, x_txt, " " + Trim(x_desc) + " ")
    x_psel = x_pos
    
    If x_pos = 0 Then
       Label1.Caption = ""
       Exit Sub
    End If
    
    x_res = Mid(x_txt, x_pos - 15, 30 + Len(x_desc))
    
    x_pos = InStr(1, x_res, " ")
    If x_pos = 0 Then
       Label1.Caption = ""
       Exit Sub
    End If
    
    x_res = Mid(x_res, x_pos, Len(x_res))
    
    Label1.Caption = Trim(x_res)
    
    Text2.Text = x_txt
    Text2.Refresh
    Text2.SelStart = x_psel
    Text2.SelLength = Len(x_desc)
    
    
    Text3.Text = Me.neigbors(Label1.Caption, x_desc, 0)
    Text4.Text = Me.neigbors(Label1.Caption, x_desc, 1)
    Label4.Caption = "Descriptor " + Format(Data2.Recordset.PercentPosition, "##0") + "% Complete"

    
    

End Sub




Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

Screen.MousePointer = 11

If KeyCode = 13 Then

Dim x_txt As String

    Set x_sel = x_db.OpenRecordset("Select artigo, descritor, fragmento from busca_descritores_casos where descritor='" + Data1.Recordset.Fields("descritor") + "' ")
    
    With x_sel
         .MoveFirst
         While Not .EOF()
               x_txt = x_txt + IIf(Not IsNull(x_sel.Fields("fragmento")), x_sel.Fields("fragmento"), "")
               .MoveNext
         Wend
    End With
    
    x = single_w_count(x_txt)

End If

Screen.MousePointer = 0
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 78 Then 'NEXT N key

       
    Data2.Recordset.FindNext "aceite IS NULL"

   
    'With Data2.Recordset
    '       Do While Not .EOF
    '          .MoveNext
    '            If .Fields("aceite") = "" Or IsNull(.Fields("aceite")) Then
    '               Exit Do
    '            End If
    '       Loop
    'End With
End If

If KeyCode = 83 Then 'SKIP  S Key
    'With Data2.Recordset
    '     .Edit
    '     .Fields("aceite") = "SKIPED"
    '     .Update
    '      Data2.Recordset.FindNext "aceite IS NULL"
          'Do While Not .EOF
          '    .MoveNext
          '      If .Fields("aceite") = "" Or IsNull(.Fields("aceite")) Then
          '        Exit Do
          '      End If
          ' Loop
    'End With
    
    
    ' analisa vizinhos anterior e posterior com k=0
    x_pre = Trim(Text3.Text)
    x_suc = Trim(Text4.Text)
    
    x_desc = Trim(Mid(Data2.Recordset.Fields(1), 4, 255))
    
    x_query = "Select artigo, descritor, fragmento, aceite from busca_descritores_casos " + _
    "WHERE descritor = '" + Data2.Recordset.Fields(1) + "' AND (" + _
    IIf(x_pre <> "", "fragmento Like '* " + x_pre + " " + x_desc + " *' ", "") + _
    IIf(x_pre <> "" And x_suc <> "", " OR ", "") + _
    IIf(x_suc <> "", "fragmento Like '* " + x_desc + " " + x_suc + " *'", "") + _
    ")"
    
    
    Set x_selcasos = x_db.OpenRecordset(x_query)
    
    x_a = x_selcasos.RecordCount
    With x_selcasos
         Do While Not .EOF
            If IsNull(.Fields("aceite")) Or .Fields("aceite") = "" Then
               .Edit
               .Fields("aceite") = "SKIPED"
               .Update
            End If
            .MoveNext
         Loop
    End With
    
    Data2.Refresh
    DBGrid2.Columns(0).Width = 500
        
    Data2.Recordset.FindNext "aceite IS NULL"
        
    DBGrid2.SetFocus

    
End If

If KeyCode = 77 Then '(UN)MARK AS SKIPED ONLY M key
   With Data2.Recordset
        .Edit
        If .Fields("aceite") = "SKIPED" Then
           .Fields("aceite") = ""
        Else
           .Fields("aceite") = "SKIPED"
        End If
        .Update
   End With

    Data2.Refresh
    DBGrid2.Columns(0).Width = 500
        
    Data2.Recordset.FindNext "aceite IS NULL"
        
    DBGrid2.SetFocus


End If



If KeyCode = 76 Then 'LEARN L Key
   Command8_Click
End If

End Sub

Private Sub Form_Load()
Set x_db = OpenDatabase(App.Path + "\base_acm_crossref.accdb", False, False, "")
Set x_totcasos = x_db.OpenRecordset("Select * from cons_total_casos order by cases desc")
Set Data1.Recordset = x_totcasos
End Sub


Function single_w_count(x_text As String) As String
        
        
        
        List2.Clear
        
        'Dim tmpNum As Integer
        Const onespace As String = " "
        'Const twospace As String = "  "
        Dim tmpWords() As String
        Dim tmpText As String
 
        tmpNum = -1
        tmpText = LCase(x_text)
        'Do Until tmpNum = 0
        '    tmpNum = InStr(tmpText, twospace)
        '    If tmpNum > 0 Then
        '        tmpText = Replace(tmpText, twospace, onespace)
        '    End If
        'Loop
 
        tmpWords = Split(tmpText, onespace)

        single_w_count = "Number of Words: " & UBound(tmpWords) + 1

        Label3.Caption = "Ordering "
        Me.Refresh
        SnakeSort1 tmpWords
         
         
         'For i = 0 To UBound(tmpWords)
             
         '     Label3.Caption = "Ordering " + Str(i) + "/" + Str(UBound(tmpWords))
         '     Me.Refresh
         '    For j = i + 1 To UBound(tmpWords)
         '       If UCase(tmpWords(i)) > UCase(tmpWords(j)) Then
         '          Temp = tmpWords(j)
         '          tmpWords(j) = tmpWords(i)
         '          tmpWords(i) = Temp
         '       End If
         '    Next j
         'Next i
         
         
         x_elem = tmpWords(0)
         x_cont = 1
         List2.Clear
         
         For i = 0 To UBound(tmpWords)
         
              Label3.Caption = "Counting " + Str(i) + "/" + Str(UBound(tmpWords))
              Me.Refresh
              If Len(tmpWords(i)) > 0 Then
                    If Mid(tmpWords(i), Len(tmpWords(i)), 1) = "," Or Mid(tmpWords(i), Len(tmpWords(i)), 1) = "." Then
                       tmpWords(i) = Mid(tmpWords(i), 1, Len(tmpWords(i)) - 1)
                    End If
    
             
                  
                  If x_elem = tmpWords(i) Then
                     x_cont = x_cont + 1
                  Else
                     If x_cont > 1 And Len(x_elem) > 3 Then
                        List2.AddItem Format(x_cont, "000000") + " " + x_elem
                     End If
                     x_elem = tmpWords(i)
                     x_cont = 1
                  End If
              End If
         Next
         
         
         List2.ListIndex = List2.ListCount - 1
         
End Function

Public Sub SnakeSort1(ByRef pvarArray As Variant)
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim lngIndex() As Long
    Dim lngLevel As Long
    Dim lngOldLevel As Long
    Dim lngNewLevel As Long
    Dim varMirror As Variant
    Dim lngDirection As Long
    Dim blnMirror As Boolean
    Dim varSwap As Variant
    
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray)
    ReDim lngIndex((iMax - iMin + 3) \ 2)
    lngIndex(0) = iMin
    i = iMin
    ' Initial loop: locate cutoffs for each ordered section
    Do Until i >= iMax
        Select Case lngDirection
            Case 1
                Do Until i = iMax
                    If pvarArray(i) > pvarArray(i + 1) Then Exit Do
                    i = i + 1
                Loop
            Case -1
                Do Until i = iMax
                    If pvarArray(i) < pvarArray(i + 1) Then Exit Do
                    i = i + 1
                Loop
            Case Else
                Do Until i = iMax
                    If pvarArray(i) <> pvarArray(i + 1) Then Exit Do
                    i = i + 1
                Loop
                If i = iMax Then lngDirection = 1
        End Select
        If lngDirection = 0 Then
            If pvarArray(i) > pvarArray(i + 1) Then
                lngDirection = -1
            Else
                lngDirection = 1
            End If
        Else
            lngLevel = lngLevel + 1
            lngIndex(lngLevel) = i * lngDirection
            lngDirection = 0
        End If
        i = i + 1
    Loop
    If Abs(lngIndex(lngLevel)) < iMax Then
        If lngDirection = 0 Then lngDirection = 1
        lngLevel = lngLevel + 1
        lngIndex(lngLevel) = i * lngDirection
    End If
    ' If the list is already sorted, exit
    If lngLevel <= 1 Then
        ' If sorted descending, reverse before exiting
        If lngIndex(lngLevel) < 0 Then
            For i = 0 To (iMax - iMin) \ 2
                varSwap = pvarArray(iMin + i)
                pvarArray(iMin + i) = pvarArray(iMax - i)
                pvarArray(iMax - i) = varSwap
            Next
        End If
        Exit Sub
    End If
    ' Main loop - merge section pairs together until only one section left
    ReDim varMirror(iMin To iMax)
    Do Until lngLevel = 1
        lngOldLevel = lngLevel
        For lngLevel = 1 To lngLevel - 1 Step 2
            If blnMirror Then
                SnakeSortMerge varMirror, lngIndex(lngLevel - 1), lngIndex(lngLevel), lngIndex(lngLevel + 1), pvarArray
            Else
                SnakeSortMerge pvarArray, lngIndex(lngLevel - 1), lngIndex(lngLevel), lngIndex(lngLevel + 1), varMirror
            End If
            lngNewLevel = lngNewLevel + 1
            lngIndex(lngNewLevel) = Abs(lngIndex(lngLevel + 1))
        Next
        If lngOldLevel Mod 2 = 1 Then
            If blnMirror Then
                For i = lngIndex(lngNewLevel) + 1 To iMax
                    pvarArray(i) = varMirror(i)
                Next
            Else
                For i = lngIndex(lngNewLevel) + 1 To iMax
                    varMirror(i) = pvarArray(i)
                Next
            End If
            lngNewLevel = lngNewLevel + 1
            lngIndex(lngNewLevel) = lngIndex(lngOldLevel)
        End If
        lngLevel = lngNewLevel
        lngNewLevel = 0
        blnMirror = Not blnMirror
    Loop
    ' Copy back to main array if necessary
    If blnMirror Then
        For i = iMin To iMax
            pvarArray(i) = varMirror(i)
        Next
    End If
End Sub
 
Private Sub SnakeSortMerge(pvarSource As Variant, plngLeft As Long, plngMid As Long, plngRight As Long, pvarDest As Variant)
    Dim L As Long
    Dim LMin As Long
    Dim LMax As Long
    Dim LStep As Long
    Dim R As Long
    Dim RMin As Long
    Dim RMax As Long
    Dim RStep As Long
    Dim O As Long
    
    If plngLeft <> 0 Then O = Abs(plngLeft) + 1
    If plngMid > 0 Then
        LMin = O
        LMax = Abs(plngMid)
        LStep = 1
    Else
        LMin = Abs(plngMid)
        LMax = O
        LStep = -1
    End If
    If plngRight > 0 Then
        RMin = Abs(plngMid) + 1
        RMax = Abs(plngRight)
        RStep = 1
    Else
        RMin = Abs(plngRight)
        RMax = Abs(plngMid) + 1
        RStep = -1
    End If
    L = LMin
    R = RMin
    Do
        If pvarSource(L) <= pvarSource(R) Then
            pvarDest(O) = pvarSource(L)
            If L = LMax Then
                For R = R To RMax Step RStep
                    O = O + 1
                    pvarDest(O) = pvarSource(R)
                Next
                Exit Do
            End If
            L = L + LStep
        Else
            pvarDest(O) = pvarSource(R)
            If R = RMax Then
                For L = L To LMax Step LStep
                    O = O + 1
                    pvarDest(O) = pvarSource(L)
                Next
                Exit Do
            End If
            R = R + RStep
        End If
        O = O + 1
    Loop
End Sub

Private Sub Text3_DblClick()
Text3.Text = ""
End Sub


Private Sub Text4_DblClick()
Text4.Text = ""
End Sub


