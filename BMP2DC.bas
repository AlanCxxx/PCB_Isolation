'------------------------------------------
' Bitmap Outline Tracer (for PCB Isolation)
'------------------------------------------
Option Explicit
' Data for processing
Dim bmImage() As Integer
Dim bmWidth As Long
Dim bmHeight As Long
Dim bmScale As Long
' Image Options
Dim Scale As Double
Dim Threshold As Integer
Dim PixelOption As Long
Dim SharpenOption As Long

' Image Option Dialog Box
Begin Dialog OUTLINEOPTIONS 180,41, 157, 196, "Outline Options"
  OptionGroup .PIXELOPTIONS
    OptionButton 12,96,129,12, "Thin outline (-1 pixel overall)"
    OptionButton 12,112,129,12, "Normal outline (better isolation)"
    OptionButton 12,128,129,12, "Normal outine (improved lettering)"
    OptionButton 12,144,129,12, "Thicken outline (+1 pixel overall)"
  GroupBox 4,84,144,80, "Outline Options"
  CheckBox 4, 8,144,12, "Clear any existing  drawing", .Clear
  CheckBox 4,20,144,12, "Sharpen image", .Sharpen
  Text      4,40,96,12, "Image scale (DPI)"
  TextBox 112,40,36,12, .Scale
  Text      4,54,96,12, "Image threshold (0 to 255)"
  TextBox 112,54,36,12, .Threshold
  OKButton 8,172,53,16
  CancelButton 92,172,53,16
End Dialog
Dim Options As OUTLINEOPTIONS

Declare Function GetDlgCtrlID Lib "User32" ( ByVal hwndCtl As Long ) As Long
Declare Function DlgDirListA Lib "User32" ( ByVal hDlg As Long, ByVal PathSpec As String, ByVal nIDListBox As Long, ByVal nIDStatic As Long, ByVal nFileType As Long ) As Long
Static PathName As String
Static FileName As String
Static nListID As Integer
Static nPathID As Integer
Static FileSpec As String
Dim MyList$()
Begin Dialog BitmapDlg 60, 60, 285, 225, "Select BMP", .DlgFunc
 TextBox 10, 5, 80, 12, .Text1
 Text 100, 8, 178, 9, "", .Path1
 ListBox 10, 20, 80, 178, MyList(), .List1, 2  ' Sorted
 Picture 100, 20, 178, 178, "", 0, .Picture1
 CancelButton 42, 203, 40, 12    
 OKButton 90, 203, 40, 12
End Dialog
Dim frame As BitmapDlg

' -------------------------------------
' BEGIN
' -------------------------------------

 ' Show the bitmap
 If (Dialog(frame)=0) Then
   ' MsgBox "Cancel"
 Else
   ' MsgBox "Path "+PathName
   ' MsgBox "File Name "+FileName
   Call ReadBMP(FileName)
   ' MsgBox "BMP Width  "+Format(bmWidth)
   ' MsgBox "BMP Height "+Format(bmHeight)
   ' MsgBox "BMP Scale "+Format(bmScale)
   ' GetOptions
   Options.Clear=1
   Options.Sharpen=0
   Options.Threshold=128
   Options.Scale=-Int(-(bmScale/50.8))*2
   Options.PixelOptions=2
   If (Dialog(Options)=0) Then
     ' MsgBox "Cancel"
   Else 
     Threshold=Options.Threshold
     Scale=Int(Abs(Options.Scale)+0.5)
     If (Scale=0) Then Scale=72
     PixelOption=Options.PixelOptions
     SharpenOption=Options.Sharpen

     ' Clear drawing
     If (Options.Clear) Then
       dcCreatePoint 0,0
       dcSelectAll
       dcEraseSelObjs
       dcSetDrawingUnits dcMillimeters 
     End If
     Call ProcessBMP()

     dcSetPointParms dcRED, True
     dcSetLineParms dcBLACK, dcSOLID, dcNORMAL
     dcViewAll
     dcUpdateDisplay True

   End If

 End If

 End 
' -------------------------------------
' END
' -------------------------------------

' Used by the File Manager (with BMP Previewer)
Function DlgFunc(controlID As String, action As Integer, suppValue As Integer )
 Select Case action
 Case 1 ' Initialize
   PathName = "."
   DlgText "Text1", PathName & "\*.bmp"
   nListID = GetDlgCtrlID( DlgControlHWND("List1") )
   nPathID = GetDlgCtrlID( DlgControlHWND("Path1") )
   DlgDirListA DlgHWND, DlgText("Text1"), nListID, nPathID, &h10
 Case 2 ' Click
   If controlID = "OK" Then        ' OK Button
     DlgDirListA DlgHWND, DlgText("Text1"), nListID, nPathID, &h10
     PathName = DlgText("Path1")
     DlgFunc = 1
   Else
     FileSpec = DlgText("List1")
     If Left(FileSpec,1) = "[" Then        ' FileSpec is a directory
       PathName = Mid( FileSpec, 2, Len(FileSpec)-2 )
       DlgDirListA DlgHWND, PathName & "\*.bmp", nListID, nPathID, &h10
       PathName = DlgText("Path1")
       DlgText "Text1", PathName & "\*.bmp"
       Exit Function
     End If
     FileName = PathName & "\" & DlgText("List1")
     DlgText "Path1", fileName
     DlgSetPicture "Picture1", fileName
   End If
 End Select
End Function

' ---------------------------------------------------------------
' ReadFile using the Win32 API
' ---------------------------------------------------------------
Const GENERIC_WRITE = &H40000000
Const GENERIC_READ = &H80000000
Const FILE_ATTRIBUTE_NORMAL = &H80
Const CREATE_ALWAYS = 2
Const OPEN_ALWAYS = 4
Const INVALID_HANDLE_VALUE = -1

Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
  lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
  lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Declare Function CloseHandle Lib "kernel32" ( _
 ByVal hObject As Long) As Long

Declare Function WriteFile Lib "kernel32" ( _
 ByVal hFile As Long, lpBuffer As Any, _
 ByVal nNumberOfBytesToWrite As Long, _
 lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Declare Function CreateFile Lib "kernel32" _
 Alias "CreateFileA" (ByVal lpFileName As String, _
 ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
 ByVal lpSecurityAttributes As Long, _
 ByVal dwCreationDisposition As Long, _
 ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) _
 As Long

Declare Function FlushFileBuffers Lib "kernel32" ( _
 ByVal hFile As Long) As Long

Sub ReadBMP(FileName As String)
 On Error GoTo ErrorHandler
 Dim Handle As Long
 Dim Success As Long
 Dim BytesRead As Long
 Dim BytesToRead As Long
 Dim Offset As Long
 Dim Image() As Byte
 Dim RowOffset As Long
 Dim i As  Long
 Dim j As  Long
 Dim k As  Long
 Dim l As  Long
 Dim B0 As Byte
 Dim B1 As Byte
 Dim B2 As Byte

 BytesToRead = FileLen(FileName)
 ReDim Image(BytesToRead)
 ' Get a handle to file
 Handle=CreateFile(FileName,GENERIC_WRITE Or GENERIC_READ,0,0,OPEN_ALWAYS,FILE_ATTRIBUTE_NORMAL,0)
 If (Handle<>INVALID_HANDLE_VALUE) Then
   ' Read file contents into a byte array
   Success=ReadFile(Handle,Image(0),BytesToRead,BytesRead,0)
   CloseHandle Handle
   If (Success) Then
     ' MsgBox "Success "+Format(BytesRead)+" out of "+Format(BytesToRead)
   Else
     GoTo ErrorHandler
   End If
 
   ' Read the BMP header
   If (CInt(Image(0))<>66) Then
     MsgBox "Not a BMP file"
     End
   End If
   Offset=CLng(Image(10))
   ' MsgBox "Offset "+Format(Offset)
   bmWidth=CLng(Image(18))
   ' MsgBox "Width "+Format(bmWidth)
   bmHeight=CLng(Image(22))
   ' MsgBox "Height "+Format(bmHeight)
   If (CInt(Image(28))<>24) Then
     MsgBox "Only 24 bit BMP files"
     End
   End If
   If (CLng(Image(30))<>0) Then
     MsgBox "Only uncompressed BMP files"
     End
   End If
   bmScale=CLng(Image(38))
   If (bmScale=0) Then bmScale=1829 ' 72 dpi
   ' Convert to integer array (Cyprus Enable does not support byte maths)
   ' This is where all the Time is spent loading the file!!!!!!!!!!!!!!!!
   l=0
   k=Offset
   ReDim bmImage(bmWidth*bmHeight)
   For j=0 to bmHeight-1
     For i=0 to bmWidth-1
       B0=Image(k)
       B1=Image(k+1)
       B2=Image(k+2)
       k=k+3
       bmImage(l)=CInt(B0)+CInt(B1)+CInt(B2)
       l=l+1
     Next i
     k=k+(4-3*bmWidth Mod 4) Mod 4
   Next j

 End If

 Exit Sub

 ErrorHandler:
   MsgBox "Unable to open file?"
   End

End Sub

Sub ProcessBMP()
 ' Image copy for sharpen
 Dim bmCopy() As Long
 Dim pMax As Long
 Dim pMin As Long

 ' Segments
 Dim px1() As Long
 Dim py1() As Long
 Dim px2() As Long
 Dim py2() As Long
 Dim available() As Boolean
 Dim segs As Long

 ' General
 Dim i As Long
 Dim j As Long
 Dim k As Long
 Dim t As Long
 Dim p0 As Long
 Dim p1 As Long
 Dim p2 As Long
 Dim p3 As Long

 ' Binary Search
 Dim L As Long
 Dim R As Long
 Dim M As Long 
 Dim A As Long
 Dim B As Long

 ' Sort
 Dim Flag As Boolean
 Dim tx1 As Long
 Dim ty1 As Long
 Dim tx2 As Long
 Dim ty2 As Long

 ' String (=PolyLine)
 Dim NextPtr() As Long
 Dim StrId() As Long
 Dim StrN() As Long
 Dim StrX() As Double
 Dim StrY() As Double
 Dim StrA() As Double
 Dim StrCnt As Long
 Dim StrPtn As Long
 Dim Shape() As Double

 ' -----------------------------------
 ' Process the image
 ' -----------------------------------
 ' Sharpen Image
 If (SharpenOption=1) Then
   ReDim bmCopy(bmWidth*bmHeight)
   For j=1 to bmHeight-2
     k=j*bmWidth+1
     For i=1 to bmWidth-2
       pMin=255
       pMax=0
       If (bmImage(k-bmWidth-1)>pMax) Then pMax=bmImage(k-bmWidth-1)
       If (bmImage(k-bmWidth  )>pMax) Then pMax=bmImage(k-bmWidth  )
       If (bmImage(k-bmWidth+1)>pMax) Then pMax=bmImage(k-bmWidth+1)
       If (bmImage(k        -1)>pMax) Then pMax=bmImage(k        -1)
       If (bmImage(k          )>pMax) Then pMax=bmImage(k          )
       If (bmImage(k        +1)>pMax) Then pMax=bmImage(k        +1)
       If (bmImage(k+bmWidth-1)>pMax) Then pMax=bmImage(k+bmWidth-1)
       If (bmImage(k+bmWidth  )>pMax) Then pMax=bmImage(k+bmWidth  )
       If (bmImage(k+bmWidth+1)>pMax) Then pMax=bmImage(k+bmWidth+1)
       If (bmImage(k-bmWidth-1)<pMin) Then pMin=bmImage(k-bmWidth-1)
       If (bmImage(k-bmWidth  )<pMin) Then pMin=bmImage(k-bmWidth  )
       If (bmImage(k-bmWidth+1)<pMin) Then pMin=bmImage(k-bmWidth+1)
       If (bmImage(k        -1)<pMin) Then pMin=bmImage(k        -1)
       If (bmImage(k          )<pMin) Then pMin=bmImage(k          )
       If (bmImage(k        +1)<pMin) Then pMin=bmImage(k        +1)
       If (bmImage(k+bmWidth-1)<pMin) Then pMin=bmImage(k+bmWidth-1)
       If (bmImage(k+bmWidth  )<pMin) Then pMin=bmImage(k+bmWidth  )
       If (bmImage(k+bmWidth+1)<pMin) Then pMin=bmImage(k+bmWidth+1)
       If (bmImage(k)*2>=pMax+pMin) Then
         bmCopy(k)=pMax
       Else
         bmCopy(k)=pMin
       End If
       k=k+1
     Next i
   Next j
   For j=1 to bmHeight-2
     k=j*bmWidth+1
     For i=1 to bmWidth-2
       bmImage(k)=bmCopy(k)
       k=k+1
     Next i
   Next j
 End If 

 ' Set Black and White
 t=3*Threshold
 For j=1 to bmHeight-2
   k=j*bmWidth+1
   For i=1 to bmWidth-2
     If (bmImage(k)>=t) Then
       bmImage(k)=0
     Else
       bmImage(k)=1
     End If
     k=k+1
   Next i
 Next j
 ' Clear boarders
 For j=0 to bmHeight-1
   bmImage(j*bmWidth)=0
   bmImage(j*bmWidth+bmWidth-1)=0
 Next j
 For i=0 to bmWidth-1
   bmImage(i)=0
   bmImage(i+(bmHeight-1)*bmWidth)=0
 Next i
 
 ' -----------------------------------
 ' Create segment arrays (worst case)
 ' -----------------------------------
 ReDim px1(bmWidth*bmHeight)
 ReDim py1(bmWidth*bmHeight)
 ReDim px2(bmWidth*bmHeight)
 ReDim py2(bmWidth*bmHeight)
 ReDim available(bmWidth*bmHeight)

 ' ------------------------------------------------------
 ' Load segments - Careful not to change the search range
 ' ------------------------------------------------------
 segs=0
 If (PixelOption=0) Then

   ' Thin lines 
   segs=0
   For j=1 to bmHeight-1
     k=j*bmWidth
     For i=0 to bmWidth-2
       p0=bmImage(k-bmWidth)
       p1=bmImage(k-bmWidth+1)
       p2=bmImage(k)
       p3=bmImage(k+1)
       If ((p0=1) And (p1=1) And (p2=1) And (p3=0)) Then
         '**
         '*.
          px1(segs)=i
          py1(segs)=j
          px2(segs)=i+1
          py2(segs)=j-1
          segs=segs+1
       ElseIf ((p0=1) And (p1=0) And (p2=1) And (p3=1)) Then
         '*.
         '**
         px1(segs)=i+1
         py1(segs)=j
         px2(segs)=i
         py2(segs)=j-1
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=1) And (p3=1)) Then
         '.*
         '**
         px1(segs)=i+1
         py1(segs)=j-1
         px2(segs)=i
         py2(segs)=j
         segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=0) And (p3=1)) Then
         '**
         '.*
         px1(segs)=i
         py1(segs)=j-1
         px2(segs)=i+1
         py2(segs)=j
         segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=0) And (p3=0)) Then
         '**
         '..
         px1(segs)=i
         py1(segs)=j-1
         px2(segs)=i+1
         py2(segs)=j-1
         segs=segs+1 
       ElseIf ((p0=1) And (p1=0) And (p2=1) And (p3=0)) Then
         '*.
         '*.
         px1(segs)=i
         py1(segs)=j
         px2(segs)=i
         py2(segs)=j-1
         segs=segs+1  
       ElseIf ((p0=0) And (p1=0) And (p2=1) And (p3=1)) Then
         '..
         '**
         px1(segs)=i+1
         py1(segs)=j
         px2(segs)=i
         py2(segs)=j
         segs=segs+1 
       ElseIf ((p0=0) And (p1=1) And (p2=0) And (p3=1)) Then
         '.*
         '.*
         px1(segs)=i+1
         py1(segs)=j-1
         px2(segs)=i+1
         py2(segs)=j
         segs=segs+1
       End If
       k=k+1
       ' ' Show points
       ' If (p2=1) Then
       '   dcCreatePoint i,j
       ' End If
     Next i
   Next j

 ElseIf (PixelOption=1) Then

   ' Normal lines    
   For j=1 to bmHeight-1
     k=j*bmWidth
     For i=0 to bmWidth-2
       p0=bmImage(k-bmWidth)
       p1=bmImage(k-bmWidth+1)
       p2=bmImage(k)
       p3=bmImage(k+1)
       If ((p0=1) And (p1=0) And (p2=1) And (p3=1)) Then
         '*.
         '**
          px1(segs)=2*i+2
          py1(segs)=2*j-1
          px2(segs)=2*i+1
          py2(segs)=2*j-2
          segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=1) And (p3=0)) Then
         '** 
         '*. 
         px1(segs)=2*i+1
         py1(segs)=2*j
         px2(segs)=2*i+2
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=0) And (p3=1)) Then
         '**
         '.*
         px1(segs)=2*i
         py1(segs)=2*j-1
         px2(segs)=2*i+1
         py2(segs)=2*j
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=1) And (p3=1)) Then
         '.*
         '**
         px1(segs)=2*i+1
         py1(segs)=2*j-2
         px2(segs)=2*i
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=1) And (p1=0) And (p2=1) And (p3=0)) Then
         '*.
         '*.
         px1(segs)=2*i+1
         py1(segs)=2*j
         px2(segs)=2*i+1
         py2(segs)=2*j-2
         segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=0) And (p3=0)) Then
         '**
         '..
         px1(segs)=2*i
         py1(segs)=2*j-1
         px2(segs)=2*i+2
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=0) And (p3=1)) Then
         '.*
         '.*
         px1(segs)=2*i+1
         py1(segs)=2*j-2
         px2(segs)=2*i+1
         py2(segs)=2*j
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=1) And (p3=1)) Then
         '..
         '**
         px1(segs)=2*i+2
         py1(segs)=2*j-1
         px2(segs)=2*i
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=1) And (p3=0)) Then
         '*.
         '.*
         px1(segs)=2*i+1
         py1(segs)=2*j
         px2(segs)=2*i
         py2(segs)=2*j-1
         segs=segs+1
         px1(segs)=2*i+1
         py1(segs)=2*j-2
         px2(segs)=2*i+2
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=1) And (p1=0) And (p2=0) And (p3=1)) Then
         '.*
         '*.
         px1(segs)=2*i+2
         py1(segs)=2*j-1
         px2(segs)=2*i+1
         py2(segs)=2*j
         segs=segs+1
         px1(segs)=2*i
         py1(segs)=2*j-1
         px2(segs)=2*i+1
         py2(segs)=2*j-2
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=0) And (p3=1)) Then
         '.*
         '..
         px1(segs)=2*i+2
         py1(segs)=2*j-1
         px2(segs)=2*i+1
         py2(segs)=2*j
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=1) And (p3=0)) Then
         '*.
         '..
         px1(segs)=2*i+1
         py1(segs)=2*j
         px2(segs)=2*i
         py2(segs)=2*j-1
         segs=segs+1  
       ElseIf ((p0=1) And (p1=0) And (p2=0) And (p3=0)) Then
         '..
         '*.
         px1(segs)=2*i
         py1(segs)=2*j-1
         px2(segs)=2*i+1
         py2(segs)=2*j-2
         segs=segs+1

       ElseIf ((p0=0) And (p1=1) And (p2=0) And (p3=0)) Then
         '..
         '.*
         px1(segs)=2*i+1
         py1(segs)=2*j-2
         px2(segs)=2*i+2
         py2(segs)=2*j-1
         segs=segs+1 
       End If
       k=k+1
       ' ' Show points
       ' If (p2=1) Then
       '   dcCreatePoint 2*i,2*j
       ' End If
     Next i
   Next j


ElseIf (PixelOption=2) Then

   ' Normal lines    
   For j=1 to bmHeight-1
     k=j*bmWidth
     For i=0 to bmWidth-2
       p0=bmImage(k-bmWidth)
       p1=bmImage(k-bmWidth+1)
       p2=bmImage(k)
       p3=bmImage(k+1)
       If ((p0=1) And (p1=0) And (p2=1) And (p3=1)) Then
         '*.
         '**
          px1(segs)=2*i+2
          py1(segs)=2*j-1
          px2(segs)=2*i+1
          py2(segs)=2*j-2
          segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=1) And (p3=0)) Then
         '** 
         '*. 
         px1(segs)=2*i+1
         py1(segs)=2*j
         px2(segs)=2*i+2
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=0) And (p3=1)) Then
         '**
         '.*
         px1(segs)=2*i
         py1(segs)=2*j-1
         px2(segs)=2*i+1
         py2(segs)=2*j
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=1) And (p3=1)) Then
         '.*
         '**
         px1(segs)=2*i+1
         py1(segs)=2*j-2
         px2(segs)=2*i
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=1) And (p1=0) And (p2=1) And (p3=0)) Then
         '*.
         '*.
         px1(segs)=2*i+1
         py1(segs)=2*j
         px2(segs)=2*i+1
         py2(segs)=2*j-2
         segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=0) And (p3=0)) Then
         '**
         '..
         px1(segs)=2*i
         py1(segs)=2*j-1
         px2(segs)=2*i+2
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=0) And (p3=1)) Then
         '.*
         '.*
         px1(segs)=2*i+1
         py1(segs)=2*j-2
         px2(segs)=2*i+1
         py2(segs)=2*j
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=1) And (p3=1)) Then
         '..
         '**
         px1(segs)=2*i+2
         py1(segs)=2*j-1
         px2(segs)=2*i
         py2(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=1) And (p3=0)) Then
         '*.
         '.*
         px2(segs)=2*i+2
         py2(segs)=2*j-1
         px1(segs)=2*i+1
         py1(segs)=2*j
         segs=segs+1
         px2(segs)=2*i
         py2(segs)=2*j-1
         px1(segs)=2*i+1
         py1(segs)=2*j-2
         segs=segs+1
       ElseIf ((p0=1) And (p1=0) And (p2=0) And (p3=1)) Then
         '.*
         '*.
         px2(segs)=2*i+1
         py2(segs)=2*j
         px1(segs)=2*i
         py1(segs)=2*j-1
         segs=segs+1
         px2(segs)=2*i+1
         py2(segs)=2*j-2
         px1(segs)=2*i+2
         py1(segs)=2*j-1
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=0) And (p3=1)) Then
         '.*
         '..
         px1(segs)=2*i+2
         py1(segs)=2*j-1
         px2(segs)=2*i+1
         py2(segs)=2*j
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=1) And (p3=0)) Then
         '*.
         '..
         px1(segs)=2*i+1
         py1(segs)=2*j
         px2(segs)=2*i
         py2(segs)=2*j-1
         segs=segs+1  
       ElseIf ((p0=1) And (p1=0) And (p2=0) And (p3=0)) Then
         '..
         '*.
         px1(segs)=2*i
         py1(segs)=2*j-1
         px2(segs)=2*i+1
         py2(segs)=2*j-2
         segs=segs+1

       ElseIf ((p0=0) And (p1=1) And (p2=0) And (p3=0)) Then
         '..
         '.*
         px1(segs)=2*i+1
         py1(segs)=2*j-2
         px2(segs)=2*i+2
         py2(segs)=2*j-1
         segs=segs+1 
       End If
       k=k+1
       ' ' Show points
       ' If (p2=1) Then
       '   dcCreatePoint 2*i,2*j
       ' End If
     Next i
   Next j


 ElseIf (PixelOption=3) Then

   ' Thick lines    
   For j=1 to bmHeight-1
     k=j*bmWidth
     For i=0 to bmWidth-2
       p0=bmImage(k-bmWidth)
       p1=bmImage(k-bmWidth+1)
       p2=bmImage(k)
       p3=bmImage(k+1)
       If ((p0=1) And (p1=0) And (p2=1) And (p3=0)) Then
         '*.
         '*.
          px1(segs)=i+1
          py1(segs)=j
          px2(segs)=i+1
          py2(segs)=j-1
          segs=segs+1
       ElseIf ((p0=1) And (p1=1) And (p2=0) And (p3=0)) Then
         '**
         '..
         px1(segs)=i
         py1(segs)=j
         px2(segs)=i+1
         py2(segs)=j
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=0) And (p3=1)) Then
         '.*
         '.*
         px1(segs)=i
         py1(segs)=j-1
         px2(segs)=i
         py2(segs)=j
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=1) And (p3=1)) Then
         '..
         '**
         px1(segs)=i+1
         py1(segs)=j-1
         px2(segs)=i
         py2(segs)=j-1
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=0) And (p3=1)) Then
         '..
         '.*
         px1(segs)=i+1
         py1(segs)=j-1
         px2(segs)=i
         py2(segs)=j
         segs=segs+1
       ElseIf ((p0=0) And (p1=0) And (p2=1) And (p3=0)) Then
         '..
         '*.
         px1(segs)=i+1
         py1(segs)=j
         px2(segs)=i
         py2(segs)=j-1
         segs=segs+1 
       ElseIf ((p0=1) And (p1=0) And (p2=0) And (p3=0)) Then
         '*.
         '..
         px1(segs)=i
         py1(segs)=j
         px2(segs)=i+1
         py2(segs)=j-1
         segs=segs+1
       ElseIf ((p0=0) And (p1=1) And (p2=0) And (p3=0)) Then
         '.*
         '..
         px1(segs)=i
         py1(segs)=j-1
         px2(segs)=i+1
         py2(segs)=j
         segs=segs+1 
       End If
       k=k+1
       ' ' Show points
       ' If (p2=1) Then
       '   dcCreatePoint i,j
       ' End If
     Next i
   Next j

 End If

 ' ' Show segments
 ' For i=0 to segs-1 
 '   dcSetLineParms dcBLACK,dcSOLID,dcTHIN
 '   dcCreateLine px1(i),py1(i),px2(i),py2(i)
 '   dcSetLineParms dcBLACK,dcSOLID,dcTHICK
 '   dcCreateLine 0.2*px1(i)+0.8*px2(i),0.2*py1(i)+0.8*py2(i),px2(i),py2(i)
 ' Next i
 ' MsgBox "Segs   "+Str(Segs)
 If (Segs=0) Then
   MsgBox "No Segments generatered"
   End
 End If 

 ' ------------------------------------
 ' Delete simple lines (segments pairs)
 ' ------------------------------------
 ' Quick Sort (p1)
 Dim bx(1000) As Long
 Dim ex(1000) As Long

 i=0
 bx(0)=0
 ex(0)=segs
 While (i>=0)
   L=bx(i)
   R=ex(i)-1
   If (L<R) Then
     If (i=1000) Then
       MsgBox "Increase QSort Stack size"
       End
     End If
     tx1=px1(L)
     ty1=py1(L)
     tx2=px2(L)
     ty2=py2(L)
     While (L<R)
       While (((px1(R)>tx1) Or (px1(R)=tx1) And (py1(R)>ty1)) And (L<R))
         R=R-1
       Wend
       If (L<R) Then
         px1(L)=px1(R)
         py1(L)=py1(R)
         px2(L)=px2(R)
         py2(L)=py2(R)
         L=L+1
       End If
       While (((px1(L)<tx1) Or (px1(L)=tx1) And (py1(L)<ty1)) And (L<R))
         L=L+1
       Wend
       If (L<R) Then
         px1(R)=px1(L)
         py1(R)=py1(L)
         px2(R)=px2(L)
         py2(R)=py2(L)
         R=R-1
       End If
     Wend
     px1(L)=tx1
     py1(L)=ty1
     px2(L)=tx2
     py2(L)=ty2
     bx(i+1)=L+1
     ex(i+1)=ex(i)
     ex(i)=L
     i=i+1
   Else 
     i=i-1
   End If
 Wend

 ' Binary search for segment pairs (set available()=False if a pair)
 For i=0 to segs-1
   available(i)=True
 Next i
 For i=0 to segs-1
   If (available(i)) Then
     L=0
     R=segs-1
     While (L<=R)
       M=Int((L+R)/2)
       If ((px1(M)=px2(i)) And (py1(M)=py2(i))) Then
         ' Scan for the range A to B
         A=M
         While ((A>0) And (px1(A)=px2(i)) And (py1(A)=py2(i)))
           A=A-1
         Wend
         If ((px1(A)<>px2(i)) Or (py1(A)<>py2(i))) Then A=A+1
         B=M
         While ((B<segs-1) And (px1(B)=px2(i)) And (py1(B)=py2(i)))
           B=B+1
         Wend
         If ((px1(B)<>px2(i)) Or (py1(B)<>py2(i))) Then B=B-1
         For k=A to B
           ' Test if a pair
           If ((px2(k)=px1(i)) And (py2(k)=py1(i))) Then
             available(i)=False
             available(k)=False
           End If
         Next k 
         L=R+1
       Else
         If ((px1(M)<px2(i)) Or ((px1(M)=px2(i))  And (py1(M)<py2(i)))) Then
           L=M+1
         Else   
           R=M-1
         End If
       End If
     Wend
   End If
 Next i

 ' Delete segments that are not available()=True
 j=0
 For i=0 to segs-1
   If (available(i)) Then
     px1(j)=px1(i)
     py1(j)=py1(i)
     px2(j)=px2(i)
     py2(j)=py2(i)
     available(j)=available(i)
     j=j+1
   End If
 Next i
 segs=j
 If (Segs=0) Then
   MsgBox "No Segments Remaining"
   End
 End If 

 ' Set up an index for segment links
 ReDim NextPtr(segs)
 For i=0 to segs-1
   NextPtr(i)=segs
   available(i)=True
 Next i

 For i=0 to segs-1
   ' Find all p1 that match target p2(i)
   L=0
   R=segs-1
   While (L<=R)
     M=Int((L+R)/2)
     If ((px1(M)=px2(i)) And (py1(M)=py2(i))) Then
       ' Scan for the range A to B
       A=M
       While ((A>0) And (px1(A)=px2(i)) And (py1(A)=py2(i)))
         A=A-1
       Wend
       If ((px1(A)<>px2(i)) Or (py1(A)<>py2(i))) Then A=A+1
       B=M
       While ((B<segs-1) And (px1(B)=px2(i)) And (py1(B)=py2(i)))
         B=B+1
       Wend
       If ((px1(B)<>px2(i)) Or (py1(B)<>py2(i))) Then B=B-1
       For k=A to B
         If (available(k)) Then
           NextPtr(i)=k
           available(k)=False
           Exit For
          End If
       Next k 
       L=R+1
     Else
       If ((px1(M)<px2(i)) Or ((px1(M)=px2(i))  And (py1(M)<py2(i)))) Then
         L=M+1
       Else   
         R=M-1
       End If
     End If
   Wend
 Next i

 ' Get number of string points
 For i=0 to segs-1
   available(i)=True
 Next i
 j=0
 For i=0 to segs-1
   If (available(i)) Then
     k=i
     While (available(k))
       j=j+1
       available(k)=False
       k=NextPtr(k)
     Wend
     j=j+1
   End If
 Next i
 StrPtn=j
 ' MsgBox "StrPtn "+Str(StrPtn)

 ' ---------------------------------------
 ' Create string
 ' ---------------------------------------
 ReDim StrN(StrPtn)
 ReDim StrX(StrPtn)
 ReDim StrY(StrPtn)
 ReDim StrA(StrPtn)

 For i=0 to segs-1
   available(i)=True
 Next i
 StrCnt=0
 j=0
 For i=0 to segs-1
   If (available(i)) Then
     k=i
     StrCnt=StrCnt+1
     While (available(k))
       StrN(j)=StrCnt
       StrX(j)=px1(k)
       StrY(j)=py1(k)
       j=j+1
       StrN(j)=StrCnt
       StrX(j)=px2(k)
       StrY(j)=py2(k)
       available(k)=False
       k=NextPtr(k)
     Wend
     j=j+1
   End If
 Next i
 
 ' Rescale if Option 1 or 2 is selected
 If ((PixelOption=1) Or (PixelOption=2)) Then
   For i=0 to StrPtn-1
     StrX(i)=StrX(i)/2
     StrY(i)=StrY(i)/2
   Next i
 End If
 
 ' -----------------------------------
 ' Filter points
 ' -----------------------------------
 ' Show Strings
 . dcSetLineParms dcBLACK,dcSOLID,dcTHIN
 ' For i=0 to StrPtn-2
 '   If (StrN(i)=StrN(i+1)) Then
 '     dcCreateLine StrX(i),StrY(i),StrX(i+1),StrY(i+1)
 '   End If
 ' Next i
 for t=3 to 1 step -1  
   for k=1 to 5
     ' Calculate area
     StrA(0)=1000
     StrA(StrPtn-1)=1000
     For i=1 to StrPtn-2
       StrA(i)=((StrY(i)-StrY(i-1))*(StrX(i+1)-StrX(i-1))-(StrY(i+1)-StrY(i-1))*(StrX(i)-StrX(i-1)))/2
       If (StrN(i-1)<>StrN(i)) Then StrA(i)=1000
       If (StrN(i)<>StrN(i+1)) Then StrA(i)=1000
     Next i
     ' Test every second point
     j=1
     For i=1 to StrPtn-1
       If ((t*Abs(StrA(i))>1) or (i Mod 2<>0 )) Then
         StrN(j)=StrN(i)
         StrX(j)=StrX(i)
         StrY(j)=StrY(i)
         j=j+1 
       End If
     Next i
     StrPtn=j
   Next k
 Next t
 ' Show Strings
 ' dcSetLineParms dcRED,dcSOLID,dcTHIN
 ' For i=0 to StrPtn-2
 '   If (StrN(i)=StrN(i+1)) Then
 '     dcCreateLine StrX(i),StrY(i),StrX(i+1),StrY(i+1)
 '   End If
 ' Next i

 ' ------------------------------------------
 ' Determine string area for string direction
 ' ------------------------------------------ 
 j=0
 StrA(j)=0
 For i=0 to StrPtn-2
   If (StrN(i)=StrN(i+1)) Then 
     StrA(j)=StrA(j)+(StrY(i)*StrX(i+1)-StrY(i+1)*StrX(i))/2
   Else
     If (Abs(StrA(j))<1e-6) Then StrA(j)=0
     j=i+1     
     StrA(j)=0
   End If
 Next i

 ' -----------------------------------
 ' Scale and offset string
 ' -----------------------------------
 For i=0 to StrPtn-1
   StrX(i)=(StrX(i)+0.5)*25.4/Scale
   StrY(i)=(StrY(i)+0.5)*25.4/Scale
 Next i

 ' -----------------------------------------------
 ' Convert strings to Shape and export to DeltaCad  
 ' ----------------------------------------------- 
 A=0
 For i=0 to StrPtn-1
   If (i<StrPtn-1) Then
     StrCnt=StrN(i+1)
   Else
     StrCnt=StrCnt+1
   End If
   If (StrN(i)<>StrCnt) Then
     B=i-A
     If ((B>=3) And (B<=3072)) Then

       ' Export as shapes
       If (StrA(A)>0) Then
         dcSetShapesParms dcBLUE,dcSOLID,dcTHIN
       ElseIf (StrA(A)<0) Then
         dcSetShapesParms dcRED,dcSOLID,dcTHIN
       Else
         ' Not interested in zero area shapes
         dcSetShapesParms dcBLACK,dcSOLID,dcTHIN
       End If
       
       ReDim Shape(2*B)
       For j=A to i-1
         Shape(2*(j-A)+1)=StrX(j)
         Shape(2*(j-A)+2)=StrY(j)
       Next j
       dcCreateShape Shape(1), B

     ElseIf (B>3072) Then

       ' Export as ordered segments
       dcSetLineParms dcBLACK,dcSOLID,dcTHIN
       For j=A to i-1 
         dcCreateLine StrX(j),StrY(j),StrX(j+1),StrY(j+1)
       Next j

     End If
     A=i+1
   End If
 Next i

 ' Put a box around the area
 dcSetShapesParms dcBLACK,dcSOLID,dcTHIN
 dcCreateBox 0,0,bmWidth*25.4/Scale,bmHeight*25.4/Scale

End Sub