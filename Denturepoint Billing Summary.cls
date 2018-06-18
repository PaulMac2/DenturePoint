VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call ReadDataFromCloseFile
End Sub

Sub ReadDataFromCloseFile()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim CalcWb As Workbook
    Dim SrcWb As Workbook
    Dim SrcFileStart As Integer
    Dim SrcFile As String
    Dim SrcSheet As String
    Dim TradeSheetOpen As String
    Dim YearReference As String
    Dim YearArray As Variant
    Dim YearOffset As Integer
    Dim YearElement As String

    Set CalcWb = ThisWorkbook
    CalcWb.Activate
    ActiveSheet.Unprotect Password:="lock"
    
    Dim iCnt As Integer         ' COUNTER.
    Dim iCntOffset As Integer
    Dim SrcCntOffset As Integer
    
    '-----------------------------------------------------------------
    ' Sheet Resets
    
    Worksheets("Summary").Range("C7:J16").Formula = ""
    Worksheets("Summary").Range("B20:J29").Formula = ""
    Worksheets("Summary").Range("B33:J57").Formula = ""
    
    '-----------------------------------------------------------------
    
    ' Get each of the seperate client job sheets
    
    iCntOffset = 7  ' line offset on Summary page
    SrcSheet = "Patients"
            
    For iCnt = iCntOffset To iCntOffset + 10
    
        SrcFile = Worksheets("Summary").Range("L" & iCnt).Value
        
        If Not Trim(SrcFile) = "" Then
        
            SrcFile = Replace(SrcFile, "/", Application.PathSeparator)
            SrcFile = Replace(SrcFile, "\", Application.PathSeparator)
            SrcFile = Replace(SrcFile, ":", Application.PathSeparator)
            ' SrcFileStart = InStrRev(SrcFile, Application.PathSeparator)
            ' SrcFile = Right(SrcFile, Len(SrcFile) - InStrRev(SrcFile, Application.PathSeparator))
            ' SrcFile = ActiveWorkbook.Path & Application.PathSeparator & SrcFile
            
            If AlreadyOpen(SrcFile) Then
                TradeSheetOpen = True
                Set SrcWb = Workbooks(SrcFile)
            Else
                TradeSheetOpen = False
                Set SrcWb = Workbooks.Open(SrcFile, True, True)
            End If
    
            Worksheets("Summary").Range("O" & (iCnt)).Formula = SrcWb.Worksheets(SrcSheet).Range("K2").Value
            Worksheets("Summary").Range("C" & (iCnt)).Formula = SrcWb.Worksheets(SrcSheet).Range("E2").Value
            Worksheets("Summary").Range("E" & (iCnt)).Formula = SrcWb.Worksheets(SrcSheet).Range("E3").Value
            Worksheets("Summary").Range("G" & (iCnt)).Formula = SrcWb.Worksheets(SrcSheet).Range("E4").Value
            Worksheets("Summary").Range("I" & (iCnt)).Formula = SrcWb.Worksheets(SrcSheet).Range("R4").Value
        
            ' CLOSE THE SOURCE FILE.
            If TradeSheetOpen Then
                ' do nothing
            Else
                SrcWb.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
                Set SrcWb = Nothing
            End If
    
            End If
    Next iCnt
 
    '-----------------------------------------------------------------
       
    ' Get the Asset summaries from the Asset register.
    
    SrcSheet = "Asset Categories"
    SrcFile = Worksheets("Summary").Range("L20").Value
     
    If Not Trim(SrcFile) = "" Then
    
        SrcFile = Replace(SrcFile, "/", Application.PathSeparator)
        SrcFile = Replace(SrcFile, "\", Application.PathSeparator)
        SrcFile = Replace(SrcFile, ":", Application.PathSeparator)
        ' SrcFileStart = InStrRev(SrcFile, Application.PathSeparator)
        ' SrcFile = Right(SrcFile, Len(SrcFile) - InStrRev(SrcFile, Application.PathSeparator))
        ' SrcFile = ActiveWorkbook.Path & Application.PathSeparator & SrcFile
        
        If AlreadyOpen(SrcFile) Then
            TradeSheetOpen = True
            Set SrcWb = Workbooks(SrcFile)
        Else
            TradeSheetOpen = False
            Set SrcWb = Workbooks.Open(SrcFile, True, True)
        End If
        
        YearArray = SrcWb.Worksheets(SrcSheet).Range("YearList2").Value
        YearReference = Worksheets("Summary").Range("M3").Value

        For iCnt = 1 To 10
        YearElement = YearArray(1, iCnt)
            If YearElement = YearReference Then
                YearOffset = iCnt
                iCnt = 10
            End If
        Next iCnt
            
        iCntOffset = 19     ' line offset on Summary page
        SrcCntOffset = 6   ' line offset on Asset register

        For iCnt = 1 To 10

            If Not SrcWb.Worksheets(SrcSheet).Range("F" & iCnt + SrcCntOffset).Value = "" Then
                Worksheets("Summary").Range("B" & (iCnt + iCntOffset)).Formula = SrcWb.Worksheets(SrcSheet).Range("F" & iCnt + SrcCntOffset).Value
                Worksheets("Summary").Range("C" & (iCnt + iCntOffset)).Formula = SrcWb.Worksheets(SrcSheet).Range("F" & iCnt + SrcCntOffset).Offset(0, YearOffset).Value
            End If
        
        Next iCnt
        
        ' CLOSE THE SOURCE FILE
        If TradeSheetOpen Then
                ' do nothing
        Else
            SrcWb.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
            Set SrcWb = Nothing
        End If
    End If
 
    '-----------------------------------------------------------------
    
    ' Get the supplier summaries from the Suppliers expense sheet.
    
    SrcSheet = "Suppliers"
    SrcFile = Worksheets("Summary").Range("L33").Value
    
    If Not Trim(SrcFile) = "" Then
    
        SrcFile = Replace(SrcFile, "/", Application.PathSeparator)
        SrcFile = Replace(SrcFile, "\", Application.PathSeparator)
        SrcFile = Replace(SrcFile, ":", Application.PathSeparator)
        ' SrcFileStart = InStrRev(SrcFile, Application.PathSeparator)
        ' SrcFile = Right(SrcFile, Len(SrcFile) - InStrRev(SrcFile, Application.PathSeparator))
        ' SrcFile = ActiveWorkbook.Path & Application.PathSeparator & SrcFile
        
        If AlreadyOpen(SrcFile) Then
            TradeSheetOpen = True
            Set SrcWb = Workbooks(SrcFile)
        Else
            TradeSheetOpen = False
            Set SrcWb = Workbooks.Open(SrcFile, True, True)
        End If
            
        iCntOffset = 32     ' line offset on Summary page
        SrcCntOffset = 6   ' line offset on Suppliers expense sheet

        For iCnt = 1 To 25

            If Not SrcWb.Worksheets(SrcSheet).Range("L" & iCnt + SrcCntOffset).Value = "" Then
                Worksheets("Summary").Range("B" & (iCnt + iCntOffset)).Formula = SrcWb.Worksheets(SrcSheet).Range("L" & iCnt + SrcCntOffset).Value
                Worksheets("Summary").Range("C" & (iCnt + iCntOffset)).Formula = SrcWb.Worksheets(SrcSheet).Range("M" & iCnt + SrcCntOffset).Value
                Worksheets("Summary").Range("E" & (iCnt + iCntOffset)).Formula = SrcWb.Worksheets(SrcSheet).Range("N" & iCnt + SrcCntOffset).Value
            End If
        
        Next iCnt
        
        ' CLOSE THE SOURCE FILE
        If TradeSheetOpen Then
                ' do nothing
        Else
            SrcWb.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
            Set SrcWb = Nothing
        End If
    End If
     
    '-----------------------------------------------------------------
            
    Worksheets("Summary").Range("H4").Formula = Now
    ActiveSheet.Protect Password:="lock"
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Function AlreadyOpen(sFname As String) As Boolean
    Dim wkb As Workbook
    On Error Resume Next
    Set wkb = Workbooks(sFname)
    AlreadyOpen = Not wkb Is Nothing
    Set wkb = Nothing
End Function
