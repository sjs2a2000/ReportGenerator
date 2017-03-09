Attribute VB_Name = "modSontag"
' ***************************************************************************************************************************************
'**                                                    Sontage Technical Indicators                                                     **
'**                                                                                                                                     **
'**                                           (C) Copyright 2016, Fullman Technologies, Inc.                                            **
'**                                                        All rights reserved.                                                         **
' ***************************************************************************************************************************************
Global dtTradeDate As Date
Type Historic
    Date        As String * 10
    Open        As Currency
    High        As Currency
    Low         As Currency
    Close       As Currency
    Volume      As Long
End Type
Global Hist(300) As Historic

Sub Main()

    'CheckForHoliday
    CreateDailyResults
    'CreateWeeklyResults
    'CreateSummary
    'SendReport "Open attachment", "Technicals", "SontagTechnicals", "c:\Users\scott\refdatavb6\Sontag\sontagtechnicals-2.xls"

End Sub
Sub CreateSummary()

    '-----------------------------------------------------------------------------------------------------------------------------
    ' This routine creates the summary page and writes the results for the short-term (daily) and intermediate-term (weekly)
    ' indexes to the historic table.
    '-----------------------------------------------------------------------------------------------------------------------------
    
    Dim sSymbol             As String * 15
    Dim dLastPrice          As Double
    Dim dNetChange          As Double
    Dim dPctChange          As Double
    Dim sSector             As String
    Dim lDailySignal        As Long
    Dim lWeeklySignal       As Long
    Dim lMonthlySignal      As Long
    Dim lCompSignal         As Long
    
    Dim ExcelApp            As Excel.Application
    Dim ExcelWorkBook       As Excel.Workbook
    Dim ExcelSheet          As Excel.Worksheet
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelWorkBook = ExcelApp.Workbooks.Open("c:\Users\scott\refdatavb6\sontag\sontagtechnicals-2.xls")
    Set ExcelSheet = ExcelWorkBook.Worksheets(2)
        
'MsgBox "got here 1"
        
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=Sontag", "", ""
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    
'MsgBox "got here 2"

    Dim rstStockList As ADODB.Recordset
    Dim cmdStockList As ADODB.Command
    Set rstStockList = New ADODB.Recordset
    Set cmdStockList = New ADODB.Command
    rstStockList.CursorType = adOpenDynamic
    rstStockList.LockType = adLockOptimistic
    Set cmdStockList.ActiveConnection = cnnl
    Excel.Application.Quit
    ExcelApp.Application.Quit
    
End Sub
Sub CreateWeeklyResults()

    '-----------------------------------------------------------------------------------------------------------------------------------
    ' This routine goes through the database of stocks and creates the data and signals for each and writes the results to the
    ' database table.
    '-----------------------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Dim sSymbol             As String * 15
    Dim dLastPrice          As Double
    Dim dNetChange          As Double
    Dim dPctChange          As Double
    Dim sSector             As String
    Dim dtEarnDate          As Date
    Dim dtDivDate           As Date
    Dim dOpen(300)          As Double
    Dim dHigh(300)          As Double
    Dim dLow(300)           As Double
    Dim dClose(300)         As Double
    Dim lVolume(300)        As Long
    Dim dDMA5               As Double
    Dim dDMA20              As Double
    Dim dDMA50              As Double
    Dim dDMA150             As Double
    Dim dDMA200             As Double
    Dim iDMASig             As Double
    Dim dMFVolume(500)      As Currency
    Dim dChaikenMF(500)     As Currency
    Dim sChaikenMF          As String
    Dim dChaikenOsc(500)    As Double
    Dim sChaikenOscSig      As String
    Dim dMyMACD(300)        As Double
    Dim dMyMACDSignal(300)  As Double
    Dim sMyMACDSignal       As String
    Dim sSignal             As String
    Dim dtSignalDate        As Date
    Dim lRow                As Long
    Dim lCols               As Long
    Dim lCounter            As Double
    Dim lPoints             As Long
    
    Dim ExcelApp            As Excel.Application
    Dim ExcelWorkBook       As Excel.Workbook
    Dim ExcelSheet          As Excel.Worksheet
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelWorkBook = ExcelApp.Workbooks.Open("c:\Users\scott\refdatavb6\sontag\sontagtechnicals-1.xls")
    Set ExcelSheet = ExcelWorkBook.Worksheets(2)
' MsgBox "got here 3"
 
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=Sontag", "", ""
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstOptions.Open "WeeklyTechnicals", cnnl, , , adCmdTable '
'MsgBox "got here 4"
    Dim rstStockList As ADODB.Recordset
    Dim cmdStockList As ADODB.Command
    Set rstStockList = New ADODB.Recordset
    Set cmdStockList = New ADODB.Command
    rstStockList.CursorType = adOpenDynamic
    rstStockList.LockType = adLockOptimistic
    Set cmdStockList.ActiveConnection = cnnl
    
    cmdStockList.CommandText = "Select * From StockList Where Symbol <>'zzz1234';"
    Set rstStockList = cmdStockList.Execute
    
    rstStockList.MoveFirst
    
    Do While Not rstStockList.EOF
    
        DoEvents
        
        sSymbol = rstStockList!Symbol
        dtDivDate = none
        dtExDate = ""
        lPoints = 0
   '     GetStockDates sSymbol, dtDivDate, dtEarnDate, sSector
        GetPublicPrices sSymbol, 2
        
        lCounter = 0
    
        Do
            DoEvents
            
            lCounter = lCounter + 1
            dOpen(lCounter) = Hist(lCounter).Open
            dHigh(lCounter) = Hist(lCounter).High
            dLow(lCounter) = Hist(lCounter).Low
            dClose(lCounter) = Hist(lCounter).Close
            lVolume(lCounter) = Hist(lCounter).Volume
            
            If lCounter = 200 Then Exit Do
        
        Loop
        
        dDMA5 = 0
        dDMA20 = 0
        dDMA50 = 0
        dDMA150 = 0
        dDMA200 = 0
        iDMASignal = 0
        
        dDMA5 = dSMA(1, 5)
        dDMA20 = dSMA(1, 20)
        dDMA50 = dSMA(1, 50)
        dDMA150 = dSMA(1, 150)
        dDMA200 = dSMA(1, 200)
        
        If Hist(1).Close > dDMA5 Then
            lPoints = lPoints + 0.25
            iDMASignal = iDMASignal + 0.25
        ElseIf Hist(1).Close < dDMA5 Then
            lPoints = lPoints - 0.25
            iDMASignal = iDMASignal - 0.25
        End If
        
        If Hist(1).Close > dDMA20 Then
            lPoints = lPoints + 0.75
            iDMASignal = iDMASignal + 0.75
        ElseIf Hist(1).Close < dDMA20 Then
            lPoints = lPoints - 0.75
            iDMASignal = iDMASignal - 0.75
        End If
        
        If Hist(1).Close > dDMA50 Then
            lPoints = lPoints + 1
            iDMASignal = iDMASignal + 1
        ElseIf Hist(1).Close < dDMA50 Then
            lPoints = lPoints - 1
            iDMASignal = iDMASignal - 1
        End If
        
        If Hist(1).Close > dDMA150 Then
            lPoints = lPoints + 1.25
            iDMASignal = iDMASignal + 1.25
        ElseIf Hist(1).Close < dDMA150 Then
            lPoints = lPoints - 1.25
            iDMASignal = iDMASignal - 1.25
        End If
        
        If Hist(1).Close > dDMA200 Then
            lPoints = lPoints + 1
            iDMASignal = iDMASignal + 1
        ElseIf Hist(1).Close < dDMA200 Then
            lPoints = lPoints - 1
            iDMASignal = iDMASignal - 1
        End If
        
        GetMACD dClose(), 12, 26, 9, dMyMACD(), dMyMACDSignal()
        
        If dMyMACD(1) > dMyMACDSignal(1) Then
            lPoints = lPoints + 2
            sMyMACDSignal = "Positive"
            If dMyMACD(2) < dMyMACDSignal(2) Then
                RecordSignalWeekly sSymbol, "MACD", "Positive"
            End If
        ElseIf dMyMACD(1) < dMyMACDSignal(1) Then
            lPoints = lPoints - 2
            sMyMACDSignal = "Negative"
            If dMyMACD(2) > dMyMACDSignal(2) Then
                RecordSignalWeekly sSymbol, "MACD", "Negative"
            End If
        End If
        
        ChaikenMoneyFlow dOpen(), dHigh(), dLow(), dClose(), lVolume(), dChaikenMF(), dMFVolume()
        ChaikenOscillator dOpen(), dHigh(), dLow(), dClose(), lVolume(), dChaikenOsc()

        If dChaikenOsc(1) > 0 Then
            sChaikenOscString = "Positive"
            lPoints = lPoints + 1.75
            If dChaikenOsc(2) < 0 Then
                RecordSignalWeekly sSymbol, "Chaiken Oscillator", "Positive"
            End If
        ElseIf dChaikenOsc(1) < 0 Then
            sChaikenOscString = "Negative"
            lPoints = lPoints - 1.75
            If dChaikenOsc(2) > 0 Then
                RecordSignalWeekly sSymbol, "Chaiken Oscillator", "Negative"
            End If
        End If
        
        If dChaikenMF(1) > 0 Then
            sChaikenMFString = "Positive"
            lPoints = lPoints + 1.5
            If dChaikenMF(2) < 0 Then
                RecordSignalWeekly sSymbol, "Chaiken Money Flow", "Positive"
            End If
        ElseIf dChaikenMF(1) < 0 Then
            sChaikenMFString = "Negative"
            lPoints = lPoints - 1.5
            If dChaikenMF(2) > 0 Then
                RecordSignalWeekly sSymbol, "Chaiken Money Flow", "Negative"
            End If
        End If
        
        rstOptions.AddNew
        rstOptions!SignalDate = Format(Date, "mm/dd/yyyy")
        rstOptions!Symbol = sSymbol
        rstOptions!LastPrice = Format(Hist(1).Close, "#,##0.00")
        rstOptions!NetChange = Format(Hist(1).Close - Hist(2).Close, "##0.00")
        rstOptions!TechIndex = lPoints
 '       rstOptions!Sector = sSector
 '       rstOptions!ExDate = dtDivDate
 '       rstOptions!EPSDate = dtEarnDate
        rstOptions!PctChange = Format(((Hist(1).Close - Hist(2).Close) / Hist(2).Close) * 100, "#,##0.0")
        rstOptions!DMA5 = Format(dDMA5, "##0.00")
        rstOptions!DMA20 = Format(dDMA20, "##0.00")
        rstOptions!DMA50 = Format(dDMA50, "##0.00")
        rstOptions!DMA150 = Format(dDMA150, "##0.00")
        rstOptions!DMA200 = Format(dDMA200, "##0.00")
        rstOptions!DMASignal = iDMASignal
        rstOptions!MACD = Format(dMyMACD(1), "##0.00")
        rstOptions!MACDSignal = Format(dMyMACDSignal(1), "##0.00")
        rstOptions!MACDSignalString = sMyMACDSignal
        rstOptions!ChaikenMoneyFlow = Format(dChaikenMF(1), "#00.00")
        rstOptions!ChaikenMoneyFlowString = sChaikenMFString
        rstOptions!ChaikenOsc = Format(dChaikenOsc(1), "#00.00")
        rstOptions!ChaikenOscString = sChaikenOscString
        rstOptions.Update
        
        rstStockList.MoveNext
        
    Loop
     
    cmdChange.CommandText = "Select * From WeeklyTechnicals Where SignalDate=#" & Format(Date, "mm/dd/yyyy") & "#;"
    Set rstOptions = cmdChange.Execute
    rstOptions.MoveFirst
    
    ExcelSheet.Cells(1, 2).Value = dtTradeDate
    
    lRow = 3
    
    Do While Not rstOptions.EOF
    
        DoEvents
        
        lRow = lRow + 1
        
        ExcelSheet.Cells(lRow, 1).Value = rstOptions!Symbol
        ExcelSheet.Cells(lRow, 3).Value = Format(rstOptions!LastPrice, "#,##0.00")
        ExcelSheet.Cells(lRow, 4).Value = Format(rstOptions!NetChange, "##0.00")
        ExcelSheet.Cells(lRow, 5).Value = Format(rstOptions!PctChange, "##0.0")
        ExcelSheet.Cells(lRow, 6).Value = Format(rstOptions!TechIndex, "##0.00")
        ExcelSheet.Cells(lRow, 7).Value = Format(rstOptions!DMASignal, "##.0")
        
'        If rstOptions!DMASignal > 2 Then
'            ExcelSheet.Cells(lRow, 7).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!DMASignal > -2 Then
'            ExcelSheet.Cells(lRow, 7).Interior.Color = RGB(255, 255, 0)
'        ElseIf rstOptions!DMASignal < -1 Then
'            ExcelSheet.Cells(lRow, 7).Interior.Color = RGB(255, 0, 0)
'        End If
        
'        If rstOptions!TechIndex > 9 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(113, 215, 69)
'        ElseIf rstOptions!TechIndex > 7 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(106, 202, 65)
'        ElseIf rstOptions!TechIndex > 5 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(11, 220, 11)
'        ElseIf rstOptions!TechIndex > 3 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(54, 236, 54)
'        ElseIf rstOptions!TechIndex > 0 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!TechIndex < -9 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(255, 0, 0)
'        ElseIf rstOptions!TechIndex < -7 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(200, 0, 0)
'        ElseIf rstOptions!TechIndex < -5 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(150, 0, 0)
'        ElseIf rstOptions!TechIndex < -3 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(125, 0, 0)
'        ElseIf rstOptions!TechIndex < 0 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(80, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(255, 255, 255)
'        End If
        
        GetWeeklySignal rstOptions!Symbol, "MACD", sSignal, dtSignalDate
        ExcelSheet.Cells(lRow, 8).Value = rstOptions!MACDSignalString
        If dtSignalDate <> #12:00:00 AM# Then
            ExcelSheet.Cells(lRow, 9).Value = dtSignalDate
        End If
'        If rstOptions!MACDSignalString = "Positive" Then
'            ExcelSheet.Cells(lRow, 8).Interior.Color = RGB(0, 255, 0)
'            ExcelSheet.Cells(lRow, 9).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!MACDSignalString = "Negative" Then
'            ExcelSheet.Cells(lRow, 8).Interior.Color = RGB(255, 0, 0)
'            ExcelSheet.Cells(lRow, 9).Interior.Color = RGB(255, 0, 0)
'        End If
        
        GetWeeklySignal rstOptions!Symbol, "Chaiken Oscillator", sSignal, dtSignalDate
        ExcelSheet.Cells(lRow, 10).Value = rstOptions!ChaikenOscString
        If dtSignalDate <> #12:00:00 AM# Then
            ExcelSheet.Cells(lRow, 11).Value = dtSignalDate
        End If
'        If rstOptions!ChaikenOscString = "Positive" Then
'            ExcelSheet.Cells(lRow, 10).Interior.Color = RGB(0, 255, 0)
'            ExcelSheet.Cells(lRow, 11).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!ChaikenOscString = "Negative" Then
'            ExcelSheet.Cells(lRow, 10).Interior.Color = RGB(255, 0, 0)
'            ExcelSheet.Cells(lRow, 11).Interior.Color = RGB(255, 0, 0)
'        End If
        
        GetWeeklySignal rstOptions!Symbol, "Chaiken Money Flow", sSignal, dtSignalDate
        ExcelSheet.Cells(lRow, 12).Value = rstOptions!ChaikenMoneyFlowString
        If dtSignalDate <> #12:00:00 AM# Then
            ExcelSheet.Cells(lRow, 13).Value = dtSignalDate
        End If
'        If rstOptions!ChaikenMoneyFlowString = "Positive" Then
'            ExcelSheet.Cells(lRow, 12).Interior.Color = RGB(0, 255, 0)
'            ExcelSheet.Cells(lRow, 13).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!ChaikenMoneyFlowString = "Negative" Then
'            ExcelSheet.Cells(lRow, 12).Interior.Color = RGB(255, 0, 0)
'            ExcelSheet.Cells(lRow, 13).Interior.Color = RGB(255, 0, 0)
'        End If
        
        'ExcelSheet.Cells(lRow, 2).Value = rstOptions!Sector
        'ExcelSheet.Cells(lRow, 14).Value = rstOptions!ExDate
       ' ExcelSheet.Cells(lRow, 15).Value = rstOptions!EPSDate
        
'        If rstOptions!ExDate < (Date - 1) + 5 Then
'            If rstOptions!ExDate > Date - 1 Then
'                ExcelSheet.Cells(lRow, 14).Interior.Color = RGB(255, 0, 0)
'            End If
'        ElseIf rstOptions!ExDate = Date - 1 Then
'            ExcelSheet.Cells(lRow, 14).Interior.Color = RGB(0, 255, 0)
'        End If
        
'        If rstOptions!EPSDate < (Date - 1) + 5 Then
'            If rstOptions!EPSDate > Date - 1 Then
'                ExcelSheet.Cells(lRow, 15).Interior.Color = RGB(255, 0, 0)
'            End If
'        ElseIf rstOptions!EPSDate = Date - 1 Then
'            ExcelSheet.Cells(lRow, 15).Interior.Color = RGB(0, 255, 0)
'        End If
            
    
        ExcelSheet.Cells(lRow, 17).Value = Format(rstOptions!DMA5, "#,##0.00")
        If rstOptions!LastPrice > rstOptions!DMA5 Then
            ExcelSheet.Cells(lRow, 17).Interior.Color = RGB(0, 255, 0)
        ElseIf rstOptions!LastPrice < rstOptions!DMA5 Then
            ExcelSheet.Cells(lRow, 17).Interior.Color = RGB(255, 0, 0)
        Else
            ExcelSheet.Cells(lRow, 17).Interior.Color = RGB(0, 0, 0)
        End If
        
        ExcelSheet.Cells(lRow, 18).Value = Format(rstOptions!DMA20, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA20 Then
'            ExcelSheet.Cells(lRow, 18).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA20 Then
'            ExcelSheet.Cells(lRow, 18).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 18).Interior.Color = RGB(0, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 19).Value = Format(rstOptions!DMA50, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA50 Then
'            ExcelSheet.Cells(lRow, 19).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA50 Then
'            ExcelSheet.Cells(lRow, 19).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 19).Interior.Color = RGB(0, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 20).Value = Format(rstOptions!DMA150, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA150 Then
'            ExcelSheet.Cells(lRow, 20).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA150 Then
'            ExcelSheet.Cells(lRow, 20).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 20).Interior.Color = RGB(0, 0, 0)
'        End If
                
        ExcelSheet.Cells(lRow, 21).Value = Format(rstOptions!DMA200, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA200 Then
'            ExcelSheet.Cells(lRow, 21).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA200 Then
'            ExcelSheet.Cells(lRow, 21).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 21).Interior.Color = RGB(0, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 22).Value = Format(rstOptions!ChaikenOsc, "#,##0.00")
'        If rstOptions!ChaikenOsc > 0 Then
'            ExcelSheet.Cells(lRow, 22).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!ChaikenOsc < 0 Then
'            ExcelSheet.Cells(lRow, 22).Interior.Color = RGB(255, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 23).Value = Format(rstOptions!ChaikenMoneyFlow, "#,##0.00")
'        If rstOptions!ChaikenMoneyFlow > 0 Then
'            ExcelSheet.Cells(lRow, 23).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!ChaikenMoneyFlow < 0 Then
'            ExcelSheet.Cells(lRow, 23).Interior.Color = RGB(255, 0, 0)
'        End If
    
        rstOptions.MoveNext
        
    Loop
    
    ExcelSheet.Range("b3:b" & lRow).NumberFormat = "#,##0.00"
    ExcelSheet.Range("c3:c" & lRow).NumberFormat = "##0.00"
    ExcelSheet.Range("d3:d" & lRow).NumberFormat = "##0.0"
    ExcelSheet.Range("q3:u" & lRow).NumberFormat = "#,##0.00"
    ExcelSheet.Range("n3:o" & lRow).NumberFormat = "MM/DD/YYYY"
    
    
    ExcelSheet.EnableAutoFilter = True
    ExcelWorkBook.SaveCopyAs ("c:\Users\scott\refdatavb6\sontag\sontagtechnicals-2.xls")
    ExcelWorkBook.Close savechanges = False
    Excel.Application.Quit
    ExcelApp.Application.Quit
    Set ExcelApp = Nothing
    Set ExcelWorkBook = Nothing
    Set excelworksheet = Nothing
'    ExcelWorkBook.Close savechanges = False
    Set ExcelApp = Nothing
    Set ExcelWorkBook = Nothing
    Set excelworksheet = Nothing
    
End Sub
Sub CreateDailyResults()

    '-----------------------------------------------------------------------------------------------------------------------------------
    ' This routine goes through the database of stocks and creates the data and signals for each and writes the results to the
    ' database table.
    '-----------------------------------------------------------------------------------------------------------------------------------
    
  '  On Error Resume Next
    On Error GoTo ShowError:
    
    Dim sSymbol             As String * 15
    Dim dLastPrice          As Double
    Dim dNetChange          As Double
    Dim dPctChange          As Double
    Dim sSector             As String
    Dim dtEarnDate          As Date
    Dim dtDivDate           As Date
    Dim dOpen(300)          As Double
    Dim dHigh(300)          As Double
    Dim dLow(300)           As Double
    Dim dClose(300)         As Double
    Dim lVolume(300)        As Long
    Dim dDMA5               As Double
    Dim dDMA20              As Double
    Dim dDMA50              As Double
    Dim dDMA150             As Double
    Dim dDMA200             As Double
    Dim iDMASig             As Double
    Dim dMFVolume(500)      As Currency
    Dim dChaikenMF(500)     As Currency
    Dim sChaikenMF          As String
    Dim dChaikenOsc(500)    As Double
    Dim sChaikenOscSig      As String
    Dim dMyMACD(300)        As Double
    Dim dMyMACDSignal(300)  As Double
    Dim sMyMACDSignal       As String
    Dim sSignal             As String
    Dim dtSignalDate        As Date
    Dim lRow                As Long
    Dim lCols               As Long
    Dim lCounter            As Long
    Dim lPoints             As Double
    Dim iDataCount          As Integer
    
    Dim ExcelApp            As Excel.Application
    Dim ExcelWorkBook       As Excel.Workbook
    Dim ExcelSheet          As Excel.Worksheet
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelWorkBook = ExcelApp.Workbooks.Open("c:\Users\scott\refdatavb6\sontag\sontagtechnicals.xls")
    Set ExcelSheet = ExcelWorkBook.Worksheets(1)
 
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=Sontag", "", ""
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstOptions.Open "DailyTechnicals", cnnl, , , adCmdTable

    Dim rstStockList As ADODB.Recordset
    Dim cmdStockList As ADODB.Command
    Set rstStockList = New ADODB.Recordset
    Set cmdStockList = New ADODB.Command
    rstStockList.CursorType = adOpenDynamic
    rstStockList.LockType = adLockOptimistic
    Set cmdStockList.ActiveConnection = cnnl
    
    'ExcelSheet.Range("b3:zz500").Delete
    
    cmdStockList.CommandText = "Select * From StockList Where Symbol <>'zzz1234';"
    Set rstStockList = cmdStockList.Execute
    
    rstStockList.MoveFirst
    
    Do While Not rstStockList.EOF
    
        DoEvents
        
        sSymbol = Trim(rstStockList!Symbol)
Debug.Print sSymbol

        dtDivDate = none
        dtExDate = ""
        lPoints = 0
        'GetStockDates sSymbol, dtDivDate, dtEarnDate, sSector
        iDataCount = 0
        Hist(1).Close = 0
        Do
            GetPublicPrices sSymbol, 1
            If Hist(1).Close <> 0 Then Exit Do
            iDataCount = iDataCount + 1
            If iDataCount > 5 Then Exit Do
        Loop
        
        lCounter = 0
    
        Do
            DoEvents
            
            lCounter = lCounter + 1
            dOpen(lCounter) = Hist(lCounter).Open
            dHigh(lCounter) = Hist(lCounter).High
            dLow(lCounter) = Hist(lCounter).Low
            dClose(lCounter) = Hist(lCounter).Close
            lVolume(lCounter) = Hist(lCounter).Volume
            
            If lCounter = 200 Then Exit Do
        
        Loop
        
        dDMA5 = 0
        dDMA20 = 0
        dDMA50 = 0
        dDMA150 = 0
        dDMA200 = 0
        iDMASignal = 0
        
        dDMA5 = dSMA(1, 5)
        dDMA20 = dSMA(1, 20)
        dDMA50 = dSMA(1, 50)
        dDMA150 = dSMA(1, 150)
        dDMA200 = dSMA(1, 200)
        
        If Hist(1).Close > dDMA5 Then
            lPoints = lPoints + 0.25
            iDMASignal = iDMASignal + 0.25
        ElseIf Hist(1).Close < dDMA5 Then
            lPoints = lPoints - 0.25
            iDMASignal = iDMASignal - 0.25
        End If
        
        If Hist(1).Close > dDMA20 Then
            lPoints = lPoints + 0.75
            iDMASignal = iDMASignal + 0.75
        ElseIf Hist(1).Close < dDMA20 Then
            lPoints = lPoints - 0.75
            iDMASignal = iDMASignal - 0.75
        End If
        
        If Hist(1).Close > dDMA50 Then
            lPoints = lPoints + 1
            iDMASignal = iDMASignal + 1
        ElseIf Hist(1).Close < dDMA50 Then
            lPoints = lPoints - 1
            iDMASignal = iDMASignal - 1
        End If
        
        If Hist(1).Close > dDMA150 Then
            lPoints = lPoints + 1.25
            iDMASignal = iDMASignal + 1.25
        ElseIf Hist(1).Close < dDMA150 Then
            lPoints = lPoints - 1.25
            iDMASignal = iDMASignal - 1.25
        End If
        
        If Hist(1).Close > dDMA200 Then
            lPoints = lPoints + 1.75
            iDMASignal = iDMASignal + 1.75
        ElseIf Hist(1).Close < dDMA200 Then
            lPoints = lPoints - 1.75
            iDMASignal = iDMASignal - 1.75
        End If
        
        GetMACD dClose(), 4, 12, 2, dMyMACD(), dMyMACDSignal()
        
        If dMyMACD(1) > dMyMACDSignal(1) Then
            lPoints = lPoints + 1.9
            sMyMACDSignal = "Positive"
            If dMyMACD(2) < dMyMACDSignal(2) Then
                RecordSignal sSymbol, "MACD", "Positive"
            End If
        ElseIf dMyMACD(1) < dMyMACDSignal(1) Then
            lPoints = lPoints - 1.9
            sMyMACDSignal = "Negative"
            If dMyMACD(2) > dMyMACDSignal(2) Then
                RecordSignal sSymbol, "MACD", "Negative"
            End If
        End If
        
        ChaikenMoneyFlow dOpen(), dHigh(), dLow(), dClose(), lVolume(), dChaikenMF(), dMFVolume()
        ChaikenOscillator dOpen(), dHigh(), dLow(), dClose(), lVolume(), dChaikenOsc()

        If dChaikenOsc(1) > 0 Then
            sChaikenOscString = "Positive"
            lPoints = lPoints + 1.8
            If dChaikenOsc(2) < 0 Then
                RecordSignal sSymbol, "Chaiken Oscillator", "Positive"
            End If
        ElseIf dChaikenOsc(1) < 0 Then
            sChaikenOscString = "Negative"
            lPoints = lPoints - 1.8
            If dChaikenOsc(2) > 0 Then
                RecordSignal sSymbol, "Chaiken Oscillator", "Negative"
            End If
        End If
        
        If dChaikenMF(1) > 0 Then
            sChaikenMFString = "Positive"
            lPoints = lPoints + 1.9
            If dChaikenMF(2) < 0 Then
                RecordSignal sSymbol, "Chaiken Money Flow", "Positive"
            End If
        ElseIf dChaikenMF(1) < 0 Then
            sChaikenMFString = "Negative"
            lPoints = lPoints - 1.9
            If dChaikenMF(2) > 0 Then
                RecordSignal sSymbol, "Chaiken Money Flow", "Negative"
            End If
        End If
        
        rstOptions.AddNew
        rstOptions!SignalDate = Format(Date, "mm/dd/yyyy")
        rstOptions!Symbol = sSymbol
        rstOptions!LastPrice = Format(Hist(1).Close, "#,##0.00")
        rstOptions!NetChange = Format(Hist(1).Close - Hist(2).Close, "##0.00")
        rstOptions!TechIndex = lPoints
        'rstOptions!Sector = sSector
        'rstOptions!ExDate = dtDivDate
        'rstOptions!EPSDate = dtEarnDate
        rstOptions!PctChange = Format(((Hist(1).Close - Hist(2).Close) / Hist(2).Close) * 100, "#,##0.0")
        rstOptions!DMA5 = Format(dDMA5, "##0.00")
        rstOptions!DMA20 = Format(dDMA20, "##0.00")
        rstOptions!DMA50 = Format(dDMA50, "##0.00")
        rstOptions!DMA150 = Format(dDMA150, "##0.00")
        rstOptions!DMA200 = Format(dDMA200, "##0.00")
        rstOptions!DMASignal = iDMASignal
        rstOptions!MACD = Format(dMyMACD(1), "##0.00")
        rstOptions!MACDSignal = Format(dMyMACDSignal(1), "##0.00")
        rstOptions!MACDSignalString = sMyMACDSignal
        rstOptions!ChaikenMoneyFlow = Format(dChaikenMF(1), "#00.00")
        rstOptions!ChaikenMoneyFlowString = sChaikenMFString
        rstOptions!ChaikenOsc = Format(dChaikenOsc(1), "#00.00")
        rstOptions!ChaikenOscString = sChaikenOscString
        rstOptions.Update
        
        rstStockList.MoveNext
        
    Loop
     
    cmdChange.CommandText = "Select * From DailyTechnicals Where SignalDate=#" & Format(Date, "mm/dd/yyyy") & "#;"
    Set rstOptions = cmdChange.Execute
    rstOptions.MoveFirst
    
    dtTradeDate = Hist(1).Date
    ExcelSheet.Cells(1, 2).Value = Date$
    
    lRow = 3
    
    Do While Not rstOptions.EOF
    
        DoEvents
        
        lRow = lRow + 1
        
        ExcelSheet.Cells(lRow, 1).Value = rstOptions!Symbol
        ExcelSheet.Cells(lRow, 3).Value = Format(rstOptions!LastPrice, "#,##0.00")
        ExcelSheet.Cells(lRow, 4).Value = Format(rstOptions!NetChange, "##0.00")
        ExcelSheet.Cells(lRow, 5).Value = Format(rstOptions!PctChange, "##0.0")
        ExcelSheet.Cells(lRow, 6).Value = Format(rstOptions!TechIndex, "##0.00")
        ExcelSheet.Cells(lRow, 7).Value = Format(rstOptions!DMASignal, "#0.00")
        
'        If rstOptions!DMASignal > 2 Then
'            ExcelSheet.Cells(lRow, 7).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!DMASignal > -2 Then
'            ExcelSheet.Cells(lRow, 7).Interior.Color = RGB(255, 255, 0)
'        ElseIf rstOptions!DMASignal < -1 Then
'            ExcelSheet.Cells(lRow, 7).Interior.Color = RGB(255, 0, 0)
'        End If
        
'        If rstOptions!TechIndex > 9 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(113, 215, 69)
'        ElseIf rstOptions!TechIndex > 7 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(106, 202, 65)
'        ElseIf rstOptions!TechIndex > 5 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(11, 220, 11)
'        ElseIf rstOptions!TechIndex > 3 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(54, 236, 54)
'        ElseIf rstOptions!TechIndex > 0 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!TechIndex < -9 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(255, 0, 0)
'        ElseIf rstOptions!TechIndex < -7 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(200, 0, 0)
'        ElseIf rstOptions!TechIndex < -5 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(150, 0, 0)
'        ElseIf rstOptions!TechIndex < -3 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(125, 0, 0)
'        ElseIf rstOptions!TechIndex < 0 Then
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(80, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 6).Interior.Color = RGB(255, 255, 255)
'        End If
        
        GetSignal rstOptions!Symbol, "MACD", sSignal, dtSignalDate
        ExcelSheet.Cells(lRow, 8).Value = rstOptions!MACDSignalString
        If dtSignalDate <> #12:00:00 AM# Then
            ExcelSheet.Cells(lRow, 9).Value = dtSignalDate
        End If
'        If rstOptions!MACDSignalString = "Positive" Then
'            ExcelSheet.Cells(lRow, 8).Interior.Color = RGB(0, 255, 0)
'            ExcelSheet.Cells(lRow, 9).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!MACDSignalString = "Negative" Then
'            ExcelSheet.Cells(lRow, 8).Interior.Color = RGB(255, 0, 0)
'            ExcelSheet.Cells(lRow, 9).Interior.Color = RGB(255, 0, 0)
'        End If
        
        GetSignal rstOptions!Symbol, "Chaiken Oscillator", sSignal, dtSignalDate
        ExcelSheet.Cells(lRow, 10).Value = rstOptions!ChaikenOscString
        If dtSignalDate <> #12:00:00 AM# Then
            ExcelSheet.Cells(lRow, 11).Value = dtSignalDate
        End If
'        If rstOptions!ChaikenOscString = "Positive" Then
'            ExcelSheet.Cells(lRow, 10).Interior.Color = RGB(0, 255, 0)
'            ExcelSheet.Cells(lRow, 11).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!ChaikenOscString = "Negative" Then
'            ExcelSheet.Cells(lRow, 10).Interior.Color = RGB(255, 0, 0)
'            ExcelSheet.Cells(lRow, 11).Interior.Color = RGB(255, 0, 0)
'        End If
'
'        GetSignal rstOptions!Symbol, "Chaiken Money Flow", sSignal, dtSignalDate
'        ExcelSheet.Cells(lRow, 12).Value = rstOptions!ChaikenMoneyFlowString
'        If dtSignalDate <> #12:00:00 AM# Then
'            ExcelSheet.Cells(lRow, 13).Value = dtSignalDate
'        End If
'        If rstOptions!ChaikenMoneyFlowString = "Positive" Then
'            ExcelSheet.Cells(lRow, 12).Interior.Color = RGB(0, 255, 0)
'            ExcelSheet.Cells(lRow, 13).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!ChaikenMoneyFlowString = "Negative" Then
'            ExcelSheet.Cells(lRow, 12).Interior.Color = RGB(255, 0, 0)
'            ExcelSheet.Cells(lRow, 13).Interior.Color = RGB(255, 0, 0)
'        End If
        
        'ExcelSheet.Cells(lRow, 2).Value = rstOptions!Sector
        'ExcelSheet.Cells(lRow, 14).Value = rstOptions!ExDate
        'ExcelSheet.Cells(lRow, 15).Value = rstOptions!EPSDate
        
'        If rstOptions!ExDate < (Date - 1) + 5 Then
'            If rstOptions!ExDate > Date - 1 Then
'                ExcelSheet.Cells(lRow, 14).Interior.Color = RGB(255, 0, 0)
'            End If
'        ElseIf rstOptions!ExDate = Date - 1 Then
'            ExcelSheet.Cells(lRow, 14).Interior.Color = RGB(0, 255, 0)
'        End If
'
'        If rstOptions!EPSDate < (Date - 1) + 5 Then
'            If rstOptions!EPSDate > Date - 1 Then
'                ExcelSheet.Cells(lRow, 15).Interior.Color = RGB(255, 0, 0)
'            End If
'        ElseIf rstOptions!EPSDate = Date - 1 Then
'            ExcelSheet.Cells(lRow, 15).Interior.Color = RGB(0, 255, 0)
'        End If
            
    
        ExcelSheet.Cells(lRow, 17).Value = Format(rstOptions!DMA5, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA5 Then
'            ExcelSheet.Cells(lRow, 17).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA5 Then
'            ExcelSheet.Cells(lRow, 17).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 17).Interior.Color = RGB(0, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 18).Value = Format(rstOptions!DMA20, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA20 Then
'            ExcelSheet.Cells(lRow, 18).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA20 Then
'            ExcelSheet.Cells(lRow, 18).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 18).Interior.Color = RGB(0, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 19).Value = Format(rstOptions!DMA50, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA50 Then
'            ExcelSheet.Cells(lRow, 19).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA50 Then
'            ExcelSheet.Cells(lRow, 19).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 19).Interior.Color = RGB(0, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 20).Value = Format(rstOptions!DMA150, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA150 Then
'            ExcelSheet.Cells(lRow, 20).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA150 Then
'            ExcelSheet.Cells(lRow, 20).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 20).Interior.Color = RGB(0, 0, 0)
'        End If
                
        ExcelSheet.Cells(lRow, 21).Value = Format(rstOptions!DMA200, "#,##0.00")
'        If rstOptions!LastPrice > rstOptions!DMA200 Then
'            ExcelSheet.Cells(lRow, 21).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!LastPrice < rstOptions!DMA200 Then
'            ExcelSheet.Cells(lRow, 21).Interior.Color = RGB(255, 0, 0)
'        Else
'            ExcelSheet.Cells(lRow, 21).Interior.Color = RGB(0, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 22).Value = Format(rstOptions!ChaikenOsc, "#,##0.00")
'        If rstOptions!ChaikenOsc > 0 Then
'            ExcelSheet.Cells(lRow, 22).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!ChaikenOsc < 0 Then
'            ExcelSheet.Cells(lRow, 22).Interior.Color = RGB(255, 0, 0)
'        End If
        
        ExcelSheet.Cells(lRow, 23).Value = Format(rstOptions!ChaikenMoneyFlow, "#,##0.00")
'        If rstOptions!ChaikenMoneyFlow > 0 Then
'            ExcelSheet.Cells(lRow, 23).Interior.Color = RGB(0, 255, 0)
'        ElseIf rstOptions!ChaikenMoneyFlow < 0 Then
'            ExcelSheet.Cells(lRow, 23).Interior.Color = RGB(255, 0, 0)
'        End If
    
        rstOptions.MoveNext
        
    Loop
    
    ExcelSheet.Range("b3:b" & lRow).NumberFormat = "#,##0.00"
    ExcelSheet.Range("c3:c" & lRow).NumberFormat = "##0.00"
    ExcelSheet.Range("d3:d" & lRow).NumberFormat = "##0.0"
    ExcelSheet.Range("q3:u" & lRow).NumberFormat = "#,##0.00"
    ExcelSheet.Range("n3:o" & lRow).NumberFormat = "MM/DD/YYYY"
    
    
    ExcelSheet.EnableAutoFilter = True
    ExcelWorkBook.SaveCopyAs ("c:\Users\scott\refdatavb6\sontag\sontagtechnicals-1.xls")
    ExcelWorkBook.Close savechanges = False
    Excel.Application.Quit
    ExcelApp.Application.Quit
    Set ExcelApp = Nothing
    Set ExcelWorkBook = Nothing
    Set excelworksheet = Nothing
'    ExcelWorkBook.Close savechanges = False
    Set ExcelApp = Nothing
    Set ExcelWorkBook = Nothing
    Set excelworksheet = Nothing
    
    Exit Sub
    
ShowError:
    If Err = 6 Then Resume Next
    MsgBox Err & " " & Error$ & " " & sSymbol
    Resume Next
    
End Sub
Sub RecordSignal(sSymbol As String, sIndicator As String, sSignal As String)
'MsgBox "got here 7"
    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=Sontag", "", ""
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstOptions.Open "SignalTracker", cnnl, , , adCmdTable
'MsgBox "got here 8"
    rstOptions.AddNew
    rstOptions!DateofSignal = Format(Date$, "mm/dd/yyyy")
    rstOptions!TimeofSignal = Format(Time$, "hh:mm:ss")
    rstOptions!Symbol = sSymbol
    rstOptions!Indicator = sIndicator
    rstOptions!SignalString = sSignal
    rstOptions.Update
    
    rstOptions.Close
    cnnl.Close
    
End Sub
Sub RecordSignalWeekly(sSymbol As String, sIndicator As String, sSignal As String)

    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=Sontag", "", ""
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstOptions.Open "SignalTracker", cnnl, , , adCmdTable
    
    rstOptions.AddNew
    rstOptions!DateofSignal = Format(Date$, "mm/dd/yyyy")
    rstOptions!TimeofSignal = Format(Time$, "hh:mm:ss")
    rstOptions!Symbol = sSymbol
    rstOptions!Indicator = sIndicator
    rstOptions!SignalString = sSignal
    rstOptions.Update
    
    rstOptions.Close
    cnnl.Close
    

End Sub
Sub GetSignal(sSymbol As String, sIndicator As String, sSignal As String, dtSignal As Date)

    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=Sontag", "", ""
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstOptions.Open "SignalTracker", cnnl, , , adCmdTable
    
    cmdChange.CommandText = "Select * From SignalTracker Where Symbol='" & sSymbol & "' and Indicator='" & sIndicator & "' Order By DateofSignal*-1;"
    Set rstOptions = cmdChange.Execute
    
    If rstOptions.BOF Then
        'Do nothing
    Else
        rstOptions.MoveFirst
        sSignal = rstOptions!SignalString
        dtSignal = rstOptions!DateofSignal
    End If
    
    rstOptions.Close
    cnnl.Close
    
End Sub
Sub GetWeeklySignal(sSymbol As String, sIndicator As String, sSignal As String, dtSignal As Date)

    Dim cnnl As New ADODB.Connection
    Dim rstOptions As ADODB.Recordset
    Set cnnl = New ADODB.Connection
    Dim cmdChange       As ADODB.Command
    cnnl.Open "DSN=Sontag", "", ""
    cnnl.CommandTimeout = 100
    Set rstOptions = New ADODB.Recordset
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnl
    rstOptions.CursorType = adOpenDynamic
    rstOptions.LockType = adLockOptimistic
    rstOptions.Open "SignalTracker", cnnl, , , adCmdTable
    
    cmdChange.CommandText = "Select * From SignalTracker Where Symbol='" & sSymbol & "' and Indicator='" & sIndicator & "' Order By DateofSignal*-1;"
    Set rstOptions = cmdChange.Execute
    
    If rstOptions.BOF Then
        'Do nothing
    Else
        rstOptions.MoveFirst
        sSignal = rstOptions!SignalString
        dtSignal = rstOptions!DateofSignal
    End If
    
    rstOptions.Close
    cnnl.Close

End Sub
Sub GetMACD(ByRef dPrices() As Double, ByVal iEMALen1 As Integer, iEMALen2 As Integer, iSigLen As Integer, dMACD() As Double, dMACDSignal() As Double)
    '----------------------------------------------------------------------------------
    ' This process calculates two EMA values and then the MACD and the MACD signal.
    '----------------------------------------------------------------------------------
    
    On Error Resume Next
    
     Dim dEMA1(300)      As Double
    Dim dEMA2(300)      As Double
    Dim dSMA1           As Double
    Dim dSMA2           As Double
    Dim dSMA3           As Double
    Dim iCounter1       As Integer
    Dim iCounter2       As Integer
    Dim iStartFrom      As Integer
    
    ' Calculate the two EMAs
    
    dSMA1 = 0
    
    For iCounter1 = 200 To 200 + iEMALen1
        dSMA1 = dSMA1 + dPrices(iCounter1)
    Next iCounter1
    
    dSMA1 = dSMA1 / iEMALen1
    dEMA1(199) = dPrices(199) * (2 / (iEMALen1 + 1)) + dSMA1 * (1 - (2 / (iEMALen1 + 1)))
    
    For iCounter1 = 198 To 0 Step -1
        dEMA1(iCounter1) = dPrices(iCounter1) * (2 / (iEMALen1 + 1)) + dEMA1(iCounter1 + 1) * (1 - (2 / (iEMALen1 + 1)))
    Next iCounter1
    
    dSMA2 = 0
    
    For iCounter2 = 200 To 200 + iEMALen2
        dSMA2 = dSMA2 + dPrices(iCounter2)
    Next iCounter2
    
    dEMA2(199) = dPrices(199) * (2 / (iEMALen2 + 1)) + dSMA2 * (1 - (2 / (iEMALen2 + 1)))
    
    For iCounter2 = 198 To 0 Step -1
        dEMA2(iCounter2) = dPrices(iCounter2) * (2 / (iEMALen2 + 1)) + dEMA2(iCounter2 + 1) * (1 - (2 / (iEMALen2 + 1)))
    Next iCounter2
    
    For iCounter1 = 0 To 198
        dMACD(iCounter1) = dEMA1(iCounter1) - dEMA2(iCounter1)
    Next iCounter1
    
    dSMA3 = 0
    
    For iCounter1 = 200 To 200 + iSigLen
        dSMA3 = dSMA3 + dMACD(iCounter1)
    Next iCounter1
    
    dSMA3 = dSMA3 / iSigLen
    dMACDSignal(197) = dMACD(197) * (2 / (iSigLen + 1)) + dSMA3 * (1 - (2 / (iSigLen + 1)))
    
    For iCounter1 = 196 To 0 Step -1
        dMACDSignal(iCounter1) = dMACD(iCounter1) * (2 / (iSigLen + 1)) + dMACDSignal(iCounter1 + 1) * (1 - (2 / (iSigLen + 1)))
    Next iCounter1
    
End Sub
'Sub GetStockDates(sStockSymbol As String, dtMyExDate As Date, dtMyEarnDate As Date, sMySector As String)
'
'    '---------------------------------------------------------------------------------
'    ' This routine gets the ex-dividend and earnings dates
'    '---------------------------------------------------------------------------------
'
'    On Error Resume Next
'
'    Dim cnnl As ADODB.Connection
'    Dim rstOptions As ADODB.Recordset
'    Set cnnl = New ADODB.Connection
'    Dim rst2 As ADODB.Recordset
'    cnnl.Open "DSN=UAM"
'    cnnl.CommandTimeout = 100
'    Dim cmdChange       As ADODB.Command
'    Dim cmdChange2      As ADODB.Command
'    Set rstOptions = New ADODB.Recordset
'    Set rst2 = New ADODB.Recordset
'    Set cmdChange = New ADODB.Command
'    Set cmdChange.ActiveConnection = cnnl
'    Set cmdChange2 = New ADODB.Command
'    Set cmdChange2.ActiveConnection = cnnl
'    rst2.CursorType = adOpenDynamic
'    rst2.LockType = adLockOptimistic
'    rstOptions.Open "NewSectors", cnnl, adOpenDynamic, adLockOptimistic
'
'    cmdChange.CommandText = "Select * From UnderlyingSecurity Where StockSymbol ='" & Trim(sStockSymbol) & "';"
'    Set rstOptions = cmdChange.Execute
'    rstOptions.MoveFirst
'
'    sMySector = rstOptions!Sector
'    dtMyEarnDate = rstOptions!EarningsDate
'
'    cmdChange2.CommandText = "Select * From StockTable Where UnderlyingSymbol='" & Trim(sStockSymbol) & "';"
'    Set rst2 = cmdChange2.Execute
'    rst2.MoveFirst
'
'    dtMyExDate = rst2!DivDate
'    rst2.Close
'    rstOptions.Close
'    cnnl.Close
'
'
'End Sub
Sub CreateDailySpreadsheet()

End Sub
Sub ChaikenMoneyFlow(dOpen() As Double, dHigh() As Double, dLow() As Double, dClose() As Double, lVolume() As Long, dMFInd() As Currency, dMFVolume() As Currency)

    '----------------------------------------------------------------------------------
    ' This formula creates the ChaikenMoneyFlow Reading for Day N.
    ' Added March 5, 2016 at Fullman Technologies Inc. by SHF.
    '----------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Dim dMF             As Double
    Dim dMFV            As Double
    Dim lCount          As Long
    
    lCount = 501
    
    Do
        
        DoEvents
        
        lCount = lCount - 1
        
        dMF = 0
        dMFV = 0
        
        dMF = ((dClose(lCount) - dLow(lCount)) - (dHigh(lCount) - dClose(lCount))) / (dHigh(lCount) - dLow(lCount))
        dMFV = dMF * lVolume(lCount)
        
        dMFInd(lCount) = dMF
        dMFVolume(lCount) = dMFV
        If lCount < 2 Then Exit Do
    
    Loop
    
End Sub
Sub ChaikenOscillator(dOpen() As Double, dHigh() As Double, dLow() As Double, dClose() As Double, lVolume() As Long, dOscillator() As Double)

    '----------------------------------------------------------------------------------
    ' This formula creates the ChaikenOscillator Reading for Day N.
    ' Added March 5, 2016 at Fullman Technologies Inc. by SHF.
    '----------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Dim dADLine(500)            As Currency
    Dim dEMA3(500)              As Currency
    Dim dEMA10(500)             As Currency
    Dim dMFInd(500)             As Currency
    Dim dMFVolume(500)          As Currency
    Dim dOscill                 As Currency
    Dim lCount                  As Long
    Dim lCount3                 As Long
    Dim lCount10                As Long
    
    ChaikenMoneyFlow dOpen(), dHigh(), dLow(), dClose(), lVolume(), dMFInd(), dMFVolume()
    
    lCount = 501
    
    Do
        
        DoEvents
        
        lCount = lCount - 1
        
        dADLine(lCount) = dADLine(lCount + 1) + dMFVolume(lCount)
Debug.Print dADLine(lCount), dMFVolume(lCount)
        If lCount < 497 Then
            
            dEMA3(lCount) = dADLine(lCount) * (2 / (4)) + dEMA3(lCount + 1) * (1 - (2 / (4)))
        
        ElseIf lCount = 498 Then
            
            dEMA3(498) = (dADLine(500) + dADLine(499) + dADLine(498)) / 3
            
        End If
        
        If lCount < 491 Then
        
            dEMA10(lCount) = dADLine(lCount) * (2 / (11)) + dEMA10(lCount + 1) * (1 - (2 / (11)))
            
            dOscillator(lCount) = dEMA3(lCount) - dEMA10(lCount)
        
        ElseIf lCount = 491 Then
            
            dEMA10(490) = (dADLine(500) + dADLine(499) + dADLine(498) + dADLine(497) + dADLine(496) + dADLine(495) + dADLine(494) + dADLine(493) + dADLine(492) + dADLine(491)) / 10
        
        End If
        
        If lCount < 1 Then Exit Do
        
    Loop
Debug.Print dOscillator(1)
'Stop
        
End Sub
Function dSMA(lStart As Long, lPeriods As Long) As Double

    Dim lCount As Long
    dSMA = 0
    
    lCount = -1
    lCount = 0
    
    Do
        
        DoEvents
        lCount = lCount + 1
        dSMA = dSMA + Hist(lCount + lStart).Close
'Debug.Print Hist(lCount).Close, lPeriods, dSMA, lCount
        If lCount = (lPeriods) Then Exit Do

    Loop
Debug.Print Hist(1).Close, dSMA, lPeriods, dSMA / lPeriods, lCount
    dSMA = dSMA / lPeriods
'Stop
    
End Function

Sub GetPublicPrices(sSymbol As String, iPeriod As Integer)

    '---------------------------------------------------------------------------------------------------------------------------------
    ' This routine retrieves data that was collected from the public domain.
    '---------------------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
   

    Dim sPath               As String
    Dim sPeriod(3)          As String
    Dim sDate               As String
    Dim dOpen               As Double
    Dim dHigh               As Double
    Dim dLow                As Double
    Dim dClose              As Double
    Dim lVolumeInput        As Long
    Dim dAdjClose           As Double

    sPath = "c:\prices\"
 '
    sPeriod(1) = "Daily"
    sPeriod(2) = "Weekly"
    sPeriod(3) = "Monthly"
    
    sSymbol = Trim(sSymbol)

    iFile = FreeFile
   ' Open sPath & sPeriod(iPeriod) & "\" & sSymbol & ".txt" For Input As iFile
    Open sPath & sPeriod(iPeriod) & "\" & sSymbol & ".txt" For Input As iFile
        sMyDummy = Input$(LOF(iFile), #iFile)
    Close iFile

    iMyLines = Split(sMyDummy, vbLf)

    num_rows = UBound(iMyLines)
    oneline = Split(iMyLines(0), ",")
    num_cols = UBound(oneline)
    online = Split(iMyLines(0), ",")
    
    lCounter = 0
    
    Do
    
        DoEvents

        lCounter = lCounter + 1
        online = Split(iMyLines(lCounter), ",")
        sDate = online(0)
        dOpen = online(1)
        dHigh = online(2)
        dLow = online(3)
        dClose = online(4)
        lVolumeInput = online(5)
        dAdjClose = online(6)
        
        Hist(lCounter).Date = Format(sDate, "m/d/yyyy")
        Hist(lCounter).Open = dOpen
        Hist(lCounter).High = dHigh
        Hist(lCounter).Low = dLow
        Hist(lCounter).Close = dAdjClose
        Hist(lCounter).Close = dClose
        Hist(lCounter).Volume = lVolumeInput
        If lCounter = 5000 Then Exit Do
        If lCounter >= num_rows - 1 Then Exit Do
      '  If lCounter > MaxSize - 1 Then Exit Do
        
    '    Debug.Print lCounter, Hist(lCounter).Date; Hist(lCounter).Open, Hist(lCounter).High, Hist(lCounter).Low, Hist(lCounter).Close, Hist(lCounter).Volume
    Loop
    
End Sub
