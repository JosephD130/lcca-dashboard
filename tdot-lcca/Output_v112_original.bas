VBA MACRO Output.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Output'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Public Sub SetupSummaryWs()
'This sub writes data to the summary worksheet
    'Declarations
        Dim colAltNum As Integer, colAltName As Integer, colInitConstCost As Integer, colPWCost As Integer, colDescription As Integer
        Dim startWsRow As Integer, curWsRow As Integer
        
        Dim curAltObj As clsAlternative
        Dim curAltNum As Integer
        
        Dim formulaStr_InitConst As String, formulaStr_PWConst As String
        
        Dim myChtObj As ChartObject, chartSeries As Series
        Dim rngChartXValues As Range, rngChartYValues As Range
        
        Dim rng As Range
        
        Dim numAlts As Integer
        
    'Initialization
        colAltNum = 1
        colAltName = 2
        colInitConstCost = 3
        colPWCost = 4
        colDescription = 5
        
        startWsRow = 4
        curWsRow = startWsRow
        curAltNum = 1
        numAlts = MainDoc.AlternativeWsCollection.Count
    
    'Code
        'Clear the data range'
            With GlobalVars.wsSummary
                Set rng = .Range(.Cells(curWsRow, colAltNum), .Cells(curWsRow + 100, colDescription))
            End With
            rng.Clear
        
        Dim foundCellNet As Range
        Dim foundCellTotal As Range
        Dim pwValue As Variant
        'Walk alternative list and write details to summary ws
            For Each curAltObj In MainDoc.AlternativeWsCollection
                
                ' Find what row "Net Present Worth" is on
                Set foundCellNet = Worksheets(curAltObj.NameWsAlt).Columns("B").Find( _
                        What:="Net Present Worth", _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole)
                        
                        
                Set foundCellTotal = Worksheets(curAltObj.NameWsAlt).Columns("A").Find( _
                        What:="Total", _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole)
                        
                'Determine needed function strings (depends on alternative type)
                
                    Select Case curAltObj.AltType
                        Case EnumAltType.NewHMA
                            formulaStr_InitConst = "='" & curAltObj.NameWsAlt & "'!" & Cells(foundCellTotal.Row, 7).Address
                            formulaStr_PWConst = "='" & curAltObj.NameWsAlt & "'!" & Cells(foundCellNet.Row, 5).Address
                            
                        Case EnumAltType.NewPCC
                            formulaStr_InitConst = "='" & curAltObj.NameWsAlt & "'!" & Cells(foundCellTotal.Row, 7).Address
                            formulaStr_PWConst = "='" & curAltObj.NameWsAlt & "'!" & Cells(foundCellNet.Row, 5).Address
                            
'                        Case EnumAltType.RehabHMA
'                            formulaStr_InitConst = "='" & curAltObj.NameWsAlt & "'!" & Cells(26, 7).Address
'                            formulaStr_PWConst = "='" & curAltObj.NameWsAlt & "'!" & Cells(45, 5).Address
'
'                        Case EnumAltType.RehabPCC
'                            formulaStr_InitConst = "='" & curAltObj.NameWsAlt & "'!" & Cells(26, 7).Address 'TO DO
'                            formulaStr_PWConst = "='" & curAltObj.NameWsAlt & "'!" & Cells(45, 5).Address  'TO DO
                    End Select
                    
                'Write data to summary ws
                    GlobalVars.wsSummary.Cells(curWsRow, colAltNum).value = "Alt " & curAltNum
                    GlobalVars.wsSummary.Cells(curWsRow, colAltName).value = curAltObj.Name
                    GlobalVars.wsSummary.Cells(curWsRow, colInitConstCost).Formula = formulaStr_InitConst
                    GlobalVars.wsSummary.Cells(curWsRow, colPWCost).Formula = formulaStr_PWConst
                    GlobalVars.wsSummary.Cells(curWsRow, colDescription).value = curAltObj.Description
            
                'Advance counters
                    curAltNum = curAltNum + 1
                    curWsRow = curWsRow + 1
            Next
        
        'Set chart information
            Set myChtObj = GlobalVars.wsSummary.ChartObjects("Chart 1")
            
        'Series 1: Initial construction cost
            Set chartSeries = myChtObj.Chart.SeriesCollection(1)
            
            'Set XValues range
            With GlobalVars.wsSummary
                Set rngChartXValues = .Range(.Cells(startWsRow, colAltNum), .Cells(startWsRow + numAlts - 1, colAltNum))
                Set rngChartYValues = .Range(.Cells(startWsRow, colInitConstCost), .Cells(startWsRow + numAlts - 1, colInitConstCost))
            End With
            
            chartSeries.XValues = rngChartXValues
            chartSeries.Values = rngChartYValues
            
        'Series 2: Present worth cost
            Set chartSeries = myChtObj.Chart.SeriesCollection(2)
            
            'Set XValues range
            With GlobalVars.wsSummary
                Set rngChartXValues = .Range(.Cells(startWsRow, colAltNum), .Cells(startWsRow + numAlts - 1, colAltNum))
                Set rngChartYValues = .Range(.Cells(startWsRow, colPWCost), .Cells(startWsRow + numAlts - 1, colPWCost))
            End With
            
            chartSeries.XValues = rngChartXValues
            chartSeries.Values = rngChartYValues
            
            'Debug.Print chartSeries.Name

        'Clean up
            Set curAltObj = Nothing
            Set rng = Nothing
            Set rngChartXValues = Nothing
            Set rngChartYValues = Nothing
End Sub
-------------------------------------------------------------------------------
