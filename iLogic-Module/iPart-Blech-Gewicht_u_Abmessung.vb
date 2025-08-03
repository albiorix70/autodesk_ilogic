AddReference "Microsoft.Office.Interop.Excel.dll"
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Sub Main()
	Dim oDoc As Document = ThisDoc.Document
	Dim bUpdate As Boolean = True

	Dim bScreenUpd = MessageBox.Show("Bildschirmaktualisierung ausschalten", "Update Screen", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

	If bScreenUpd = vbYes Then
		bUpdate = False
	End If		

	ThisApplication.SilentOperation = True
	ThisApplication.ScreenUpdating = bUpdate

	If oDoc.DocumentSubType.DocumentSubTypeID = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
        AbwicklungAnlegen()
		AbwicklungDimBerechnen(bUpdate)
    End If
	
	ThisApplication.SilentOperation = False
	ThisApplication.ScreenUpdating = True

End Sub

Private Sub AbwicklungAnlegen()
    Dim oCompDef as SheetMetalComponentDefinition = ThisDoc.Document.ComponentDefinition

    If oCompDef.HasFlatPattern() Then
		Logger.Info(iLogicVb.RuleName & ": " & ThisDoc.FileName(false) & " Abwicklung bereits vorhanden")
	Else
		Logger.Warn(iLogicVb.RuleName & ": " & ThisDoc.FileName(false) & " Abwicklung nicht vorhanden, wird erstellt")
		
		oCompDef.Unfold()
		oCompDef.FlatPattern.ExitEdit()
	End If
End Sub

Private Sub AbwicklungDimBerechnen(Optional pProgessInStatus As Boolean = True)
	Dim oCompDef As SheetMetalComponentDefinition = ThisDoc.Document.ComponentDefinition
	Dim UoM As UnitsOfMeasure = ThisDoc.Document.UnitsOfMeasure
		logger.info(UoM.LengthUnits & " " & UoM.GetStringFromType(UoM.LengthUnits))
	' Exit, wenn keine iPartFactory
	If Not  oCompDef.IsiPartFactory Then	
		Exit Sub
	End If
	
	Dim oPartFac As iPartFactory = oCompDef.iPartFactory
	Dim oProgress As Inventor.ProgressBar = ThisApplication.CreateProgressBar(pProgessInStatus, oPartFac.TableRows.Count, "Fortschritt")
	oProgress.Message = "Starting ...."

	oProgress.Message = "Init Excel Sheet"
	CreateExcelCols()
	oProgress.Message =  "Excel Sheet prepared"

    Dim oFlatPattern As FlatPattern = oCompDef.FlatPattern
	' Reload iPartFactory, nach dem Anlegen der Excel-Spalten
	oPartFac = oCompDef.iPartFactory
	Dim initRow As iPartTableRow = oPartFac.DefaultRow

	' durch die Parts laufen
	For Each oPartRow In oPartFac.TableRows
		oProgress.Message = "Step " & oPartFac.DefaultRow.Index & " von " & oPartFac.TableRows.Count
		oPartFac.DefaultRow = oPartRow

		Dim dWeight As Double = oCompDef.MassProperties.Mass
		Dim dFlatWeight As Double = oFlatPattern.MassProperties.Mass

		Dim dLength As Double = Round(UoM.ConvertUnits(oFlatPattern.Length, _ 
				UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits), 1)
		Dim dWidth As Double = Round(UoM.ConvertUnits(oFlatPattern.Width, _ 
				UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits), 1)

		' Aufrunden auf den nächsten mm nach oben
		dLength = Ceil(dLength)
		dWidth = Ceil(dWidth)
		
		SetColumnValue(oPartFac.DefaultRow, "Gewicht", UoM.GetStringFromValue( dWeight, UoM.MassUnits))
		SetColumnValue(oPartFac.DefaultRow, "Gewicht Abwicklung", UoM.GetStringFromValue( dFlatWeight, UoM.MassUnits))
		SetColumnValue(oPartFac.DefaultRow, "Länge Abwicklung", dLength)
		SetColumnValue(oPartFac.DefaultRow, "Breite Abwicklung", dWidth)
		oProgress.UpdateProgress
	Next
	' letzten Wert der Zeile wieder setzen
	oPartFac.DefaultRow = initRow

	oProgress.Message = "Fertig"
	oProgress.Close
End Sub

Private Sub CreateExcelCols()
	Dim lstColumnNames = New String(){"Gewicht", "Gewicht Abwicklung", "Länge Abwicklung", "Breite Abwicklung"}
	Dim oFactory As iPartFactory = ThisDoc.Document.ComponentDefinition.iPartFactory

	Dim xlApp As Excel.Application = New Excel.Application
	Dim xlWorkbook As Excel.Workbook
	Dim xlWorksheet as Worksheet = oFactory.ExcelWorkSheet
	
	' Abkürzung für die iProperties
	Dim invCustomPropertySet As PropertySet = _
    	ThisDoc.Document.PropertySets.Item("Inventor User Defined Properties")
    ' Workbook öffnen
	For Each sColName In lstColumnNames
		On Error Resume Next
		Dim invProp As Inventor.Property = invCustomPropertySet.Item(sColName)
		' Ok, Property nicht vorhanden, also anlegen
		If Err.Number <> 0 Then
			Logger.Info(sColName & " nicht in iProperties vorhanden, wird angelegt" )
			invCustomPropertySet.Add("", sColName)
			Err.Clear
		End If
		
		If Not ColumnExists(sColName, xlWorksheet) Then
			Dim iFreeCol As Integer = FindFirstFreeColumn(xlWorksheet)
			Logger.Info("Create new Col " & sColName & " at Position " & iFreeCol)
			xlWorksheet.Cells(1, iFreeCol).value = sColName & " [Custom]"
		End If
	Next

	 xlWorkbook = xlWorksheet.Parent

	' Save and close workbook
	xlWorkbook.Save
	xlWorkbook.Close

	' Quit Excel application
	xlApp.Quit
	xlApp = Nothing
End Sub

' Function To check If a Column exists
Private Function ColumnExists(columnName As String, xlWorksheet As Excel.Worksheet) As Boolean
	Dim cell As Excel.Range
	For Each cell In xlWorksheet.Rows(1).Cells
		If cell.Value = columnName & " [Custom]" Then
			ColumnExists = True
			Exit Function
		End If

		If cell.Value Is Nothing Then
			ColumnExists = False
			Exit Function
		end if
	Next

	ColumnExists = False
End Function

' Find the first free column
Private Function FindFirstFreeColumn(xlWorksheet As Excel.Worksheet) As Integer
	Dim cell As Excel.Range
	
	For Each cell In xlWorksheet.Rows(1).Cells
		If cell.Value Is Nothing Then
			FindFirstFreeColumn = cell.Column
			Exit Function
		End If
	Next

	' If no free column exists, add a new one at the end
	FindFirstFreeColumn = Nothing
End Function

Private Sub SetColumnValue(ByRef pRow as iPartTableRow, pPropertyName As String, _
			Optional pPropertyValue as String = "")
	
	If not pRow.item(pPropertyName + " [Custom]") is Nothing Then
		pRow.item(pPropertyName + " [Custom]").Value = pPropertyValue
	end if
End Sub
