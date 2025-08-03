Dim MassProperty As String

'Type in here the iProperty Name
MassProperty = "Gewicht"


Dim ColIndex As Integer
Dim oPrt As Inventor.PartDocument
oPrt = ThisDoc.Document
Dim oFact As iPartFactory
oFact = oPrt.ComponentDefinition.iPartFactory

'Code runs only in iPart Factory 
If oPrt.ComponentDefinition.IsiPartFactory = True Then
    Dim oFacCol As iPartTableColumn
	'Get the Colum with the iProperty Name
    For Each oFacCol In oFact.TableColumns
        If oFacCol.DisplayHeading = MassProperty Then
            ColIndex = oFacCol.Index
        End If
    Next

	
    Dim oFacRow As iPartTableRow
    For Each oFacRow In oFact.TableRows
        ' Set Row as active Row
		oFact.DefaultRow = oFacRow
        
		' Get the weight.
        Dim dWeight As Double
        dWeight = oPrt.ComponentDefinition.MassProperties.Mass
		
		'Convert in weight from doc settings
        Dim strWeight As String
		strWeight = oPrt.UnitsOfMeasure.GetStringFromValue(dWeight, 11281) 'In VBA kDefaultDisplayMassUnits
                   
        ' Set the row value for the weight column.
        oFacRow.Item(ColIndex).Value = strWeight
    Next

End If