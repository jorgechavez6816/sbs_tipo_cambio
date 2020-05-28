Sub Main
	IgnoreWarning(True)
	'Client.RunPython "C:\Users\Intel\Documents\Mis documentos IDEA\Samples\Macros.ILB\tipo_cambio_sbs.py"
	Call TextImport()	'C:\ProgramData\Anaconda3\Python_Notebooks\Tipo_cambio.csv
	Call JoinDatabase()	'Ejemplo-Detalle de ventas.IMD
	Call AppendField()	'Ejemplo_Detalle de ventas_Moneda.IMD
	Client.RefreshFileExplorer
End Sub


' Archivo - Asistente de importación: Texto delimitado
Function TextImport
	dbName = "Tipo_cambio.IMD"
	Client.ImportUTF8DelimFile "C:\Users\Intel\Documents\Mis documentos IDEA\Samples\Archivos fuente.ILB\Tipo_cambio.csv", dbName, FALSE, "", "C:\Users\Intel\Documents\Mis documentos IDEA\Samples\Definiciones de importación.ILB\Tipo_cambio.RDF", TRUE
	Client.OpenDatabase (dbName)
End Function


' Archivo: Unir bases de datos
Function JoinDatabase
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Tipo_cambio.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "MONEDA", "MONEDA", "A"
	task.CreateVirtualDatabase = False
	dbName = "Ejemplo_Detalle de ventas_Moneda.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Anexar campo
Function AppendField
	Set db = Client.OpenDatabase("Ejemplo_Detalle de ventas_Moneda.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TOTAL_MONEDA"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = " TOTAL * VENTA "
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function







