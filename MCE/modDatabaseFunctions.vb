'------------------------------------------------------------------------------------------------
' Filename    : modDatabaseFunctions.vb
' Purpose     : This is the common module that provides functions that is database related
' Created By  : Felix Kang - I-CAT Computing (28 JUL 2005)
' Note        : 
' Assumptions : 
'   - Code is based on Visual Basic .NET (Visual Studio 2003)
'   - The necessary database connector and adapters have been loaded
'------------------------------------------------------------------------------------------------
' History
' - 28 JUL 2005 : Creation date of the module
'------------------------------------------------------------------------------------------------

#Region " System Imports "

'mySQL DB library
Imports MySql.Data.MySqlClient
'MS SQL Server
Imports System.Data.SqlClient

#End Region

Module modDatabaseFunctions

#Region " Constants "

'Public Constants
Public Const NULL_KEYWORD = "NULL"
Public Const MYSQL_STRING_QUOTE = "'"

#End Region

#Region " Local Variables "

#End Region

#Region " Procedures / Functions "

Public Function ConnectToMySQLDB(ByVal strServerName As String, ByVal strUsername As String, ByVal strPassword As String, _
  ByVal strDBName As String, ByRef cnDatabaseConnection As MySqlConnection, Optional ByVal intTimeout As Integer = 15) As Boolean
'---------------------------------------------------------------------
'Purpose     : This procedure is to specifically connect to mySQL database
'Assumptions :
'Input       :
'   - strServerName, the string containing server name 
'   - strUsername, the string containing username of the account that can connect to the server
'   - strPassword, the string containing the account's password that can connect to the server
'   - strDBName, the string containing the database name 
'   - cnDatabaseConnection, if successful, this variable holds the actual connection
'   - intTimeout, an integer specifying the timeout in seconds
'Returns     :
'   - TRUE if connection was succesful and FAIL if it wasn't
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim strConnectionString As String


  'Check if there's a previous connection
  If Not cnDatabaseConnection Is Nothing Then
    'If there is then disconnect it first
    cnDatabaseConnection.Close()
  End If

  'Compose a connection string based on the information supplied
  strConnectionString = "server=" & strServerName & "; user id=" & strUsername & "; password=" & strPassword & _
  "; database=" & strDBName & "; pooling=false"

  'Connect!  
  cnDatabaseConnection = New MySqlConnection(strConnectionString)  
  cnDatabaseConnection.Open()

  'If we get to here, that means everything is fine
  ConnectToMySQLDB = True

  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "ConnectToMySQLDB", Err.Number, Err.Source, Err.Description)
  ConnectToMySQLDB = False
End Function

Public Function ExecuteMySQLStatement(ByVal strSQL As String, ByVal cnConnection As MySqlConnection, _
  Optional ByVal strConnection As String = "") As Boolean
'---------------------------------------------------------------------
'Purpose     : This procedure is to specifically connect to mySQL database
'Assumptions :
'Input       :
'   - strSQL, the string containing the SQL statement 
'   - cnConnection, the connection object that holds the database connection
'   - strConnection (optional), the string containing the connection string. It's optional because it is only needed if 
'     the connection somehow got disconnected
'Returns     :
'   - TRUE if connection was succesful and FAIL if it wasn't
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim cmdSQLStatement As MySqlCommand
  Dim cnNewConnection As MySqlConnection


  'Create a new instance of mySQL command
  cmdSQLStatement = New MySqlCommand(strSQL, cnConnection)
  'Check if a connection is already open
  If cmdSQLStatement.Connection.State.Closed Then
    'If it's closed, open it again
    '<Insert code here>
  Else
    'Execute it
    cmdSQLStatement.ExecuteNonQuery()
  End If

  'If we get this far then everything is ok
  ExecuteMySQLStatement = True

  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "ExecuteMySQLStatement", Err.Number, Err.Source, Err.Description)
  ExecuteMySQLStatement = False
End Function

Public Function RetrieveMySQLData(ByVal strSQLQuery As String, ByVal cnDBConnection As MySqlConnection, _
  ByRef dsReturnedData As DataSet, Optional ByRef intRowCount As Integer = 0, _
  Optional ByVal intRowLimit As Integer = 0, Optional ByVal strXMLOutputPath As String = "") As Boolean
'---------------------------------------------------------------------
'Purpose     : This procedure is to retrieve data from mySQL Database
'Assumptions :
'Input       :
' - strSQLQuery, the string containing the SQL Query 
' - cnDBConnection, the connection object that holds the database connection
' - dsReturnedData, the DataSet object that will hold the data
' - intRowCount (optional), an integer showing how many records retrieved from the table
' - intRowLimit, an integer indicating a limit on the number of records retrieved, by default is 0, which means all row
' - strXMLOutputPath, a string containg a complete path + filename to an XML dump of the data retrieved
'Returns     :
'   - TRUE if connection was succesful and FAIL if it wasn't
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim adpMySQLAdapter As MySqlDataAdapter
  Dim dsTemp As DataSet


  'Create a new instance of mySQL adapter which will bridge between mySQL provider and .NET's DataSet(?)
  adpMySQLAdapter = New MySqlDataAdapter

  'Create a new instance of mySQL command which will be used to retrieve the data  
  If intRowLimit > 0 Then
    adpMySQLAdapter.SelectCommand = New MySqlCommand(strSQLQuery & " LIMIT " & intRowLimit, cnDBConnection)
  Else
    adpMySQLAdapter.SelectCommand = New MySqlCommand(strSQLQuery, cnDBConnection)
  End If

  'Fill the DataSet with returning recordset
  dsReturnedData = New DataSet
  adpMySQLAdapter.Fill(dsReturnedData)
  adpMySQLAdapter.FillSchema(dsReturnedData, SchemaType.Source)

  'Try to check out how many records we got from this query
  intRowCount = dsReturnedData.Tables(0).Rows.Count()

  'If user specify an output file, dump the content of the output into an XML file
  If Trim(strXMLOutputPath) <> "" Then
    dsReturnedData.WriteXml(strXMLOutputPath, XmlWriteMode.WriteSchema)
  End If

  'If we get this far that means we're fine
  RetrieveMySQLData = True

  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "RetrieveMySQLData", Err.Number, Err.Source, Err.Description)
  RetrieveMySQLData = False
  intRowCount = -1
End Function

Public Function LoadMySQLTableIntoDataSet(ByVal strTableName As String, ByVal cnDatabase As MySqlConnection, Optional ByVal strWhere As String = "", _
  Optional ByVal strOrderBy As String = "", Optional ByRef intRowCount As Integer = 0, Optional ByRef strFinalSQL As String = "") As DataSet
'---------------------------------------------------------------------
'Purpose     : This procedure is to load the entire content of that table into a dataset
'Assumptions :
' - The connection passed as a parameter is active and open
'Input       :
' - strTableName, a string defining the real table name in mySQL database
' - cnDatabase, a valid and active database connection to mySQL database
' - strWhere (optional), a valid SQL's WHERE clause - without the actual WHERE keyword
' - strOrderBy (optional), a valid SQL's ORDER BY clause  - without the actual ORDER BY keyword
' - intRowCount (optional), an integer showing how many records retrieved from the table
' - strSQL (optional), a string consisting the final SQL statement used to load the table into DataSet
'Returns     :
'   - a dataset if it was successful, nothing if it fails
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim dsTempDataSet As DataSet
  Dim strSQL As String
  Dim blnSuccess As Boolean


  'Compose the base SQL query
  strSQL = "SELECT * FROM " & strTableName

  'If we have a WHERE statement, add it to the SQL query
  If Trim(strWhere) <> "" Then
    strSQL = strSQL & " WHERE " & strWhere
  End If

  'If we have an ORDER BY statement, add it to the SQL query
  If Trim(strOrderBy) <> "" Then
    strSQL = strSQL & " ORDER BY " & strOrderBy
  End If

  'Get the data
  blnSuccess = RetrieveMySQLData(strSQL, cnDatabase, dsTempDataSet, intRowCount)

  'If it's successful, return the DataSet
  If blnSuccess = True Then
    LoadMySQLTableIntoDataSet = dsTempDataSet
  Else
    'If it wasn't succesfull, just return code to indicate that it failed
    LoadMySQLTableIntoDataSet = Nothing
    intRowCount = -1
  End If

  'Return the final SQL statement
  strFinalSQL = strSQL

  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "LoadMySQLTableIntoDataSet", Err.Number, Err.Source, Err.Description)
  LoadMySQLTableIntoDataSet = Nothing
  intRowCount = -1
End Function

Public Function DBNullToString(ByVal oValue As Object, Optional ByVal blnUseNULLKeyword As Boolean = False) As String
'---------------------------------------------------------------------
'Purpose     : This procedure is to check whether a string value is NULL or not
'Assumptions :
'Input       :
'   - strValue, the value that we want to check
'Returns     :
'   - a string with "NULL" if it was NULL or the original data
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler


  'Check whether is it a null
  If IsDBNull(oValue) = True Then
    If blnUseNULLKeyword = True Then
      'If it is, return the string "NULL" so it's safe to be displayed
      DBNullToString = "<" & NULL_KEYWORD & ">"
    Else
      'If it's not, return empty string
      DBNullToString = ""
    End If
  Else
    'If it is not a NULL then simply pass back the data as a string. We assume that caller will not give us junk data
    DBNullToString = CStr(oValue)
  End If

  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "DBNullToString", Err.Number, Err.Source, Err.Description)
  DBNullToString = "ERR"
End Function

Public Function DBNullToInteger(ByVal oValue As Object, Optional ByVal intDefNULLValue As Integer = 0) As Integer
'---------------------------------------------------------------------
'Purpose     : This procedure is to check whether an integer value is NULL or not
'Assumptions :
'Input       :
'   - strValue, the value that we want to check
'Returns     :
'   - a integer substitute for "NULL", by default it will be zero
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler


  'Check whether is it a null
  If IsDBNull(oValue) = True Then
    DBNullToInteger = intDefNULLValue
  Else
    'If it is not a NULL then simply pass back the data as a string. We assume that caller will not give us junk data
    DBNullToInteger = CInt(oValue)
  End If
  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "DBNullToInteger", Err.Number, Err.Source, Err.Description)
  DBNullToInteger = "-99999"
End Function

Public Function DBNullToDate(ByVal oValue As Object, Optional ByVal intDefNULLValue As Date = #12:00:00 AM#) As Date
'---------------------------------------------------------------------
'Purpose     : This procedure is to check whether a date value is NULL or not
'Assumptions :
'Input       :
'   - strValue, the value that we want to check
'Returns     :
'   - a integer substitute for "NULL", by default it will be zero
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler


  'Check whether is it a null
  If IsDBNull(oValue) = True Then
    DBNullToDate = intDefNULLValue
  Else
    'If it is not a NULL then simply pass back the data as a string. We assume that caller will not give us junk data
    DBNullToDate = CDate(oValue)
  End If
  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "DBNullToDate", Err.Number, Err.Source, Err.Description)
  DBNullToDate = #12:00:00 AM#
End Function

Public Function DBNullToSingle(ByVal oValue As Object, Optional ByVal intDefNULLValue As Single = 0.0) As Single
'---------------------------------------------------------------------
'Purpose     : This procedure is to check whether a single value is NULL or not
'Assumptions :
'Input       :
'   - strValue, the value that we want to check
'Returns     :
'   - a single substitute for "NULL", by default it will be 0.0
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler


  'Check whether is it a null
  If IsDBNull(oValue) = True Then
    DBNullToSingle = intDefNULLValue
  Else
    'If it is not a NULL then simply pass back the data as a string. We assume that caller will not give us junk data
    DBNullToSingle = CSng(oValue)
  End If
  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "DBNullToSingle", Err.Number, Err.Source, Err.Description)
  DBNullToSingle = "-9999999.99"
End Function

Public Function EmptyToNULLString(ByVal strValue As String) As String
'---------------------------------------------------------------------
'Purpose     : This procedure is to convert empty value to NULL String
'Assumptions :
'Input       :
'   - strValue, the value that we want to check
'Returns     :
'   - a single substitute for "NULL", by default it will be 0.0
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler


  'Check whether is it empty or not
  If strValue = "" Then
    EmptyToNULLString = NULL_KEYWORD
  Else
    'If it is not empty then simply pass back the data as a string. We assume that caller will not give us junk data
    EmptyToNULLString = strValue
  End If
  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "EmptyToNULLString", Err.Number, Err.Source, Err.Description)
  EmptyToNULLString = ""
End Function

Public Function StripUnnumericSign(ByVal strValue As String) As String
'---------------------------------------------------------------------
'Purpose     : This procedure is to convert strip dollar sign ($)from string
'Assumptions :
'Input       :
'   - strValue, the value that we want to check
'Returns     :
'   - a string without the $ sign
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim strTempStr As String


  'Do replace $ (dollar sign) with empty string
  strTempStr = Trim(Replace(strValue, "$", ""))
  'Do replace , (comma) with empty string
  strTempStr = Trim(Replace(strTempStr, ",", ""))
  'Just basically remove any punctuation
  strTempStr = Trim(Replace(strTempStr, "/", ""))
  strTempStr = Trim(Replace(strTempStr, "\", ""))
  StripUnnumericSign = Trim(Replace(strTempStr, "'", ""))

  Exit Function

ErrorHandler:
  GenericErrorHandler("modDatabase.vb", "StripUnnumericSign", Err.Number, Err.Source, Err.Description)
  StripUnnumericSign = ""
End Function

Public Function PrepSngNumStrForDB(ByVal strNumericValue As String, _
  Optional ByVal blnTurnZeroIntoNULL As Boolean = True) As String
'---------------------------------------------------------------------
'Purpose     : This procedure is to cleanup / prepare a string containing some numeric (single type) for database use
'Assumptions :
'Input       :
' - strNumericValue, a string to use for database later on
' - blnTurnZeroIntoNULL, if TRUE it will return NULL string instead of zero
'Returns     :
' - a string containing clean number for database use
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim sngTempValue As Single
  Dim strTempString As String


  'First, strip any dollar sign in the string
  strTempString = Trim(StripUnnumericSign(strNumericValue))
  'and then, using val making sure that we get only numbers, junks will resulted in 0
  sngTempValue = Val(strTempString)

  'After that, convert it to string or NULL depending on the request
  If sngTempValue = 0 Then
    If blnTurnZeroIntoNULL = True Then
      PrepSngNumStrForDB = NULL_KEYWORD
    Else
      PrepSngNumStrForDB = CStr(sngTempValue)
    End If
  Else
    PrepSngNumStrForDB = CStr(sngTempValue)
  End If

  Exit Function

ErrorHandling:
  GenericErrorHandler("modDatabase.vb", "PrepSngNumStrForDB", Err.Number, Err.Source, Err.Description)
  PrepSngNumStrForDB = "ERR:" & Err.Description
End Function

Public Function GetSingleDBFieldValue(ByVal strSQL As String, ByVal strFieldName As String, ByVal cnDBConnection As MySqlConnection) As String
'---------------------------------------------------------------------
'Purpose     : This procedure is to get a single database field value
'Assumptions :
'Input       :
' - strSQL, a string containing a valid SQL statement
' - strFieldName, a string with single field name which value needs to be accessed
' - cnDBConnection, a valid and active database connection to mySQL database
'Returns     :
' - a string returning the value of the requested field, empty string ("") will be returned if there's a problem
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim blnSuccess As Boolean
  Dim dsTempDataSet As DataSet
  Dim intTempRowCounter As Integer
  Dim drTempDataRow As DataRow


  'Get our record
  blnSuccess = RetrieveMySQLData(strSQL, cnDBConnection, dsTempDataSet, intTempRowCounter)
  'Check whether we get something
  If intTempRowCounter > 0 Then
    drTempDataRow = dsTempDataSet.Tables(0).Rows(0)
    'Return the value
    GetSingleDBFieldValue = (DBNullToString(drTempDataRow(strFieldName)))
  Else
    'If nothing, just return blank
    GetSingleDBFieldValue = ""
  End If

  Exit Function

ErrorHandling:
  GenericErrorHandler("modDatabase.vb", "GetSingleDBFieldValue", Err.Number, Err.Source, Err.Description)
  GetSingleDBFieldValue = "ERR:" & Err.Description
End Function

Public Function ConnectToMSSQLDatabase(ByVal strServerName As String, ByVal strDBName As String, _
  ByVal strDBUsername As String, ByVal strDBPassword As String, Optional ByVal intTimeout As Integer = -1, _
  Optional ByRef blnSuccess As Boolean = False) As SqlConnection
'------------------------------------------------------------------------------------------------
' Purpose     : Open up connection to the MS SQL Server Database engine
' Assumption  : 
'   - Imports System.Data.SqlClient has been declared
'   - Destination variable need to be created beforehand
' Input       :
'   - strServerName, a string containing the server that we want to connect to
'   - strDBName, a string containing the database name
'   - strDBUsername, a string with the username to connect to the database
'   - strDBPassword, a string with the password to connect to the database
'   - blnSuccess, contains TRUE if the connection was succesful and FALSE if it fails
'   - intTimeout, integer value for user to specify timeout value, use -1 if no change is needed or 
'     use default value (in seconds)
' Output      :
' Created By  : Felix Kang
' Note        : - This function will automatically close and re-open the database
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim cnMSSQL As SqlConnection
  Dim strConnection As String


  'Check whether the destination variable already has an open connection
  If Not ConnectToMSSQLDatabase Is Nothing Then
    If ConnectToMSSQLDatabase.State = ConnectionState.Open Then
      'If it's open, close it first
      ConnectToMSSQLDatabase.Close()
    End If
  End If

  'Compose the connection string used to open the database
  strConnection = "server=" & strServerName & ";uid=" & strDBUsername & ";pwd=" & strDBPassword & ";" & _
  "database=" & strDBName
  If intTimeout > 0 Then
    'If the user wants to modify the length of timeout
    strConnection = strConnection & ";Connection Timeout=" & intTimeout
  End If

  'Create a new connection instance
  cnMSSQL = New SqlConnection(strConnection)
  'and then open a new connection
  cnMSSQL.Open()

  'See whether the connection is open now
  If cnMSSQL.State = ConnectionState.Open Then
    'Pass it back
    ConnectToMSSQLDatabase = cnMSSQL
    'and return success flag
    blnSuccess = True
  Else
    'if it's not open then return to come user and says something is not right
    ConnectToMSSQLDatabase = Nothing
    blnSuccess = False
  End If

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modDatabase.vb", "ConnectToMSSQLDatabase", Err.Number, Err.Source, Err.Description)
  ConnectToMSSQLDatabase = Nothing
  'and return fail flag
  blnSuccess = False
End Function

Public Function CleanUpSQLForMySQL(ByVal strInput As String, Optional ByVal blnEmptyForNULL As Boolean = False) As String
'---------------------------------------------------------------------
'Purpose     : This procedure is to parse/clean up any character in string so it can be save properly into mySQL
'Assumptions :
'Input       :
' - strInput, a string stating text input
'Returns     :
' - The clean string ready for mySQL to be executed
'Note        : 
' - mySQL will automatically remove any reserve character in string insertion. So in order to user them, we need to double
'   the amount of that character. "\\" will be saved as "\" in DB and "\\\\" will be saved as "\\", etc.
'---------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim strTemp As String


  'Move it into temp variable so it can easily manipulated
  strTemp = Trim(strInput)

  'Replace "\" with "\\"
  strTemp = Replace(strTemp, "\", "\\")
  'Replace "'" with "''"
  strTemp = Replace(strTemp, "'", "''")

  'Check whether user wants to replace empty string/value with nulls
  If (strTemp = "") And (blnEmptyForNULL = True) Then
    strTemp = NULL_KEYWORD
  End If

  'When we done, pass back the clean string
  CleanUpSQLForMySQL = strTemp

  Exit Function

ErrorHandling:
  GenericErrorHandler("modDatabase.vb", "CleanUpSQLForMySQL", Err.Number, Err.Source, Err.Description)
  CleanUpSQLForMySQL = "ERR:" & Err.Description
End Function

Public Function CreateListViewFromDS(ByVal dsData As DataSet, ByRef lstTarget As ListView, _
  Optional ByVal intColumnTextAlignment As System.Windows.Forms.HorizontalAlignment = HorizontalAlignment.Left, _
  Optional ByVal intDefColumnWidth As Integer = 110) As Boolean
'---------------------------------------------------------------------
'Purpose     : This procedure is to create the column header in list using the fields/columns from a DataSet
'Assumptions :
'   - DataSet is not empty and it has at least 1 row
'Input       :
' - dsData, a DataSet object containing the records
'Returns     :
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim intTempLoop As Integer
  Dim intTempLoop2 As Integer
  Dim clmTemp As ColumnHeader
  Dim intRowCount As Integer
  Dim drTempRow As DataRow
  Dim lstTempItem As ListViewItem


  'Create the column header first
  lstTarget.Columns.Clear()
  For intTempLoop = 0 To dsData.Tables(0).Columns.Count - 1
    'Create a temp column headers from the fields
    clmTemp = New ColumnHeader
    'Setup its properties
    clmTemp.Text = dsData.Tables(0).Columns(intTempLoop).ColumnName
    'Default width
    clmTemp.Width = intDefColumnWidth
    'Alignment
    clmTemp.TextAlign = intColumnTextAlignment

    'Add it to the listview
    lstTarget.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {clmTemp})
  Next

  'Check whether we have records to use.
  If dsData.Tables(0).Rows.Count < 1 Then
    'Return TRUE to say that everything is ok
    CreateListViewFromDS = True
    'but no need to proceed
    Exit Function
  End If

  'If we're here that means we're good to go
  lstTarget.Items.Clear()
  For intTempLoop = 0 To dsData.Tables(0).Rows.Count - 1
    'Get a row
    drTempRow = dsData.Tables(0).Rows(intTempLoop)
    'Create a "head" of that row
    lstTempItem = New ListViewItem(DBNullToString(drTempRow(lstTarget.Columns(0).Text)))

    'And then loop for each subitem
    For intTempLoop2 = 1 To lstTarget.Columns.Count - 1
      'Adding the details            
      lstTempItem.SubItems.Add(DBNullToString(drTempRow(lstTarget.Columns(intTempLoop2).Text)))
    Next

    'Add the row to the listview  
    lstTarget.Items.AddRange(New ListViewItem() {lstTempItem})
  Next

  'Return TRUE to say that everything is ok
  CreateListViewFromDS = True

  Exit Function

ErrorHandling:
  GenericErrorHandler("modDatabase", "CreateListViewFromDS", Err.Number, Err.Source, Err.Description)
  CreateListViewFromDS = False
End Function

Public Function GetMySQLTableFields(ByVal strTableName As String, ByVal cnDBConnection As MySqlConnection) As DataSet
'---------------------------------------------------------------------
'Purpose     : This procedure is to create the column header in list using the fields/columns from a DataSet
'Assumptions :
'   - DataSet is not empty and it has at least 1 row
'Input       :
' - dsData, a DataSet object containing the records
'Returns     :
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Exit Function

ErrorHandling:
  GenericErrorHandler("modDatabase", "GetMySQLTableFields", Err.Number, Err.Source, Err.Description)
  GetMySQLTableFields = Nothing
End Function

#End Region


End Module
