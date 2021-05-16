'**********************************
'* Name: Connection
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Connection
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.5
'* Create Time: 18/2/2021
'* 1.0.2  20/2/2021   Modify mExecute,Errors
'* 1.0.3  13/3/2021   Add ProviderEnum,SetConnSQLServer,SetConnAccess
'* 1.0.4  19/3/2021   Add DBTypeEnum,DBType
'* 1.0.5  16/4/2021	Remove excess Me.ClearErr(), Modify BeginTrans
'* 1.0.6  2/5/2021	Add ConnStateEnum, modify State
'**********************************
Public Class Connection
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.5"
	Public Obj As Object

	Public Enum DBTypeEnum
		Unknow = 0
		SQLServer = 10
		Oracle = 20
		MySQL = 30
		Access = 40
	End Enum

	Public Enum ConnStateEnum
		adStateClosed = 0
		adStateOpen = 1
		adStateConnecting = 2
		adStateExecuting = 4
		adStateFetching = 8
	End Enum

	Public Enum ProviderEnum
		ActiveDirectoryServices = 10
		MicrosoftJetDatabases = 20
		MicrosoftInternetPublishing = 30
		OracleDatabases = 40
		SimpleTextFiles = 50
		MicrosoftOLEDBProviderForODBC = 60
		MicrosoftDataShape = 70
		LocallySavedFiles = 80
		MicrosoftSQLServer = 90
		MicrosoftSQLServer2012NativeClient = 100
	End Enum

	Public Enum SchemaEnum
		adSchemaActions = 41
		adSchemaAsserts = 0
		adSchemaCatalogs = 1
		adSchemaCharacterSets = 2
		adSchemaCheckConstraints = 5
		adSchemaCollations = 3
		adSchemaColumnPrivileges = 13
		adSchemaColumns = 4
		adSchemaColumnsDomainUsage = 11
		adSchemaCommands = 42
		adSchemaConstraintColumnUsage = 6
		adSchemaConstraintTableUsage = 7
		adSchemaCubes = 32
		adSchemaDBInfoKeywords = 30
		adSchemaDBInfoLiterals = 31
		adSchemaDimensions = 33
		adSchemaForeignKeys = 27
		adSchemaFunctions = 40
		adSchemaHierarchies = 34
		adSchemaIndexes = 12
		adSchemaKeyColumnUsage = 8
		adSchemaLevels = 35
		adSchemaMeasures = 36
		adSchemaMembers = 38
		adSchemaPrimaryKeys = 28
		adSchemaProcedureColumns = 29
		adSchemaProcedureParameters = 26
		adSchemaProcedures = 16
		adSchemaProperties = 37
		adSchemaProviderSpecific = -1
		adSchemaProviderTypes = 22
		adSchemaReferentialConstraints = 9
		adSchemaSchemata = 17
		adSchemaSets = 43
		adSchemaSQLLanguages = 18
		adSchemaStatistics = 19
		adSchemaTableConstraints = 10
		adSchemaTablePrivileges = 14
		adSchemaTables = 20
		adSchemaTranslations = 21
		adSchemaTrustees = 39
		adSchemaUsagePrivileges = 15
		adSchemaViewColumnUsage = 24
		adSchemaViews = 23
		adSchemaViewTableUsage = 25
	End Enum
	Public Enum IsolationLevelEnum
		adXactBrowse = 256
		adXactChaos = 16
		adXactCursorStability = 4096
		adXactIsolated = 1048576
		adXactReadCommitted = 4096
		adXactReadUncommitted = 256
		adXactRepeatableRead = 65536
		adXactSerializable = 1048576
		adXactUnspecified = -1
	End Enum
	Public Enum CursorLocationEnum
		adUseClient = 3
		adUseServer = 2
	End Enum
	Public Enum ConnectModeEnum
		adModeRead = 1
		adModeReadWrite = 3
		adModeRecursive = 4194304
		adModeShareDenyNone = 16
		adModeShareDenyRead = 4
		adModeShareDenyWrite = 8
		adModeShareExclusive = 12
		adModeUnknown = 0
		adModeWrite = 2
	End Enum

	Public Sub New()
		MyBase.New(CLS_VERSION)
		Try
			Me.Obj = CreateObject("ADODB.Connection")
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", ex)
		End Try
	End Sub
	Public Property Attributes() As Long
		Get
			Try
				Return Me.Obj.Attributes
			Catch ex As Exception
				Me.SetSubErrInf("Attributes.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.Attributes = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Attributes.Set", ex)
			End Try
		End Set
	End Property
	Public Function BeginTrans() As Object
		Try
			BeginTrans = Me.Obj.BeginTrans()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("BeginTrans", ex)
			Return Nothing
		End Try
	End Function
	Public Sub Cancel()
		Try
			Me.Obj.Cancel()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Cancel", ex)
		End Try
	End Sub
	Public Sub Close()
		Try
			Me.Obj.Close()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Close", ex)
		End Try
	End Sub
	Public Property CommandTimeout() As Long
		Get
			Try
				Return Me.Obj.CommandTimeout
			Catch ex As Exception
				Me.SetSubErrInf("CommandTimeout.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.CommandTimeout = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CommandTimeout.Set", ex)
			End Try
		End Set
	End Property
	Public Sub CommitTrans()
		Try
			Me.Obj.CommitTrans()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CommitTrans", ex)
		End Try
	End Sub
	Public Property ConnectionString() As String
		Get
			Try
				If Me.Obj Is Nothing Then Throw New Exception("Obj is Nothing")
				Return Me.Obj.ConnectionString
			Catch ex As Exception
				Me.SetSubErrInf("ConnectionString.Get", ex)
				Return ""
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.ConnectionString = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ConnectionString.Set", ex)
			End Try
		End Set
	End Property
	Public Property ConnectionTimeout() As Long
		Get
			Try
				Return Me.Obj.ConnectionTimeout
			Catch ex As Exception
				Me.SetSubErrInf("ConnectionTimeout.Get", ex)
				Return 0
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.ConnectionTimeout = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ConnectionTimeout.Set", ex)
			End Try
		End Set
	End Property
	Public Property CursorLocation() As CursorLocationEnum
		Get
			Try
				Return Me.Obj.CursorLocation
			Catch ex As Exception
				Me.SetSubErrInf("CursorLocation.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As CursorLocationEnum)
			Try
				Me.Obj.CursorLocation = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CursorLocation.Set", ex)
			End Try
		End Set
	End Property
	Public Property DefaultDatabase() As String
		Get
			Try
				Return Me.Obj.DefaultDatabase
			Catch ex As Exception
				Me.SetSubErrInf("DefaultDatabase.Get", ex)
				Return ""
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.DefaultDatabase = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DefaultDatabase.Set", ex)
			End Try
		End Set
	End Property
	Public Property Errors() As Errors
		Get
			Try
				Dim oErrors As New Errors
				oErrors.Obj = Me.Obj.Errors
				Return oErrors.Obj
			Catch ex As Exception
				Me.SetSubErrInf("Errors.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Errors)
			Try
				Me.Obj.Errors = value.Obj
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Errors.Set", ex)
			End Try
		End Set
	End Property
	Public Function Execute(CommandText As String) As Recordset
		Return Me.mExecute(CommandText)
	End Function

	Public Function Execute(CommandText As String, ByRef RecordsAffected As Long) As Recordset
		Return Me.mExecute(CommandText, RecordsAffected)
	End Function

	Public Function Execute(CommandText As String, Optional ByRef RecordsAffected As Object = Nothing, Optional Options As Long = -1) As Recordset
		Return Me.mExecute(CommandText, RecordsAffected, Options)
	End Function

	Private Function mExecute(CommandText As String, Optional ByRef RecordsAffected As Object = Nothing, Optional Options As Long = -1) As Recordset
		Try
			mExecute = New Recordset
			If RecordsAffected Is Nothing Then
				If Options = -1 Then
					mExecute.Obj = Me.Obj.Execute(CommandText)
				Else
					mExecute.Obj = Me.Obj.Execute(CommandText,, Options)
				End If
			Else
				If Options = -1 Then
					mExecute.Obj = Me.Obj.Execute(CommandText, RecordsAffected)
				Else
					mExecute.Obj = Me.Obj.Execute(CommandText, RecordsAffected, Options)
				End If
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mExecute", ex)
			Return Nothing
		End Try
	End Function
	Public Property IsolationLevel() As IsolationLevelEnum
		Get
			Try
				Return Me.Obj.IsolationLevel
			Catch ex As Exception
				Me.SetSubErrInf("IsolationLevel.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As IsolationLevelEnum)
			Try
				Me.Obj.IsolationLevel = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("IsolationLevel.Set", ex)
			End Try
		End Set
	End Property
	Public Property Mode() As ConnectModeEnum
		Get
			Try
				Return Me.Obj.Mode
			Catch ex As Exception
				Me.SetSubErrInf("Mode.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As ConnectModeEnum)
			Try
				Me.Obj.Mode = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Mode.Set", ex)
			End Try
		End Set
	End Property
	Public Sub Open(Optional ConnectionString As String = "", Optional UserID As String = "", Optional Password As String = "", Optional Options As Long = -1)
		Try
			If ConnectionString = "" Then
				If Options = -1 Then
					Me.Obj.Open()
				Else
					Me.Obj.Open(,,, Options)
				End If
			Else
				If UserID <> "" Then
					If Options = -1 Then
						Me.Obj.Open(ConnectionString, UserID, Password)
					Else
						Me.Obj.Open(ConnectionString, UserID, Password, Options)
					End If
				Else
					If Options = -1 Then
						Me.Obj.Open(ConnectionString)
					Else
						Me.Obj.Open(ConnectionString,, Options)
					End If
				End If
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Open", ex)
		End Try
	End Sub
	Public Function OpenSchema(Schema As SchemaEnum, Optional Restrictions As Object = Nothing, Optional SchemaID As Object = Nothing) As Recordset
		Try
			If Restrictions Is Nothing Then
				If SchemaID Is Nothing Then
					OpenSchema = Me.Obj.OpenSchema(Schema)
				Else
					OpenSchema = Me.Obj.OpenSchema(Schema, , SchemaID)
				End If
			Else
				If SchemaID Is Nothing Then
					OpenSchema = Me.Obj.OpenSchema(Schema, Restrictions)
				Else
					OpenSchema = Me.Obj.OpenSchema(Schema, Restrictions, SchemaID)
				End If
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("OpenSchema", ex)
			Return Nothing
		End Try
	End Function
	Public Property Properties() As Properties
		Get
			Try
				Return Me.Obj.Properties
			Catch ex As Exception
				Me.SetSubErrInf("Properties.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Properties)
			Try
				Me.Obj.Properties = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Properties.Set", ex)
			End Try
		End Set
	End Property
	Public Property Provider() As String
		Get
			Try
				Return Me.Obj.Provider
			Catch ex As Exception
				Me.SetSubErrInf("Provider.Get", ex)
				Return ""
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.Provider = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Provider.Set", ex)
			End Try
		End Set
	End Property
	Public Sub RollbackTrans()
		Try
			Me.Obj.RollbackTrans()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("RollbackTrans", ex)
		End Try
	End Sub
	Public Property State() As ConnStateEnum
		Get
			Try
				Return Me.Obj.State
			Catch ex As Exception
				Me.SetSubErrInf("State.Get", ex)
				Return ConnStateEnum.adStateClosed
			End Try
		End Get
		Set(value As ConnStateEnum)
			Try
				Me.Obj.State = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("State.Set", ex)
			End Try
		End Set
	End Property

	Public Sub SetConnSQLServer(SQLServer As String, DBUser As String, DBUserPwd As String, CurrDatabase As String, Optional Provider As ProviderEnum = ProviderEnum.MicrosoftSQLServer)
		Try
			Dim strConn As String = ""
			Select Case Provider
				Case ProviderEnum.MicrosoftSQLServer, ProviderEnum.MicrosoftSQLServer2012NativeClient
					strConn = Me.mGetProviderStr(Provider)
				Case Else
					Throw New Exception("Unsupported Provider")
			End Select
			DBUserPwd = Replace(DBUserPwd, "'", "''")
			strConn &= "Data Source=" & SQLServer & ";Database=" & CurrDatabase & ";User ID='" & DBUser & "';Password='" & DBUserPwd & "';"
			Me.ConnectionString = strConn
			Me.DBType = DBTypeEnum.SQLServer
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("SetConnSQLServer", ex)
		End Try
	End Sub

	Public Sub SetConnAccess(AccessFilePath As String)
		Try
			Dim strConn As String = ""
			strConn = Me.mGetProviderStr(ProviderEnum.MicrosoftJetDatabases)
			strConn &= "Data Source=" & AccessFilePath
			Me.ConnectionString = strConn
			Me.DBType = DBTypeEnum.Access
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("SetConnAccess", ex)
		End Try
	End Sub

	Public Sub SetConnSQLServer(SQLServer As String, CurrDatabase As String, Optional Provider As ProviderEnum = ProviderEnum.MicrosoftSQLServer)
		Try
			Dim strConn As String = ""
			Select Case Provider
				Case ProviderEnum.MicrosoftSQLServer, ProviderEnum.MicrosoftSQLServer2012NativeClient
					strConn = Me.mGetProviderStr(Provider)
				Case Else
					Throw New Exception("Unsupported Provider")
			End Select
			strConn &= "Data Source=" & SQLServer & ";Database=" & CurrDatabase & ";Integrated Security=SSPI;"
			Me.ConnectionString = strConn
			Me.DBType = DBTypeEnum.SQLServer
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("SetConnSQLServer", ex)
		End Try
	End Sub

	Private Function mGetProviderStr(Provider As ProviderEnum) As String
		Try
			mGetProviderStr = "Provider="
			Select Case Provider
				Case ProviderEnum.ActiveDirectoryServices
					mGetProviderStr &= "ADSDSOObject"
				Case ProviderEnum.LocallySavedFiles
					mGetProviderStr &= "MSPersist"
				Case ProviderEnum.MicrosoftDataShape
					mGetProviderStr &= "MSDataShape"
				Case ProviderEnum.MicrosoftInternetPublishing
					mGetProviderStr &= "MSDAIPP.DSO.1"
				Case ProviderEnum.MicrosoftJetDatabases
					mGetProviderStr &= "Microsoft.Jet.OLEDB.4.0"
				Case ProviderEnum.MicrosoftOLEDBProviderForODBC
					mGetProviderStr &= "MSDASQL"
				Case ProviderEnum.MicrosoftSQLServer
					mGetProviderStr &= "SQLOLEDB"
				Case ProviderEnum.MicrosoftSQLServer2012NativeClient
					mGetProviderStr &= "SQLNCLI10"
				Case ProviderEnum.OracleDatabases
					mGetProviderStr &= "MSDAORA"
				Case ProviderEnum.SimpleTextFiles
					mGetProviderStr &= "MSDAOSP"
				Case Else
					Throw New Exception("Unknow Provider")
			End Select
			mGetProviderStr &= ";"
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mGetProviderStr", ex)
			Return ""
		End Try
	End Function

	''' <summary>
	''' 数据库类型
	''' </summary>
	Private mlngDBType As DBTypeEnum = DBTypeEnum.Unknow
	Public Property DBType() As DBTypeEnum
		Get
			Return mlngDBType
		End Get
		Set(ByVal value As DBTypeEnum)
			mlngDBType = value
		End Set
	End Property

End Class
