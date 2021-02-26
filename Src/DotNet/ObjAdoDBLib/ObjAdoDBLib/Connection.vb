'**********************************
'* Name: Connection
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Connection
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.2
'* Create Time: 18/2/2021
'*1.0.2  20/2/2021   Modify mExecute,Errors
'**********************************
Public Class Connection
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.2"
	Public Obj As Object
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
		Me.Obj = CreateObject("ADODB.Connection")
	End Sub
	Public Property Attributes() As Long
		Get
			Try
				Return Me.Obj.Attributes
				Me.ClearErr()
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
	Public Function BeginTrans() As Long
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
				Me.ClearErr()
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
				Return Me.Obj.ConnectionString
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ConnectionString.Get", ex)
				Return Nothing
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
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ConnectionTimeout.Get", ex)
				Return Nothing
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
				Me.ClearErr()
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
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DefaultDatabase.Get", ex)
				Return Nothing
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
				Me.ClearErr()
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
				Me.ClearErr()
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
				Me.ClearErr()
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
				Me.ClearErr()
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
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Provider.Get", ex)
				Return Nothing
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
	Public Property State() As Long
		Get
			Try
				Return Me.Obj.State
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("State.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.State = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("State.Set", ex)
			End Try
		End Set
	End Property
End Class
