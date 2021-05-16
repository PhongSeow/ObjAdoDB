'**********************************
'* Name: Command
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Command
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.5
'* Create Time: 2/3/2021
'* 1.0.2	3/3/2021	Modify ActiveConnection
'* 1.0.3	16/4/2021	Remove excess Me.ClearErr(), Modify New
'* 1.0.4	17/4/2021	Modify State
'* 1.0.5	16/5/2021	Modify ActiveConnection
'**********************************
Public Class Command
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.5"
	Public Obj As Object
	Public Enum CommandTypeEnum
		adCmdFile = 256
		adCmdStoredProc = 4
		adCmdTable = 2
		adCmdTableDirect = 512
		adCmdText = 1
		adCmdUnknown = 8
	End Enum

	Public Sub New()
		MyBase.New(CLS_VERSION)
		Try
			Me.Obj = CreateObject("ADODB.Command")
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", ex)
		End Try
	End Sub
	Public Property ActiveConnection() As Connection
		Get
			Try
				Dim oConnection As New Connection
				oConnection.Obj = Me.Obj.ActiveConnection
				If oConnection.LastErr <> "" Then Throw New Exception(oConnection.LastErr)
				Return oConnection
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Connection)
			Try
				Me.Obj.ActiveConnection = value.Obj
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Set", ex)
			End Try
		End Set
	End Property
	Public Sub Cancel()
		Try
			Me.Obj.Cancel()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Cancel", ex)
		End Try
	End Sub
	Public Property CommandStream() As Object
		Get
			Try
				Return Me.Obj.CommandStream
			Catch ex As Exception
				Me.SetSubErrInf("CommandStream.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				Me.Obj.CommandStream = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CommandStream.Set", ex)
			End Try
		End Set
	End Property
	Public Property CommandText() As String
		Get
			Try
				Return Me.Obj.CommandText
			Catch ex As Exception
				Me.SetSubErrInf("CommandText.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.CommandText = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CommandText.Set", ex)
			End Try
		End Set
	End Property
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
	Public Property CommandType() As CommandTypeEnum
		Get
			Try
				Return Me.Obj.CommandType
			Catch ex As Exception
				Me.SetSubErrInf("CommandType.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As CommandTypeEnum)
			Try
				Me.Obj.CommandType = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CommandType.Set", ex)
			End Try
		End Set
	End Property
	Public Function CreateParameter(Optional Name As String = "", Optional Type As Field.DataTypeEnum = Field.DataTypeEnum.adEmpty, Optional Direction As Parameter.ParameterDirectionEnum = Parameter.ParameterDirectionEnum.adParamInput, Optional Size As Long = -1, Optional Value As Object = Nothing) As Parameter
		Try
			Dim oParameter As New Parameter
			oParameter.Obj = Me.Obj.CreateParameter(Name, Type, Direction, Size, Value)
			Return oParameter

		Catch ex As Exception
			Me.SetSubErrInf("CreateParameter", ex)
			Return Nothing
		End Try
	End Function
	Public Property Dialect() As String
		Get
			Try
				Return Me.Obj.Dialect
			Catch ex As Exception
				Me.SetSubErrInf("Dialect.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.Dialect = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Dialect.Set", ex)
			End Try
		End Set
	End Property
	Public Function Execute(Optional RecordsAffected As Long = -1, Optional Parameters As Object = Nothing, Optional Options As Long = -1) As Recordset
		Try
			Dim oRecordset As New Recordset
			oRecordset.Obj = Me.Obj.Execute(RecordsAffected, Parameters, Options)
			Return oRecordset

		Catch ex As Exception
			Me.SetSubErrInf("Execute", ex)
			Return Nothing
		End Try
	End Function
	Public Property Name() As String
		Get
			Try
				Return Me.Obj.Name
			Catch ex As Exception
				Me.SetSubErrInf("Name.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.Name = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Name.Set", ex)
			End Try
		End Set
	End Property
	Public Property NamedParameters() As Boolean
		Get
			Try
				Return Me.Obj.NamedParameters
			Catch ex As Exception
				Me.SetSubErrInf("NamedParameters.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Boolean)
			Try
				Me.Obj.NamedParameters = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("NamedParameters.Set", ex)
			End Try
		End Set
	End Property
	Public Property Parameters() As Parameters
		Get
			Try
				Dim oParameters As New Parameters
				oParameters.Obj = Me.Obj.Parameters
				Return oParameters
			Catch ex As Exception
				Me.SetSubErrInf("Parameters.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Parameters)
			Try
				Me.Obj.Parameters = value.Obj
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Parameters.Set", ex)
			End Try
		End Set
	End Property
	Public Property Prepared() As Boolean
		Get
			Try
				Return Me.Obj.Prepared
			Catch ex As Exception
				Me.SetSubErrInf("Prepared.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Boolean)
			Try
				Me.Obj.Prepared = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Prepared.Set", ex)
			End Try
		End Set
	End Property
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
	Public ReadOnly Property State() As Long
		Get
			Try
				Return Me.Obj.State
			Catch ex As Exception
				Me.SetSubErrInf("State.Get", ex)
				Return Long.MinValue
			End Try
		End Get
	End Property
End Class
