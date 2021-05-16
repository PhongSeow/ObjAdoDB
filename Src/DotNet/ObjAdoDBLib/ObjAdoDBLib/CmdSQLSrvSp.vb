'**********************************
'* Name: CmdSQLSrvSp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Command for SQL Server StoredProcedure
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.6
'* Create Time: 17/4/2021
'* 1.0.2	18/4/2021	Modify ActiveConnection
'* 1.0.3	24/4/2021	Add mAdoDataType
'* 1.0.4	25/4/2021	Modify New
'* 1.0.5	28/4/2021	Add ActiveConnection,AddPara,ParaValue,Execute
'* 1.0.6	16/5/2021	SQLSrvDataTypeEnum move to ConnSQLSrv, Modify Execute,ParaValue,ActiveConnection
'**********************************
Public Class CmdSQLSrvSp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.6"
	Private moCommand As Command

	Public Sub New(SpName As String)
		MyBase.New(CLS_VERSION)
		Dim strStepName As String = ""
		Try
			moCommand = New Command
			With moCommand
				.CommandType = Command.CommandTypeEnum.adCmdStoredProc
				.CommandText = SpName
				Dim oParameter As Parameter
				strStepName = "CreateParameter(RETURN_VALUE)"
				oParameter = .CreateParameter("RETURN_VALUE", Field.DataTypeEnum.adInteger, Parameter.ParameterDirectionEnum.adParamReturnValue, 4)
				strStepName = "Append(RETURN_VALUE)"
				.Parameters.Append(oParameter)
				oParameter = Nothing
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", strStepName, ex)
		End Try
	End Sub

	''' <summary>
	''' Stored Procedure Name
	''' </summary>
	Private mstrSpName As String
	Public Property SpName() As String
		Get
			Return mstrSpName
		End Get
		Set(ByVal value As String)
			mstrSpName = value
		End Set
	End Property

	''' <summary>
	''' Stored Procedure return value
	''' </summary>
	Private mstrReturnValue As String
	Public ReadOnly Property ReturnValue() As Integer
		Get
			Return mstrReturnValue
		End Get
	End Property

	''' <summary>
	''' Records Affected by the execution of the Stored Procedure
	''' </summary>
	Private mlngRecordsAffected As Long
	Public ReadOnly Property RecordsAffected() As Long
		Get
			Return mlngRecordsAffected
		End Get
	End Property


	Public Function Execute() As Recordset
		Try
			Execute = moCommand.Execute(mlngRecordsAffected)
			If moCommand.LastErr <> "" Then Throw New Exception(moCommand.LastErr)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Execute", ex)
			Return Nothing
		End Try
	End Function

	Public Property ParaValue(ParaName As String) As Object
		Get
			Try
				ParaValue = moCommand.Parameters.Obj(ParaName).Value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ParaValue.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				moCommand.Parameters.Obj(ParaName).Value = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ParaValue.Set", ex)
			End Try
		End Set
	End Property

	Public Property ActiveConnection() As Connection
		Get
			Try
				Return moCommand.ActiveConnection
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Connection)
			Try
				moCommand.ActiveConnection = value
				If moCommand.LastErr <> "" Then Throw New Exception(moCommand.LastErr)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Set", ex)
			End Try
		End Set
	End Property

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum)
		Me.mAddPara(ParaName, DataType)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, IsOutPut As Boolean)
		Me.mAddPara(ParaName, DataType,, IsOutPut)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, Size As Long)
		Me.mAddPara(ParaName, DataType, Size)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, Size As Long, IsOutPut As Boolean)
		Me.mAddPara(ParaName, DataType, Size)
	End Sub

	Private Sub mAddPara(ParaName As String, DataType As ConnSQLSrv.SQLSrvDataTypeEnum, Optional Size As Long = -1, Optional IsOutPut As Boolean = False)
		Dim strStepName As String = ""
		Try
			Dim oParameter As Parameter
			Dim dyeAny As Field.DataTypeEnum
			dyeAny = GetSQLSrvAdoDataType(DataType)
			Dim pdeAny As Parameter.ParameterDirectionEnum
			If IsOutPut = True Then
				pdeAny = Parameter.ParameterDirectionEnum.adParamOutput
			Else
				pdeAny = Parameter.ParameterDirectionEnum.adParamInput
			End If
			strStepName = "Append(" & ParaName & ")"
			oParameter = moCommand.CreateParameter(ParaName, dyeAny, pdeAny, Size)
			strStepName = "Append(" & ParaName & ")"
			moCommand.Parameters.Append(oParameter)
			oParameter = Nothing
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAddPara", strStepName, ex)
		End Try
	End Sub


End Class
