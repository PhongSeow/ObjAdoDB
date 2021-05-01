﻿'**********************************
'* Name: CmdSQLSrvSp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Command for SQL Server StoredProcedure
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.4
'* Create Time: 17/4/2021
'* 1.0.2	18/4/2021	Modify ActiveConnection
'* 1.0.3	24/4/2021	Add mAdoDataType
'* 1.0.4	25/4/2021	Modify New
'* 1.0.5	28/4/2021	Add ActiveConnection,AddPara,ParaValue,Execute
'**********************************
Public Class CmdSQLSrvSp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.5"
	Private moCommand As Command

	Public Enum SQLSrvDataTypeEnum
		adBigint = 127
		adBinary = 173
		adBit = 104
		adChar = 175
		adDate = 40
		adDatetime = 61
		adDatetime2 = 42
		'adDatetimeoffset = 43
		adDecimal = 106
		adFloat = 62
		'adGeography = 240
		'adGeometry = 240
		'adHierarchyid = 240
		adImage = 34
		adInt = 56
		adMoney = 60
		adNChar = 239
		adNText = 99
		adNumeric = 108
		adNVarchar = 231
		adReal = 59
		adSmallDateTime = 58
		adSmallInt = 52
		adSmallMoney = 122
		adSql_Variant = 98
		adSysname = 231
		adText = 35
		'adTime = 41
		adTimeStamp = 189
		adTinyInt = 48
		adUniqueIdentifier = 36
		adVarBinary = 165
		adVarChar = 167
		'adXml = 241
	End Enum

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



	''' <summary>
	''' SQLSrvDataTypeEnum to DataTypeEnum
	''' </summary>
	Private ReadOnly Property mAdoDataType(SQLSrvDataType As SQLSrvDataTypeEnum) As Field.DataTypeEnum
		Get
			Select Case SQLSrvDataType
				Case SQLSrvDataTypeEnum.adBigint
					Return Field.DataTypeEnum.adBigInt
				Case SQLSrvDataTypeEnum.adBinary
					Return Field.DataTypeEnum.adBinary
				Case SQLSrvDataTypeEnum.adBit
					Return Field.DataTypeEnum.adBoolean
				Case SQLSrvDataTypeEnum.adChar
					Return Field.DataTypeEnum.adChar
				Case SQLSrvDataTypeEnum.adDate
					Return Field.DataTypeEnum.adDBDate
				Case SQLSrvDataTypeEnum.adDatetime, SQLSrvDataTypeEnum.adDatetime2
					Return Field.DataTypeEnum.adDBTimeStamp
				Case SQLSrvDataTypeEnum.adDecimal
					Return Field.DataTypeEnum.adNumeric
				Case SQLSrvDataTypeEnum.adFloat
					Return Field.DataTypeEnum.adDouble
				Case SQLSrvDataTypeEnum.adImage
					Return Field.DataTypeEnum.adLongVarBinary
				Case SQLSrvDataTypeEnum.adInt
					Return Field.DataTypeEnum.adInteger
				Case SQLSrvDataTypeEnum.adMoney
					Return Field.DataTypeEnum.adCurrency
				Case SQLSrvDataTypeEnum.adNChar
					Return Field.DataTypeEnum.adWChar
				Case SQLSrvDataTypeEnum.adNText
					Return Field.DataTypeEnum.adLongVarWChar
				Case SQLSrvDataTypeEnum.adNumeric
					Return Field.DataTypeEnum.adNumeric
				Case SQLSrvDataTypeEnum.adNvarchar
					Return Field.DataTypeEnum.adVarWChar
				Case SQLSrvDataTypeEnum.adReal
					Return Field.DataTypeEnum.adSingle
				Case SQLSrvDataTypeEnum.adSmallDateTime
					Return Field.DataTypeEnum.adDBTimeStamp
				Case SQLSrvDataTypeEnum.adSmallInt
					Return Field.DataTypeEnum.adSmallInt
				Case SQLSrvDataTypeEnum.adSmallMoney
					Return Field.DataTypeEnum.adCurrency
				Case SQLSrvDataTypeEnum.adSql_Variant
					Return Field.DataTypeEnum.adVariant
				Case SQLSrvDataTypeEnum.adSysname
					Return Field.DataTypeEnum.adVarWChar
				Case SQLSrvDataTypeEnum.adText
					Return Field.DataTypeEnum.adLongVarChar
				Case SQLSrvDataTypeEnum.adTimeStamp
					Return Field.DataTypeEnum.adBinary
				Case SQLSrvDataTypeEnum.adTinyInt
					Return Field.DataTypeEnum.adUnsignedTinyInt
				Case SQLSrvDataTypeEnum.adUniqueIdentifier
					Return Field.DataTypeEnum.adGUID
				Case SQLSrvDataTypeEnum.adVarBinary
					Return Field.DataTypeEnum.adVarBinary
				Case SQLSrvDataTypeEnum.adVarChar
					Return Field.DataTypeEnum.adVarChar
				Case Else
					Return Field.DataTypeEnum.adVarChar
			End Select
		End Get
	End Property

	Public Function Execute() As Recordset
		Try
			Execute = moCommand.Execute(mlngRecordsAffected)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Execute", ex)
			Return Nothing
		End Try
	End Function

	Public Property ParaValue(ParaName As String) As Object
		Get
			Try
				ParaValue = moCommand.Parameters.Item(ParaName)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ParaValue.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				moCommand.ActiveConnection = value
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
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Set", ex)
			End Try
		End Set
	End Property

	Public Sub AddPara(ParaName As String, DataType As SQLSrvDataTypeEnum)
		Me.mAddPara(ParaName, DataType)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As SQLSrvDataTypeEnum, IsOutPut As Boolean)
		Me.mAddPara(ParaName, DataType,, IsOutPut)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As SQLSrvDataTypeEnum, Size As Long)
		Me.mAddPara(ParaName, DataType, Size)
	End Sub

	Public Sub AddPara(ParaName As String, DataType As SQLSrvDataTypeEnum, Size As Long, IsOutPut As Boolean)
		Me.mAddPara(ParaName, DataType, Size)
	End Sub

	Private Sub mAddPara(ParaName As String, DataType As SQLSrvDataTypeEnum, Optional Size As Long = -1, Optional IsOutPut As Boolean = False)
		Dim strStepName As String = ""
		Try
			Dim oParameter As Parameter
			Dim dyeAny As Field.DataTypeEnum
			dyeAny = Me.mAdoDataType(DataType)
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
