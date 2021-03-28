'**********************************
'* Name: Field
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Fields
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.4
'* Create Time: 21/2/2021
'*1.0.2  19/3/2021   Add BooleanValue,DateValue,DecValue,IntValue,LngValue,StrValue,DataCategory
'*1.0.3  20/3/2021   Modify DecValue, add ValueForJSon
'*1.0.4  27/3/2021   Modify StrValue,LngValue
'**********************************
Public Class Field
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.4"
	Public Obj As Object

	Public Enum DataCategoryEnum
		OtherValue = 0
		StrValue = 10
		IntValue = 20
		DecValue = 30
		BooleanValue = 40
		DateValue = 50
	End Enum

	Public Enum DataTypeEnum
		adArray = 8192
		adBigInt = 20
		adBinary = 128
		adBoolean = 11
		adBSTR = 8
		adChapter = 136
		adChar = 129
		adCurrency = 6
		adDate = 7
		adDBDate = 133
		adDBTime = 134
		adDBTimeStamp = 135
		adDecimal = 14
		adDouble = 5
		adEmpty = 0
		adError = 10
		adFileTime = 64
		adGUID = 72
		adIDispatch = 9
		adInteger = 3
		adIUnknown = 13
		adLongVarBinary = 205
		adLongVarChar = 201
		adLongVarWChar = 203
		adNumeric = 131
		adPropVariant = 138
		adSingle = 4
		adSmallInt = 2
		adTinyInt = 16
		adUnsignedBigInt = 21
		adUnsignedInt = 19
		adUnsignedSmallInt = 18
		adUnsignedTinyInt = 17
		adUserDefined = 132
		adVarBinary = 204
		adVarChar = 200
		adVariant = 12
		adVarNumeric = 139
		adVarWChar = 202
		adWChar = 130
	End Enum

	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub

	Public ReadOnly Property ActualSize() As Long
		Get
			Try
				Return Me.Obj.ActualSize
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ActualSize.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property Attributes() As Long
		Get
			Try
				Return Me.Obj.Attributes
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Attributes.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property DefinedSize() As Long
		Get
			Try
				Return Me.Obj.DefinedSize
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DefinedSize.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property Name() As String
		Get
			Try
				Return Me.Obj.Name
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Name.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property NumericScale() As Byte
		Get
			Try
				Return Me.Obj.NumericScale
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("NumericScale.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property Precision() As Byte
		Get
			Try
				Return Me.Obj.Precision
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Precision.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property Status() As Long
		Get
			Try
				Return Me.Obj.Status
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Status.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property DataCategory() As DataCategoryEnum
		Get
			Try
				Select Case Me.Type
					Case DataTypeEnum.adChar, DataTypeEnum.adVarChar, DataTypeEnum.adWChar, DataTypeEnum.adVarWChar, DataTypeEnum.adLongVarChar, DataTypeEnum.adLongVarWChar, DataTypeEnum.adGUID
						DataCategory = DataCategoryEnum.StrValue
					Case DataTypeEnum.adBigInt, DataTypeEnum.adTinyInt, DataTypeEnum.adSmallInt, DataTypeEnum.adUnsignedBigInt, DataTypeEnum.adUnsignedInt, DataTypeEnum.adUnsignedSmallInt, DataTypeEnum.adUnsignedTinyInt, DataTypeEnum.adInteger
						DataCategory = DataCategoryEnum.IntValue
					Case DataTypeEnum.adChar, DataTypeEnum.adVarChar, DataTypeEnum.adWChar, DataTypeEnum.adVarWChar, DataTypeEnum.adLongVarChar, DataTypeEnum.adLongVarWChar, DataTypeEnum.adGUID
						DataCategory = DataCategoryEnum.DecValue
					Case Else
						DataCategory = DataCategoryEnum.OtherValue
				End Select
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DataCategory.Get", ex)
				Return DataCategoryEnum.OtherValue
			End Try
		End Get
	End Property

	Public ReadOnly Property Type() As DataTypeEnum
		Get
			Try
				Return Me.Obj.Type
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Type.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property DecValue() As Decimal
		Get
			Try
				Return CDec(Me.Obj.Value)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DecValue.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property DateValue() As DateTime
		Get
			Try
				Return CDate(Me.Obj.Value)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DateValue.Get", ex)
				Return DateTime.MinValue
			End Try
		End Get
	End Property

	Public ReadOnly Property StrValue() As String
		Get
			Try
				Return CStr(Me.Obj.Value)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("StrValue.Get", ex)
				Return ""
			End Try
		End Get
	End Property

	Public ReadOnly Property LngValue() As Long
		Get
			Try
				Return CLng(Me.Obj.Value)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("LngValue.Get", ex)
				Return 0
			End Try
		End Get
	End Property

	Public ReadOnly Property IntValue() As Integer
		Get
			Try
				Return CBool(Me.Obj.Value)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("IntValue.Get", ex)
				Return 0
			End Try
		End Get
	End Property

	Public ReadOnly Property BooleanValue() As Boolean
		Get
			Try
				Return CBool(Me.Obj.Value)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("BooleanValue.Get", ex)
				Return False
			End Try
		End Get
	End Property

	Friend ReadOnly Property ValueForJSon() As Object
		Get
			Try
				Select Case Me.DataCategory
					Case DataCategoryEnum.BooleanValue
						ValueForJSon = Me.BooleanValue
					Case DataCategoryEnum.DateValue
						ValueForJSon = Me.DateValue
					Case DataCategoryEnum.DecValue
						ValueForJSon = Me.DecValue
					Case DataCategoryEnum.IntValue
						ValueForJSon = Me.IntValue
					Case DataCategoryEnum.OtherValue
						ValueForJSon = Me.StrValue
					Case DataCategoryEnum.StrValue
						ValueForJSon = Me.StrValue
					Case Else
						ValueForJSon = Me.StrValue
				End Select
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ValueForJSon.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property Value() As Object
		Get
			Try
				Return Me.Obj.Value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Value.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

End Class
