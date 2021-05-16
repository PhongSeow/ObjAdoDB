Module modObjAdoDBLib

	''' <summary>
	''' SQLSrvDataTypeEnum to DataTypeEnum
	''' </summary>
	Public Function GetSQLSrvAdoDataType(SQLSrvDataType As ConnSQLSrv.SQLSrvDataTypeEnum) As Field.DataTypeEnum
		Select Case SQLSrvDataType
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adBigint
				Return Field.DataTypeEnum.adBigInt
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adBinary
				Return Field.DataTypeEnum.adBinary
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adBit
				Return Field.DataTypeEnum.adBoolean
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adChar
				Return Field.DataTypeEnum.adChar
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adDate
				Return Field.DataTypeEnum.adDBDate
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adDatetime, ConnSQLSrv.SQLSrvDataTypeEnum.adDatetime2
				Return Field.DataTypeEnum.adDBTimeStamp
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adDecimal
				Return Field.DataTypeEnum.adNumeric
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adFloat
				Return Field.DataTypeEnum.adDouble
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adImage
				Return Field.DataTypeEnum.adLongVarBinary
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adInt
				Return Field.DataTypeEnum.adInteger
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adMoney
				Return Field.DataTypeEnum.adCurrency
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adNChar
				Return Field.DataTypeEnum.adWChar
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adNText
				Return Field.DataTypeEnum.adLongVarWChar
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adNumeric
				Return Field.DataTypeEnum.adNumeric
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adNVarchar
				Return Field.DataTypeEnum.adVarWChar
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adReal
				Return Field.DataTypeEnum.adSingle
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adSmallDateTime
				Return Field.DataTypeEnum.adDBTimeStamp
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adSmallInt
				Return Field.DataTypeEnum.adSmallInt
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adSmallMoney
				Return Field.DataTypeEnum.adCurrency
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adSql_Variant
				Return Field.DataTypeEnum.adVariant
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adSysname
				Return Field.DataTypeEnum.adVarWChar
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adText
				Return Field.DataTypeEnum.adLongVarChar
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adTimeStamp
				Return Field.DataTypeEnum.adBinary
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adTinyInt
				Return Field.DataTypeEnum.adUnsignedTinyInt
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adUniqueIdentifier
				Return Field.DataTypeEnum.adGUID
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adVarBinary
				Return Field.DataTypeEnum.adVarBinary
			Case ConnSQLSrv.SQLSrvDataTypeEnum.adVarChar
				Return Field.DataTypeEnum.adVarChar
			Case Else
				Return Field.DataTypeEnum.adVarChar
		End Select
	End Function

End Module
