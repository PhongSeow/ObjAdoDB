'**********************************
'* Name: Recordset
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 ADODB.Recordset
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.12
'* Create Time: 18/2/2021
'* 1.0.2  20/2/2021   Modify Fields
'* 1.0.3  11/3/2021   Modify NextRecordset
'* 1.0.4  18/3/2021   Add Recordset2JSon,MaxToJSonRows
'* 1.0.5  20/3/2021   Add Recordset2JSonToEnd, Modify mRecordset2JSon
'* 1.0.6  27/3/2021   Modify mRs2JSonTypeEnum,mRecordset2JSon, add Rows2JSon,IsTrimJSonValue
'* 1.0.7  27/3/2021   Modify mRs2JSonTypeEnum,Row2JSon
'* 1.0.8  4/4/2021   Remove mRecordset2JSon, Add Recordset2JSon
'* 1.0.9  16/4/2021	Remove excess Me.ClearErr()
'* 1.0.10  1/6/2021	Add AllRecordset2JSon
'* 1.0.11  2/6/2021	Modify AllRecordset2JSon
'* 1.0.12  3/6/2021	Modify Recordset2JSon,AllRecordset2JSon,NextRecordset
'**********************************
Public Class Recordset
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.12"
	Public Obj As Object
	Private moPigJSon As PigJSon
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Private Enum mRs2JSonTypeEnum
		CurrRecordsetTopRows = 10
		CurrRecordsetTopEnd = 20
		AllRecordset = 30
	End Enum

	Public Enum EditModeEnum
		adEditAdd = 2
		adEditDelete = 4
		adEditInProgress = 1
		adEditNone = 0
	End Enum
	Public Enum MarshalOptionsEnum
		adMarshalAll = 0
		adMarshalModifiedOnly = 1
	End Enum
	Public Enum SearchDirectionEnum
		adSearchBackward = -1
		adSearchForward = 1
	End Enum
	Public Enum SeekEnum
		adSeekAfter = 8
		adSeekAfterEQ = 4
		adSeekBefore = 32
		adSeekBeforeEQ = 16
		adSeekFirstEQ = 1
		adSeekLastEQ = 2
	End Enum
	Public Enum ResyncEnum
		adResyncAllValues = 2
		adResyncUnderlyingValues = 1
	End Enum
	Public Enum ParameterDirectionEnum
		adParamInput = 1
		adParamInputOutput = 3
		adParamOutput = 2
		adParamReturnValue = 4
		adParamUnknown = 0
	End Enum
	Public Enum PersistFormatEnum
		adPersistADTG = 0
		adPersistXML = 1
	End Enum
	Public Enum StringFormatEnum
		adClipString = 2
	End Enum
	Public Enum LockTypeEnum
		adLockBatchOptimistic = 4
		adLockOptimistic = 3
		adLockPessimistic = 2
		adLockReadOnly = 1
	End Enum
	Public Enum CursorTypeEnum
		adOpenDynamic = 2
		adOpenForwardOnly = 0
		adOpenKeyset = 1
		adOpenStatic = 3
	End Enum

	Public Enum CompareEnum
		adCompareEqual = 1
		adCompareGreaterThan = 2
		adCompareLessThan = 0
		adCompareNotComparable = 4
		adCompareNotEqual = 3
	End Enum
	Public Enum CursorLocationEnum
		adUseClient = 3
		adUseServer = 2
	End Enum
	Public Enum AffectEnum
		adAffectAllChapters = 4
		adAffectCurrent = 1
		adAffectGroup = 2
	End Enum
	Public Enum PositionEnum
		adPosBOF = -2
		adPosEOF = -3
		adPosUnknown = -1
	End Enum

	''' <summary>
	''' Whether to remove the space before and after the value is converted to JSON
	''' </summary>
	Private mbolIsTrimJSonValue As Boolean = True
	Public Property IsTrimJSonValue() As Boolean
		Get
			Return mbolIsTrimJSonValue
		End Get
		Set(ByVal value As Boolean)
			mbolIsTrimJSonValue = value
		End Set
	End Property

	''' <summary>
	''' The maximum number of rows to convert the Recordset to JSON
	''' </summary>
	Private mlngMaxToJSonRows As Long = 1024
	Public Property MaxToJSonRows() As Long
		Get
			Return mlngMaxToJSonRows
		End Get
		Set(ByVal value As Long)
			mlngMaxToJSonRows = value
		End Set
	End Property

	'''' <summary>
	'''' Has Next Recordset
	'''' </summary>
	'Private mbolHasNextRecordset As Boolean
	'Public ReadOnly Property HasNextRecordset() As Boolean
	'	Get
	'		Return mbolHasNextRecordset
	'	End Get
	'End Property

	Public Property AbsolutePage() As PositionEnum
		Get
			Try
				Return Me.Obj.AbsolutePage
			Catch ex As Exception
				Me.SetSubErrInf("AbsolutePage.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As PositionEnum)
			Try
				Me.Obj.AbsolutePage = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("AbsolutePage.Set", ex)
			End Try
		End Set
	End Property
	Public Property AbsolutePosition() As PositionEnum
		Get
			Try
				Return Me.Obj.AbsolutePosition
			Catch ex As Exception
				Me.SetSubErrInf("AbsolutePosition.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As PositionEnum)
			Try
				Me.Obj.AbsolutePosition = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("AbsolutePosition.Set", ex)
			End Try
		End Set
	End Property
	Public ReadOnly Property ActiveCommand() As Object
		Get
			Try
				Return Me.Obj.ActiveCommand
			Catch ex As Exception
				Me.SetSubErrInf("ActiveCommand.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Property ActiveConnection() As Object
		Get
			Try
				Return Me.Obj.ActiveConnection
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				Me.Obj.ActiveConnection = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ActiveConnection.Set", ex)
			End Try
		End Set
	End Property
	Public Sub AddNew(Optional FieldList = "", Optional Values = "")
		Try
			Me.Obj.AddNew(FieldList, Values)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("AddNew", ex)
		End Try
	End Sub
	Public ReadOnly Property BOF() As Boolean
		Get
			Try
				Return Me.Obj.BOF
			Catch ex As Exception
				Me.SetSubErrInf("BOF.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Property Bookmark() As Object
		Get
			Try
				Return Me.Obj.Bookmark
			Catch ex As Exception
				Me.SetSubErrInf("Bookmark.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				Me.Obj.Bookmark = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Bookmark.Set", ex)
			End Try
		End Set
	End Property
	Public Property CacheSize() As Long
		Get
			Try
				Return Me.Obj.CacheSize
			Catch ex As Exception
				Me.SetSubErrInf("CacheSize.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.CacheSize = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CacheSize.Set", ex)
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
	Public Sub CancelBatch(Optional AffectRecords As AffectEnum = AffectEnum.adAffectAllChapters)
		Try
			Me.Obj.CancelBatch(AffectRecords)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CancelBatch", ex)
		End Try
	End Sub
	Public Sub CancelUpdate()
		Try
			Me.Obj.CancelUpdate()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CancelUpdate", ex)
		End Try
	End Sub
	Public Function Clone(Optional LockType As LockTypeEnum = LockTypeEnum.adLockReadOnly) As Recordset
		Try
			Clone = Me.Obj.Clone(LockType)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Clone", ex)
			Return Nothing
		End Try
	End Function
	Public Sub Close()
		Try
			Me.Obj.Close()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Close", ex)
		End Try
	End Sub
	Public Function CompareBookmarks(Bookmark1, Bookmark2) As CompareEnum
		Try
			CompareBookmarks = Me.Obj.CompareBookmarks(Bookmark1, Bookmark2)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CompareBookmarks", ex)
			Return Nothing
		End Try
	End Function
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

	Public Property CursorType() As CursorTypeEnum
		Get
			Try
				Return Me.Obj.CursorType
			Catch ex As Exception
				Me.SetSubErrInf("CursorType.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As CursorTypeEnum)
			Try
				Me.Obj.CursorType = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CursorType.Set", ex)
			End Try
		End Set
	End Property
	Public Property DataMember() As String
		Get
			Try
				Return Me.Obj.DataMember
			Catch ex As Exception
				Me.SetSubErrInf("DataMember.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.DataMember = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DataMember.Set", ex)
			End Try
		End Set
	End Property
	Public Property DataSource() As Object
		Get
			Try
				Return Me.Obj.DataSource
			Catch ex As Exception
				Me.SetSubErrInf("DataSource.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				Me.Obj.DataSource = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DataSource.Set", ex)
			End Try
		End Set
	End Property
	Public Sub Delete(Optional AffectRecords As AffectEnum = AffectEnum.adAffectCurrent)
		Try
			Me.Obj.Delete(AffectRecords)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Delete", ex)
		End Try
	End Sub
	Public ReadOnly Property EditMode() As EditModeEnum
		Get
			Try
				Return Me.Obj.EditMode
			Catch ex As Exception
				Me.SetSubErrInf("EditMode.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property EOF() As Boolean
		Get
			Try
				Return Me.Obj.EOF
			Catch ex As Exception
				Me.SetSubErrInf("EOF.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Fields() As Fields
		Get
			Try
				Dim oFields As New Fields
				oFields.Obj = Me.Obj.Fields
				Return oFields
			Catch ex As Exception
				Me.SetSubErrInf("Fields.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Property Filter() As Object
		Get
			Try
				Return Me.Obj.Filter
			Catch ex As Exception
				Me.SetSubErrInf("Filter.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				Me.Obj.Filter = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Filter.Set", ex)
			End Try
		End Set
	End Property

	Public Sub Find(Criteria As String)
		Me.Find(Criteria)
	End Sub

	Public Sub Find(Criteria As String, Optional SkipRecords As Long = 0, Optional SearchDirection As SearchDirectionEnum = SearchDirectionEnum.adSearchForward, Optional Start As Object = Nothing)
		Me.Find(Criteria, SkipRecords, SearchDirection, Start)
	End Sub

	Private Sub mFind(Criteria As String, Optional SkipRecords As Long = 0, Optional SearchDirection As SearchDirectionEnum = SearchDirectionEnum.adSearchForward, Optional Start As Object = Nothing)
		Try
			If Start Is Nothing Then
				Me.Obj.Find(Criteria, SkipRecords, SearchDirection)
			Else
				Me.Obj.Find(Criteria, SkipRecords, SearchDirection, Start)
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mFind", ex)
		End Try
	End Sub

	Public Function GetRows(Optional Rows As Long = -1, Optional Start As Long = 0, Optional Fields As Long = 0) As Object
		Try
			Return Me.Obj.GetRows(Rows = -1, Start, Fields)
		Catch ex As Exception
			Me.SetSubErrInf("GetRows", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetString(Optional StringFormat As StringFormatEnum = StringFormatEnum.adClipString, Optional NumRows As Long = -1, Optional ColumnDelimeter As String = "", Optional RowDelimeter As String = "", Optional NullExpr As String = "") As String
		Try
			Return Me.Obj.GetString(StringFormat, NumRows, ColumnDelimeter, RowDelimeter, NullExpr)
		Catch ex As Exception
			Me.SetSubErrInf("GetString", ex)
			Return Nothing
		End Try
	End Function
	Public Property Index() As String
		Get
			Try
				Return Me.Obj.Index
			Catch ex As Exception
				Me.SetSubErrInf("Index.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.Index = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Index.Set", ex)
			End Try
		End Set
	End Property
	Public Property LockType() As LockTypeEnum
		Get
			Try
				Return Me.Obj.LockType
			Catch ex As Exception
				Me.SetSubErrInf("LockType.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As LockTypeEnum)
			Try
				Me.Obj.LockType = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("LockType.Set", ex)
			End Try
		End Set
	End Property
	Public Property MarshalOptions() As MarshalOptionsEnum
		Get
			Try
				Return Me.Obj.MarshalOptions
			Catch ex As Exception
				Me.SetSubErrInf("MarshalOptions.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As MarshalOptionsEnum)
			Try
				Me.Obj.MarshalOptions = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("MarshalOptions.Set", ex)
			End Try
		End Set
	End Property
	Public Property MaxRecords() As Long
		Get
			Try
				Return Me.Obj.MaxRecords
			Catch ex As Exception
				Me.SetSubErrInf("MaxRecords.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.MaxRecords = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("MaxRecords.Set", ex)
			End Try
		End Set
	End Property
	Public Sub Move(NumRecords As Long, Optional Start As Long = 0)
		Try
			Me.Obj.Move(NumRecords, Start)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Move", ex)
		End Try
	End Sub
	Public Sub MoveFirst()
		Try
			Me.Obj.MoveFirst()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("MoveFirst", ex)
		End Try
	End Sub
	Public Sub MoveLast()
		Try
			Me.Obj.MoveLast()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("MoveLast", ex)
		End Try
	End Sub
	Public Sub MoveNext()
		Try
			Me.Obj.MoveNext()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("MoveNext", ex)
		End Try
	End Sub
	Public Sub MovePrevious()
		Try
			Me.Obj.MovePrevious()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("MovePrevious", ex)
		End Try
	End Sub
	Public Function NextRecordset(Optional RecordsAffected = "") As Recordset
		Try
			Dim oRecordset As New Recordset
			oRecordset.Obj = Me.Obj.NextRecordset(RecordsAffected)
			If oRecordset.Obj Is Nothing Then
				Throw New Exception("Not NextRecordset")
			End If
			NextRecordset = oRecordset
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("NextRecordset", ex)
			Return Nothing
		End Try
	End Function
	Public Sub Open(Optional Source As String = "", Optional ActiveConnection As String = "", Optional CursorType As CursorTypeEnum = CursorTypeEnum.adOpenForwardOnly, Optional LockType As LockTypeEnum = LockTypeEnum.adLockReadOnly, Optional Options As Long = -1)
		Try
			Me.Obj = Nothing
			Me.Obj = CreateObject("ADODB.Recordset")
			Me.Obj.Open(Source, ActiveConnection, CursorType, LockType, Options)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Open", ex)
		End Try
	End Sub
	Public ReadOnly Property PageCount() As Long
		Get
			Try
				Return Me.Obj.PageCount
			Catch ex As Exception
				Me.SetSubErrInf("PageCount.Get", ex)
				Return 0
			End Try
		End Get
	End Property
	Public Property PageSize() As Long
		Get
			Try
				Return Me.Obj.PageSize
			Catch ex As Exception
				Me.SetSubErrInf("PageSize.Get", ex)
				Return 0
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.PageSize = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("PageSize.Set", ex)
			End Try
		End Set
	End Property
	Public ReadOnly Property Properties() As Properties
		Get
			Try
				Return Me.Obj.Properties
			Catch ex As Exception
				Me.SetSubErrInf("Properties.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property RecordCount() As Long
		Get
			Try
				Return Me.Obj.RecordCount
			Catch ex As Exception
				Me.SetSubErrInf("RecordCount.Get", ex)
				Return 0
			End Try
		End Get
	End Property
	Public Sub Requery(Optional Options As Long = -1)
		Try
			Me.Obj.Requery(Options)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Requery", ex)
		End Try
	End Sub
	Public Sub Resync(Optional AffectRecords As AffectEnum = AffectEnum.adAffectAllChapters, Optional ResyncValues As ResyncEnum = ResyncEnum.adResyncAllValues)
		Try
			Me.Obj.Resync(AffectRecords, ResyncValues)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Resync", ex)
		End Try
	End Sub
	Public Sub Save(Optional Destination As String = "", Optional PersistFormat As PersistFormatEnum = PersistFormatEnum.adPersistADTG)
		Try
			Me.Obj.Save(Destination, PersistFormat)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Save", ex)
		End Try
	End Sub
	Public Sub Seek(KeyValues As Object, Optional SeekOption As SeekEnum = SeekEnum.adSeekFirstEQ)
		Try
			Me.Obj.Seek(KeyValues, SeekOption)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Seek", ex)
		End Try
	End Sub
	Public Property Sort() As String
		Get
			Try
				Return Me.Obj.Sort
			Catch ex As Exception
				Me.SetSubErrInf("Sort.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.Sort = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Sort.Set", ex)
			End Try
		End Set
	End Property
	Public Property Source() As Object
		Get
			Try
				Return Me.Obj.Source
			Catch ex As Exception
				Me.SetSubErrInf("Source.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				Me.Obj.Source = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Source.Set", ex)
			End Try
		End Set
	End Property
	Public ReadOnly Property State() As Long
		Get
			Try
				Return Me.Obj.State
			Catch ex As Exception
				Me.SetSubErrInf("State.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Status() As Long
		Get
			Try
				Return Me.Obj.Status
			Catch ex As Exception
				Me.SetSubErrInf("Status.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Property StayInSync() As Boolean
		Get
			Try
				Return Me.Obj.StayInSync
			Catch ex As Exception
				Me.SetSubErrInf("StayInSync.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Boolean)
			Try
				Me.Obj.StayInSync = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("StayInSync.Set", ex)
			End Try
		End Set
	End Property

	Public Sub Update(Optional Fields = "", Optional Values = "")
		Try
			Me.Obj.Update(Fields, Values)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Update", ex)
		End Try
	End Sub
	Public Sub UpdateBatch(Optional AffectRecords As AffectEnum = AffectEnum.adAffectAllChapters)
		Try
			Me.Obj.UpdateBatch(AffectRecords)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("UpdateBatch", ex)
		End Try
	End Sub

	''' <summary>
	''' Convert current row to JSON|当前行转换成JSON
	''' </summary>
	Public Function Row2JSon() As String
		Try
			Dim pjMain As New PigJSon
			With pjMain
				If Me.EOF = False Then
					For i = 0 To Me.Fields.Count - 1
						Dim oField As Field = Me.Fields.Item(i)
						Dim strName As String = oField.Name
						Dim strValue As String = oField.ValueForJSon
						If strName = "" Then strName = "Col" & (i + 1).ToString
						If Me.IsTrimJSonValue = True Then strValue = Trim(strValue)
						If i = 0 Then
							.AddEle(strName, strValue, True)
						Else
							.AddEle(strName, strValue)
						End If
					Next
					.AddSymbol(PigJSon.xpSymbolType.EleEndFlag)
				End If
			End With
			Row2JSon = pjMain.MainJSonStr
			pjMain = Nothing
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Row2JSon", ex)
			Return ""
		End Try
	End Function

	''' <summary>
	''' Convert all recordset to JSON|所有结果集转换成JSON
	''' </summary>
	''' <returns></returns>
	Public Function AllRecordset2JSon() As String
		Dim strStepName As String = ""
		Try
			Dim intRSNo As Integer = 0
			strStepName = "New PigJSon"
			Dim pjMain As New PigJSon
			If pjMain.LastErr <> "" Then Throw New Exception(pjMain.LastErr)
			pjMain.AddArrayEleBegin("RS", True)
			Dim strRsJSon As String
			strStepName = "Me.Recordset2JSon"
			strRsJSon = Me.Recordset2JSon(Me.MaxToJSonRows)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			pjMain.AddArrayEleValue(strRsJSon, True)
			intRSNo = 1
			strStepName = "Me.NextRecordset"
			Dim rsParent As Recordset = Nothing
			Dim rsSub As Recordset = Me.NextRecordset
			Do While Not rsSub Is Nothing
				strStepName = "rs.Recordset2JSon"
				strRsJSon = rsSub.Recordset2JSon(Me.MaxToJSonRows)
				If rsSub.LastErr <> "" Then Throw New Exception(rsSub.LastErr)
				pjMain.AddArrayEleValue(strRsJSon)
				intRSNo += 1
				rsParent = rsSub
				strStepName = "rs.NextRecordset"
				rsSub = Nothing
				rsSub = rsParent.NextRecordset
				If rsParent.LastErr <> "" Then Exit Do
			Loop
			pjMain.AddSymbol(PigJSon.xpSymbolType.ArrayEndFlag)
			pjMain.AddEle("TotalRS", intRSNo)
			pjMain.AddSymbol(PigJSon.xpSymbolType.EleEndFlag)
			AllRecordset2JSon = pjMain.MainJSonStr
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("AllRecordset2JSon", ex)
			Return ""
		End Try
	End Function


	''' <summary>
	''' Convert current recordset to JSON|当前结果集转换成JSON
	''' </summary>
	''' <param name="TopRows">Top rows|最前行数</param>
	''' <returns></returns>
	Public Function Recordset2JSon(TopRows As Long) As String
		Dim strStepName As String = ""
		Try
			Dim intRowNo As Integer = 0
			strStepName = "New PigJSon"
			Dim pjMain As New PigJSon
			If pjMain.LastErr <> "" Then Throw New Exception(pjMain.LastErr)
			pjMain.AddArrayEleBegin("ROW", True)
			Do While Not Me.EOF
				If intRowNo >= Me.MaxToJSonRows Then Exit Do
				If intRowNo >= TopRows Then Exit Do
				strStepName = "Row2JSon"
				Dim strRowJSon As String = Me.Row2JSon
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
				If intRowNo = 0 Then
					pjMain.AddArrayEleValue(strRowJSon, True)
				Else
					pjMain.AddArrayEleValue(strRowJSon)
				End If
				intRowNo += 1
				strStepName = "MoveNext"
				Me.MoveNext()
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Loop
			'			If intRowNo > 0 Then pjMain.AddSymbol(PigJSon.xpSymbolType.ArrayEndFlag)
			pjMain.AddSymbol(PigJSon.xpSymbolType.ArrayEndFlag)
			pjMain.AddEle("TotalRows", intRowNo)
			pjMain.AddEle("IsEOF", Me.EOF)
			pjMain.AddSymbol(PigJSon.xpSymbolType.EleEndFlag)
			Recordset2JSon = pjMain.MainJSonStr
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Recordset2JSon", ex)
			Return ""
		End Try
	End Function

	'Public Function Recordset2JSonToEnd() As String
	'	Try
	'		Recordset2JSonToEnd = Me.mRecordset2JSon(mRs2JSonTypeEnum.CurrRecordsetTopEnd).MainJSonStr
	'		Me.ClearErr()
	'	Catch ex As Exception
	'		Me.SetSubErrInf("Recordset2JSonToEnd", ex)
	'		Return ""
	'	End Try
	'End Function


	'Private Function mRecordset2JSon(Rs2JSonType As mRs2JSonTypeEnum, Optional TopRows As Long = 1) As PigJSon
	'	Dim strStepName As String = ""
	'	Try
	'		Dim pjMain As New PigJSon, intRows As Integer = 0
	'		Select Case Rs2JSonType
	'			Case mRs2JSonTypeEnum.CurrRecordsetTopEnd, mRs2JSonTypeEnum.CurrRecordsetTopRows
	'				Dim pjRow As New PigJSon
	'				With pjMain
	'					.Reset()
	'					.AddArrayEleBegin("RowsValueList", True)
	'					Do While True
	'						If intRows > Me.MaxToJSonRows Then Exit Do
	'						If Me.EOF = True Then Exit Do
	'						intRows += 1
	'						With pjRow
	'							.Reset()
	'							For i = 0 To Me.Fields.Count - 1
	'								Dim oField As Field = Me.Fields.Item(i)
	'								If i = 0 Then
	'									.AddEle(oField.Name, oField.ValueForJSon, True)
	'								Else
	'									.AddEle(oField.Name, oField.ValueForJSon)
	'								End If
	'							Next
	'							.AddSymbol(PigJSon.xpSymbolType.EleEndFlag)
	'						End With
	'						.AddArrayEleValue(pjRow.MainJSonStr)
	'					Loop
	'					.AddSymbol(PigJSon.xpSymbolType.ArrayEndFlag)
	'					.AddEle("Rows", intRows)
	'					.AddSymbol(PigJSon.xpSymbolType.EleEndFlag)
	'				End With
	'				mRecordset2JSon = pjMain
	'			Case mRs2JSonTypeEnum.AllRecordset
	'				Throw New Exception("Coming soon")
	'			Case Else
	'				Throw New Exception("Invalid Rs2JSonType")
	'		End Select
	'		Me.ClearErr()
	'	Catch ex As Exception
	'		Me.SetSubErrInf("mRecordset2JSon", ex)
	'		Return Nothing
	'	End Try
	'End Function

End Class


