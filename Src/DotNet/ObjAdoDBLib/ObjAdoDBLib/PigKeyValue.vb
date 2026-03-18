'**********************************
'* Name: PigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 键值项
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 3.8
'* Create Time: 11/3/2021
'* 1.0.2	6/4/2021 Add IsKeyNameToPigMD5Force
'* 1.0.3	6/5/2021 Modify New,mNew
'* 1.0.4	8/5/2021 Modify enmValueType,New, add BytesValue,StrValue
'* 1.0.5	10/5/2021 Add BytesBase64Value
'* 1.0.6	11/5/2021 Modify StrValue
'* 1.0.7	17/5/2021 Modify New
'* 1.0.8	8/8/2021  Modify New, and Add IsExpired
'* 1.0.9	10/8/2021  Add KeyValueLen, remove mstrStrValue, modify StrValue
'* 1.0.10	11/8/2021  Add SMNameHead,SMNameBody
'* 1.0.11	13/8/2021  Rename KeyValueLen to ValueLen, add ValueMD5Bytes
'* 1.0.12	13/8/2021  Modify ValueMD5Bytes
'* 1.0.13	16/8/2021  Modify mstrSMNameBody,SMNameBody,SMNameHead
'* 1.0.14	17/8/2021  Modify New
'* 1.0.15	25/8/2021 Remove Imports PigToolsLib, change to PigToolsWinLib, and add 
'* 1.1	    29/8/2021 Chanage PigToolsWinLib to PigToolsLiteLib
'* 1.2	    2/9/2021  Add mIsValueTypeOK,ValueMD5Base64
'* 1.3	    17/9/2021  Modify ValueMD5Base64, Add Check
'* 1.4	    2/10/2021  Modify SMNameBody,SMNameHead
'* 1.5	    3/10/2021  Add CompareOther
'* 1.6	    4/10/2021  Add LastRefCacheTime,IsForceRefCache,CopyToMe
'* 1.7	    13/11/2021  Modify BytesValue,New
'* 1.8	    20/11/2021  Add OriginalBytesValue, modify BytesValue
'* 1.9	    21/11/2021  Modify New, add fBodyLen,fSaveBytesValue,InitBytesBySave,IsDataReady, Rename fCompareOther,fCopyToMe,fIsForceRefCache,fSMNameBody,KeyFileTitle
'* 1.10	    24/11/2021  Modify fBodyLen,InitBytesBySave
'* 1.11	    25/11/2021  Modify StrValue,Check
'* 2.0	    27/11/2021  Add EnmSaveType,TextType，mIsValueTypeOK,mIsTextTypeOK,mNew, and modify enmValueType,New
'* 2.1	    28/11/2021  Remove KeyFileTitle,fSMNameBody
'* 2.2	    30/11/2021  Add GetSaveData, modify ValueLen,StrValue,mInitSMNameHeadBody
'* 2.3	    2/12/2021  Add mNew for Byte, modify GetSaveData,ValueLen,fCopyToMe
'* 2.5	    5/12/2021  fInitBytesBySave rename to InitBytesBySave,fGetSaveData rename to GetSaveData
'* 2.6	    7/12/2021  Remove mNew check MD5
'* 2.7	    8/12/2021  Modify EnmValueType
'* 2.8	    9/12/2021  Add ChkMD5Type,mIsChkMD5TypeOK,HeadVersion, Modify New,Check
'* 3.0		10/12/2021 A large number of code rewrites, the interface is upgraded to 3.0, and the interface is no longer compatible with versions below 3.0
'* 3.1		11/12/2021 Add LoadBody,LoadHead,HeadData,BodyData
'* 3.2		12/12/2021 Add mIsSaveDiff, modify mRefBodyData,HeadData,LoadHead,PbValue,mNew,BodyMD5
'* 3.3		13/12/2021 Modify LoadBody,mNew,LoadHead,Check,mIsSaveTypeOK,mRefBodyData,IsMatchAnother,fIsForceRefCache,BodyData, add ValueLen
'* 3.5		28/12/2021 Modify mNew,BodyData
'* 3.6		4/1/2022 Modify StrValue
'* 3.7		26/7/2022 Imports PigToolsWinLib
'* 3.8		29/7/2022 Modify Imports
'************************************

Imports PigToolsLiteLib
Public Class PigKeyValue
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "3.8.2"
    Private Const STRU_VALUE_HEAD_VERSION As Integer = 1
    ''' <summary>
    ''' 父对象
    ''' </summary>
    ''' <returns></returns>
    Public Property Parent As PigKeyValueApp

    ''' <summary>
    ''' 键值标识，为键名的PigMD5
    ''' </summary>
    ''' <returns>键值标识</returns>
    Private mstrKeyName As String
    Public Property KeyName As String
        Get
            Return mstrKeyName
        End Get
        Friend Set(value As String)
            mstrKeyName = value
        End Set
    End Property
    Private mintHeadVersion As Integer
    Friend Property HeadVersion As Integer
        Get
            Return mintHeadVersion
        End Get
        Set(value As Integer)
            mintHeadVersion = value
        End Set
    End Property


    ''' <summary>
    ''' 检查数据MD5的方式|How to check MD5 of data
    ''' </summary>
    ''' <returns></returns>
    Private mintChkMD5Type As EnmChkMD5Type
    Public Property ChkMD5Type As EnmChkMD5Type
        Get
            Return mintChkMD5Type
        End Get
        Friend Set(value As EnmChkMD5Type)
            mintChkMD5Type = value
        End Set
    End Property

    ''' <summary>
    ''' 保存数据类型|Save data type
    ''' </summary>
    ''' <returns></returns>
    Private mintSaveType As EnmSaveType
    Public Property SaveType As EnmSaveType
        Get
            Select Case Me.ValueType
                Case EnmValueType.Text

            End Select
            Return mintSaveType
        End Get
        Friend Set(value As EnmSaveType)
            mintSaveType = value
        End Set
    End Property

    ''' <summary>
    ''' 值类型
    ''' </summary>
    ''' <returns></returns>
    Private mintValueType As EnmValueType
    Public Property ValueType As EnmValueType
        Get
            Return mintValueType
        End Get
        Friend Set(value As EnmValueType)
            mintValueType = value
        End Set
    End Property

    ''' <summary>
    ''' 文本编码类型
    ''' </summary>
    ''' <returns></returns>
    Private mintTextType As PigText.enmTextType
    Public Property TextType As PigText.enmTextType
        Get
            Return mintTextType
        End Get
        Friend Set(value As PigText.enmTextType)
            mintTextType = value
        End Set
    End Property

    ''' <summary>
    '''   过期时间|Expiration time
    ''' </summary>
    Private mdteExpTime As DateTime
    Public Property ExpTime As DateTime
        Get
            Return mdteExpTime
        End Get
        Friend Set(value As DateTime)
            mdteExpTime = value
        End Set
    End Property

    Private Function mRefBodyData() As String
        Dim LOG As New PigStepLog("mRefBodyData")
        Try
            If Me.mIsNew = False Then Throw New Exception("IsNew is False. This function is not allowed to execute")
            Me.mpbBodyData = Nothing
            Select Case Me.SaveType
                Case EnmSaveType.Original
                Case EnmSaveType.SaveSpace
                    If Me.PbValue Is Nothing Then
                        LOG.StepName = "Check PbValue"
                        Throw New Exception("Is Nothing")
                    End If
                    LOG.StepName = "New PigBytes"
                    Me.mpbBodyData = New PigBytes(Me.PbValue.Main)
                    If Me.mpbBodyData.LastErr <> "" Then
                        LOG.AddStepNameInf(Me.SaveType.ToString)
                        Throw New Exception(Me.mpbBodyData.LastErr)
                    End If
                    LOG.StepName = "Compress"
                    Me.mpbBodyData.Compress()
                    If Me.mpbBodyData.LastErr <> "" Then
                        LOG.AddStepNameInf(Me.SaveType.ToString)
                        Throw New Exception(Me.mpbBodyData.LastErr)
                    End If
                    Me.mlngBodyLen = Me.mpbBodyData.Main.Length
                    Me.mabBodyMD5 = Me.mpbBodyData.PigMD5Bytes
                Case EnmSaveType.EncSaveSpace
                    LOG.StepName = "SaveType is not supported yet"
                    Throw New Exception(Me.SaveType.ToString)
                Case Else
                    LOG.StepName = "Invalid SaveType"
                    Throw New Exception(Me.SaveType.ToString)
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private ReadOnly Property mIsBodyReady As Boolean
        Get
            Try
                If Me.BodyLen <= 0 Then Throw New Exception("Me.BodyLen <=0")
                If Me.BodyMD5.Length <> 16 Then Throw New Exception("BodyMD5.Length <> 16")
                If Me.BodyData Is Nothing Then Throw New Exception("BodyMD5BodyData Is Nothing")
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property



    ''' <summary>
    ''' 保存数据体的长度
    ''' </summary>
    Private mlngBodyLen As Long
    Public Property BodyLen As Long
        Get
            Try
                If Me.mIsSaveDiff = True Then
                    If mlngBodyLen = 0 Then
                        Dim strRet As String = Me.mRefBodyData()
                        If strRet <> "OK" Then Throw New Exception(strRet)

                    End If
                    Return mlngBodyLen
                ElseIf Me.mIsNew = True Then
                    If Me.PbValue Is Nothing Then
                        Return -1
                    ElseIf Me.mpbPbValue.Main Is Nothing Then
                        Return -1
                    Else
                        Return Me.mpbPbValue.Main.Length
                    End If
                Else
                    Return mlngBodyLen
                End If
            Catch ex As Exception
                Me.PrintDebugLog("BodyLen", ex.Message.ToString)
                Return -1
            End Try
        End Get
        Friend Set(value As Long)
            mlngBodyLen = value
        End Set
    End Property


    Private mdteDateTime As DateTime
    Public Property LastRefCacheTime As DateTime
        Get
            Return mdteDateTime
        End Get
        Friend Set(value As DateTime)
            mdteDateTime = value
        End Set
    End Property

    Public ReadOnly Property ValueLen As Long
        Get
            Try
                Select Case Me.ValueType
                    Case EnmValueType.Text
                        Return Me.StrValue.Length
                    Case EnmValueType.Bytes
                        Return Me.PbValue.Main.Length
                    Case Else
                        Return -1
                End Select
            Catch ex As Exception
                Return -2
            End Try
        End Get
    End Property

    ''' <summary>
    ''' 字符串值，非文本类型以 Base64 格式表示|String value, non text type, expressed in Base64 format
    ''' </summary>
    Private mstrValue As String = ""
    Public ReadOnly Property StrValue As String
        Get
            Dim LOG As New PigStepLog("StrValue")
            Try
                If mstrValue = "" Or mstrValue Is vbNullChar Then
                    Select Case Me.ValueType
                        Case EnmValueType.Text

                            LOG.StepName = "New PigText(Text)"
                            Dim oPigText As New PigText(Me.mpbPbValue.Main, Me.TextType)
                            If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
                            mstrValue = oPigText.Text
                            oPigText = Nothing
                        Case EnmValueType.Bytes
                            LOG.StepName = "New PigText(Bytes)"
                            If mpbPbValue Is Nothing Then Throw New Exception("PbValue Is Nothing")
                            Dim oPigText As New PigText(Me.mpbPbValue.Main)
                            If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
                            mstrValue = oPigText.Base64
                            oPigText = Nothing
                        Case Else
                            LOG.StepName = "Invalid ValueType"
                            Throw New Exception(Me.ValueType.ToString)
                    End Select
                End If
                Return mstrValue
            Catch ex As Exception
                Me.PrintDebugLog(LOG.SubName, LOG.StepName, ex.Message.ToString)
                mstrValue = ""
                Return mstrValue
            End Try
        End Get
    End Property

    'Private mlngBodyLen As Long = 0
    'Friend ReadOnly Property fBodyLen As Long
    '    Get
    '        Dim LOG.StepName As String = ""
    '        Dim LOG.Ret As String = ""
    '        Try
    '            Select Case Me.ValueType
    '                Case enmValueType.Text
    '                    If mlngBodyLen <= 0 Then mlngBodyLen = mstrValue.Length
    '                Case enmValueType.Bytes
    '                    If mlngBodyLen <= 0 Then
    '                        LOG.StepName = "fRefSaveValue"
    '                        LOG.Ret = Me.fRefSaveValue
    '                        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '                    End If
    '                Case Else
    '                    Throw New Exception("Invalid ValueType is " & Me.ValueType)
    '            End Select
    '            Return mstrValue.Length
    '        Catch ex As Exception
    '            LOG.Ret = Me.GetSubErrInf("fBodyLen", LOG.StepName, ex)
    '            Me.PrintDebugLog("As Exception", LOG.Ret)
    '            mlngBodyLen = -1
    '            Return mlngBodyLen
    '        End Try
    '    End Get
    'End Property

    'Public ReadOnly Property ValueLen As Long
    '    Get
    '        Try
    '            Select Case Me.ValueType
    '                Case EnmValueType.Bytes
    '                    Return Me.BytesValue.Length
    '                Case EnmValueType.Text
    '                    Return Me.StrValue.Length
    '                Case Else
    '                    Throw New Exception("Unsupported ValueType is " & Me.ValueType.ToString)
    '            End Select
    '        Catch ex As Exception
    '            Me.SetSubErrInf("ValueLen", ex)
    '            Return -1
    '        End Try
    '    End Get
    'End Property

    ''' <summary>
    ''' How to check data MD5
    ''' </summary>
    Public Enum EnmChkMD5Type
        ''' <summary>
        ''' Check the 1K content
        ''' </summary>
        FastCheck1K = 0
        ''' <summary>
        ''' Check the 1M content
        ''' </summary>
        FastCheck1M = 1
        ''' <summary>
        ''' Complete check
        ''' </summary>
        FullCheck = 2
    End Enum


    ''' <summary>
    ''' Value types, including text and binary
    ''' </summary>
    Public Enum EnmValueType
        ''' <summary>
        ''' text
        ''' </summary>
        Unknow = 0
        ''' <summary>
        ''' text
        ''' </summary>
        Text = 1
        ''' <summary>
        ''' Byte array
        ''' </summary>
        Bytes = 2
        ''' <summary>
        ''' Compressed byte array
        ''' </summary>
    End Enum


    ''' <summary>
    ''' Save data type, Decide whether to process the saved data.
    ''' </summary>
    Public Enum EnmSaveType
        ''' <summary>
        ''' Original, not processed
        ''' </summary>
        Original = 0
        ''' <summary>
        ''' Save space and compress and save data
        ''' </summary>
        SaveSpace = 1
        ''' <summary>
        ''' It is confidential and saves space. The data is compressed and encrypted
        ''' </summary>
        EncSaveSpace = 2
    End Enum

    '''' <summary>
    '''' 键值的PigMD5
    '''' </summary>
    '''' 
    'Public ReadOnly Property ValueMD5 As String
    '    Get
    '        Try
    '            Dim oPigMD5 As New PigMD5(mabValueMD5)
    '            If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
    '            ValueMD5 = oPigMD5.PigMD5
    '            oPigMD5 = Nothing
    '        Catch ex As Exception
    '            Me.PrintDebugLog("ValueMD5", ex.Message.ToString)
    '            Return ""
    '        End Try
    '    End Get
    'End Property


    Private mpbHeadData As PigBytes = Nothing
    Public ReadOnly Property HeadData As PigBytes
        Get
            Dim LOG As New PigStepLog("HeadData.Get")
            Try
                If Me.mpbHeadData Is Nothing Then
                    If Me.mIsNew = False Then
                        If Me.BodyData Is Nothing Then
                            Throw New Exception("BodyData is Nothing")
                        End If
                    End If
                    'Dim lngSetting As String = PbHead.GetInt64Value
                    '.HeadVersion = CInt(lngSetting / 1) Mod 10
                    '.ValueType = CInt(lngSetting / 10) Mod 10
                    '.TextType = CInt(lngSetting / 100) Mod 10
                    '.SaveType = CInt(lngSetting / 1000) Mod 10
                    '.ChkMD5Type = CInt(lngSetting / 10000) Mod 10
                    '.ExpTime = PbHead.GetDateTimeValue
                    '.BodyLen = PbHead.GetInt64Value
                    '.BodyMD5 = PbHead.GetBytesValue(16)
                    mpbHeadData = New PigBytes
                    Dim lngSetting As Long
                    lngSetting = Me.HeadVersion
                    lngSetting += Me.ValueType * 10
                    lngSetting += Me.TextType * 100
                    lngSetting += Me.SaveType * 1000
                    lngSetting += Me.ChkMD5Type * 10000
                    '顺序不可以变，对应 LoadBody
                    With Me.mpbHeadData
                        LOG.StepName = "SetValue"
                        .SetValue(lngSetting)
                        If .LastErr <> "" Then
                            LOG.AddStepNameInf("Setting")
                            Throw New Exception(.LastErr)
                        End If
                        .SetValue(Me.ExpTime)
                        If .LastErr <> "" Then
                            LOG.AddStepNameInf("ExpTime")
                            Throw New Exception(.LastErr)
                        End If
                        .SetValue(Me.BodyLen)
                        If .LastErr <> "" Then
                            LOG.AddStepNameInf("BodyLen")
                            Throw New Exception(.LastErr)
                        End If
                        .SetValue(Me.BodyMD5)
                        If .LastErr <> "" Then
                            LOG.AddStepNameInf("BodyMD5")
                            Throw New Exception(.LastErr)
                        End If
                    End With
                End If
                Return mpbHeadData
            Catch ex As Exception
                Me.PrintDebugLog(LOG.SubName, LOG.StepName, ex.Message.ToString)
                Return Nothing
            End Try
        End Get
    End Property

    ''' <summary>
    ''' 保存数据体的PigMD5
    ''' </summary>
    ''' <returns></returns>
    Private mabBodyMD5 As Byte()
    Public Property BodyMD5 As Byte()
        Get
            Try
                If Me.mIsSaveDiff = True Then
                    If mabBodyMD5.Length <> 16 Then
                        Dim strRet As String = Me.mRefBodyData()
                        If strRet <> "OK" Then Throw New Exception(strRet)
                    End If
                    Return mabBodyMD5
                Else
                    If mabBodyMD5.Length <> 16 Then
                        If Me.mIsNew = True Then
                            If Me.PbValue Is Nothing Then
                                Return Nothing
                            ElseIf Me.mpbPbValue.Main Is Nothing Then
                                Return Nothing
                            Else
                                Me.mabBodyMD5 = Me.PbValue.PigMD5Bytes
                                Return mabBodyMD5
                            End If
                        Else
                            Return mabBodyMD5
                        End If
                    Else
                        Return mabBodyMD5
                    End If
                End If
            Catch ex As Exception
                ReDim mabBodyMD5(0)
                Me.PrintDebugLog("BodyMD5", ex.Message.ToString)
                Return Nothing
            End Try
        End Get
        Friend Set(value As Byte())
            mabBodyMD5 = value
        End Set
    End Property

    ''' <summary>
    ''' 键值的 PigBytes 值，只对 ValueType 为 Bytes 有效
    ''' </summary>
    ''' <returns></returns>
    Private mpbPbValue As PigBytes
    Public ReadOnly Property PbValue As PigBytes
        Get
            Dim LOG As New PigStepLog("PbValue.Get")
            Try
                If Me.mIsNew = True Then
                    Select Case Me.ValueType
                        Case EnmValueType.Text
                            LOG.StepName = "New PigText"
                            Dim oPigText As New PigText(Me.mstrValue, Me.TextType)
                            If oPigText.LastErr <> "" Then
                                LOG.AddStepNameInf(Me.ValueType.ToString)
                                LOG.AddStepNameInf(Me.TextType.ToString)
                                Throw New Exception(oPigText.LastErr)
                            End If
                            LOG.StepName = "New PigBytes"
                            Me.mpbPbValue = New PigBytes(oPigText.TextBytes)
                            If Me.mpbPbValue.LastErr <> "" Then
                                LOG.AddStepNameInf(Me.ValueType.ToString)
                                LOG.AddStepNameInf(Me.TextType.ToString)
                                Throw New Exception(Me.mpbPbValue.LastErr)
                            End If
                            oPigText = Nothing
                            Return mpbPbValue
                        Case EnmValueType.Bytes
                            Return mpbPbValue
                        Case Else
                            Return Nothing
                    End Select
                Else
                    Return mpbPbValue
                End If
            Catch ex As Exception
                Me.PrintDebugLog(LOG.SubName, LOG.StepName, ex.Message.ToString)
                Return Nothing
            End Try
        End Get
    End Property

    ''' <summary>
    ''' 保存数据体数据
    ''' </summary>
    ''' <returns></returns>
    Private mpbBodyData As PigBytes
    Public Property BodyData As PigBytes
        Get
            Try
                If Me.mIsSaveDiff = True Then
                    If mpbBodyData Is Nothing Then
                        If Me.mIsNew = False Then
                            If Me.HeadData IsNot Nothing Then
                                If Me.mpbPbValue IsNot Nothing Then
                                    Me.mIsNew = True
                                End If
                            End If
                        End If
                        Dim strRet As String = Me.mRefBodyData()
                        If strRet <> "OK" Then Throw New Exception(strRet)
                    End If
                    Return mpbBodyData
                Else
                    Return Me.PbValue
                End If
            Catch ex As Exception
                Me.PrintDebugLog("BodyData", ex.Message.ToString)
                Return Nothing
            End Try
        End Get
        Friend Set(value As PigBytes)
            mpbBodyData = value
        End Set
    End Property

    Public ReadOnly Property BytesValue As Byte()
        Get
            Dim LOG As New PigStepLog("BytesValue")
            Try
                Select Case Me.ValueType
                    Case EnmValueType.Text, EnmValueType.Bytes
                        If Me.PbValue Is Nothing Then Throw New Exception("PbValue Is Nothing")
                        Return Me.PbValue.Main
                    Case Else
                        Throw New Exception("Invalid ValueType " & Me.ValueType.ToString)
                End Select
            Catch ex As Exception
                Me.PrintDebugLog(LOG.SubName, LOG.StepName, ex.Message.ToString)
                Return Nothing
            End Try
        End Get
    End Property

    Public Sub New(KeyName As String)
        MyBase.New(CLS_VERSION)
        Dim LOG As New PigStepLog("New")
        Try
            Me.KeyName = KeyName
            Me.mIsNew = False
            ReDim Me.BodyMD5(0)
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">文本键值</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">文本键值</param>
    ''' <param name="TextType">文本类型</param>
    ''' <param name="TextType">保存数据方式</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String, TextType As PigText.enmTextType, SaveType As PigKeyValue.EnmSaveType)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, TextType, SaveType)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">文本键值</param>
    ''' <param name="TextType">文本类型</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String, TextType As PigText.enmTextType)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, TextType)
    End Sub



    Private Sub mNew(KeyName As String, ExpTime As DateTime, KeyValue As String, Optional TextType As PigText.enmTextType = PigText.enmTextType.UTF8, Optional SaveType As EnmSaveType = EnmSaveType.Original, Optional ChkMD5Type As EnmChkMD5Type = EnmChkMD5Type.FullCheck)
        Dim LOG As New PigStepLog("mNew[Text]")
        Try
            With Me
                .KeyName = KeyName
                .ExpTime = ExpTime
                .ValueType = EnmValueType.Text
                .TextType = TextType
                .SaveType = SaveType
                .ChkMD5Type = ChkMD5Type
                .mIsNew = True
                .BodyLen = 0
                .BodyMD5 = Nothing
                .HeadVersion = STRU_VALUE_HEAD_VERSION
                ReDim Me.BodyMD5(0)
            End With
            Me.mstrValue = KeyValue
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">字节数组键值</param>
    ''' <param name="SaveType">保存数据方式</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), SaveType As EnmSaveType)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, SaveType)
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">字节数组键值</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte())
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue)
    End Sub

    Private Sub mNew(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), Optional SaveType As EnmSaveType = EnmSaveType.SaveSpace, Optional ChkMD5Type As EnmChkMD5Type = EnmChkMD5Type.FullCheck)
        Dim LOG As New PigStepLog("mNew[Byte]")
        Try
            With Me
                .KeyName = KeyName
                .ExpTime = ExpTime
                .ValueType = EnmValueType.Bytes
                .TextType = PigText.enmTextType.UnknowOrBin
                .SaveType = SaveType
                .ChkMD5Type = ChkMD5Type
                .mIsNew = True
                .BodyLen = 0
                .BodyMD5 = Nothing
                .HeadVersion = STRU_VALUE_HEAD_VERSION
                ReDim Me.BodyMD5(0)
            End With
            LOG.StepName = "New PigBytes"
            Me.mpbPbValue = New PigBytes(KeyValue)
            If Me.mpbPbValue.LastErr <> "" Then Throw New Exception(Me.mpbPbValue.LastErr)
            Me.ClearErr()
        Catch ex As Exception
            Me.mpbPbValue = Nothing
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Private Function mIsBytesMatch(ByRef SrcBytes As Byte(), ByRef MatchBytes As Byte()) As Boolean
        Try
#If NET40_OR_GREATER Then
            Return SrcBytes.SequenceEqual(MatchBytes)
#Else
            Dim i As Long
            If SrcBytes.Length <> MatchBytes.Length Then
                Return False
            Else
                mIsBytesMatch = True
                For i = 0 To SrcBytes.Length - 1
                    If SrcBytes(i) <> MatchBytes(i) Then
                        mIsBytesMatch = False
                        Exit For
                    End If
                Next
            End If

#End If
        Catch ex As Exception
            Me.SetSubErrInf("mIsBytesMatch", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 对象是否通过New创建的，是则不允许 LoadHead 和 LoadBody
    ''' </summary>
    Private mbolIsNew As Boolean
    Private Property mIsNew As Boolean
        Get
            Return mbolIsNew
        End Get
        Set(value As Boolean)
            mbolIsNew = value
        End Set
    End Property

    ''' <summary>
    ''' 是否保存的数据与原始数据不同
    ''' </summary>
    Private ReadOnly Property mIsSaveDiff As Boolean
        Get
            Select Case Me.SaveType
                Case EnmSaveType.Original
                    Return False
                Case Else
                    Return True
            End Select
        End Get
    End Property

    Private ReadOnly Property mIsValueTypeOK As Boolean
        Get
            Select Case Me.ValueType
                Case EnmValueType.Bytes, EnmValueType.Text
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Private ReadOnly Property mIsChkMD5TypeOK As Boolean
        Get
            Select Case Me.ChkMD5Type
                Case EnmChkMD5Type.FastCheck1K, EnmChkMD5Type.FastCheck1M, EnmChkMD5Type.FullCheck, EnmChkMD5Type.FullCheck
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Private ReadOnly Property mIsTextTypeOK As Boolean
        Get
            Select Case Me.ValueType
                Case EnmValueType.Bytes
                    Return False
                Case EnmValueType.Text
                    Select Case Me.TextType
                        Case PigText.enmTextType.Ascii, PigText.enmTextType.Unicode, PigText.enmTextType.UTF8
                            Return True
                        Case Else
                            Return False
                    End Select
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Private ReadOnly Property mIsSaveTypeOK As Boolean
        Get
            Select Case Me.SaveType
                Case EnmSaveType.EncSaveSpace, EnmSaveType.Original, EnmSaveType.SaveSpace
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property


    Public ReadOnly Property IsExpired As Boolean
        Get
            If Me.ExpTime < Now Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property




    ''' <summary>
    ''' 检查数据是否有效
    ''' </summary>
    ''' <returns></returns>
    Public Function Check() As String
        Dim LOG As New PigStepLog("Check")
        Try
            LOG.StepName = "Check control properties"
            Select Case Me.KeyName.Length
                Case 1 To 128
                Case Else
                    Throw New Exception("The length of the keyname must be between 1 and 128")
            End Select
            If Me.IsExpired = True Then Throw New Exception("KeyValue is IsExpired")
            If Me.mIsValueTypeOK = False Then Throw New Exception("Invalid ValueType is " & Me.ValueType.ToString)
            If Me.mIsSaveTypeOK = False Then Throw New Exception("Invalid SaveType is " & Me.SaveType.ToString)
            If Me.mIsChkMD5TypeOK = False Then Throw New Exception("Invalid ChkMD5Type is " & Me.ChkMD5Type.ToString)
            Select Case Me.ValueType
                Case EnmValueType.Text
                    If Me.mIsTextTypeOK = False Then Throw New Exception("Invalid TextType is " & Me.TextType.ToString)
                Case EnmValueType.Bytes
                    If Me.mIsNew = True Then
                        If Me.PbValue Is Nothing Then
                            Throw New Exception("The KeyValue is undefined")
                        End If
                    End If
            End Select
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function


    Friend Function fIsForceRefCache() As Boolean
        Try
            If Me.Parent Is Nothing Then
                Return False
            ElseIf Math.Abs(DateDiff(DateInterval.Second, Me.LastRefCacheTime, Now)) > Me.Parent.ForceRefCacheTime Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("fIsForceRefCache", ex)
            Return False
        End Try
    End Function



    Public Function LoadHead(PbHead As PigBytes) As String
        Dim LOG As New PigStepLog("LoadHead")
        Try
            With Me
                If .mIsNew = True Then Throw New Exception("IsNew is True. This function is not allowed to execute")
                '顺序不可以变，对应 GetHeadAndBodyBytes
                LOG.StepName = "Setting"
                Dim lngSetting As String = PbHead.GetInt64Value
                .HeadVersion = CInt(lngSetting / 1) Mod 10
                .ValueType = CInt(lngSetting / 10) Mod 10
                .TextType = CInt(lngSetting / 100) Mod 10
                .SaveType = CInt(lngSetting / 1000) Mod 10
                .ChkMD5Type = CInt(lngSetting / 10000) Mod 10
                LOG.StepName = "Set ExpTime"
                .ExpTime = PbHead.GetDateTimeValue
                If PbHead.LastErr <> "" Then Throw New Exception(PbHead.LastErr)
                LOG.StepName = "Set BodyLen"
                .BodyLen = PbHead.GetInt64Value
                If PbHead.LastErr <> "" Then Throw New Exception(PbHead.LastErr)
                LOG.StepName = "Set BodyMD5"
                ReDim Me.mabBodyMD5(0)
                Me.mabBodyMD5 = PbHead.GetBytesValue(16)
                LOG.StepName = "Check"
                LOG.Ret = .Check()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            Me.mpbHeadData = PbHead
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function LoadBody(AbBody As Byte()) As String
        Dim LOG As New PigStepLog("LoadBody")
        Try
            If Me.mIsNew = True Then Throw New Exception("IsNew is True. This function is not allowed to execute")
            If Me.HeadData Is Nothing Then Throw New Exception("HeadData not loaded.")
            If Me.BodyLen <> AbBody.Length Then Throw New Exception("Bodylen mismatch.")
            Select Case Me.ValueType
                Case EnmValueType.Text, EnmValueType.Bytes
                    LOG.StepName = "Set PbValue"
                    Me.mpbPbValue = Nothing
                    LOG.AddStepNameInf(Me.SaveType.ToString)
                    Select Case Me.SaveType
                        Case EnmSaveType.Original
                            Me.mpbPbValue = New PigBytes(AbBody)
                            If Me.mpbPbValue.LastErr <> "" Then Throw New Exception(Me.mpbPbValue.LastErr)
                        Case EnmSaveType.SaveSpace
                            Me.mpbPbValue = New PigBytes(AbBody)
                            If Me.mpbPbValue.LastErr <> "" Then Throw New Exception(Me.mpbPbValue.LastErr)
                        Case EnmSaveType.EncSaveSpace
                            Throw New Exception("Type not yet supported")
                    End Select
                    LOG.AddStepNameInf(Me.ChkMD5Type)
                    Select Case Me.ChkMD5Type
                        Case EnmChkMD5Type.FastCheck1K
                            LOG.StepName = "Type not yet supported"
                            Throw New Exception(Me.ChkMD5Type.ToString)
                        Case EnmChkMD5Type.FastCheck1M
                            LOG.StepName = "Type not yet supported"
                            Throw New Exception(Me.ChkMD5Type.ToString)
                        Case EnmChkMD5Type.FullCheck
                            If Me.mpbPbValue.IsPigMD5Mate(Me.BodyMD5) = False Then
                                Throw New Exception("BodyMD5 mismatch.")
                            End If
                        Case Else
                            LOG.StepName = "Invalid ChkMD5Type"
                            Throw New Exception(Me.ChkMD5Type.ToString)
                    End Select
                    Select Case Me.SaveType
                        Case EnmSaveType.Original
                            mstrValue = ""
                        Case EnmSaveType.SaveSpace
                            LOG.StepName = "UnCompress(SaveSpace)"
                            Me.mpbPbValue.UnCompress()
                            If Me.mpbPbValue.LastErr <> "" Then Throw New Exception(Me.mpbPbValue.LastErr)
                            mstrValue = ""
                        Case EnmSaveType.EncSaveSpace
                            LOG.StepName = "Type not yet supported"
                            Throw New Exception(Me.SaveType.ToString)
                    End Select
                Case Else
                    LOG.StepName = "Invalid ValueType"
                    Throw New Exception(Me.ValueType.ToString)
            End Select
            Return "OK"
        Catch ex As Exception
            If Me.mIsNew = False Then
                Me.mpbPbValue = Nothing
                mstrValue = ""
            End If
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function IsMatchAnother(ByRef OtherItem As PigKeyValue) As Boolean
        Dim LOG As New PigStepLog("IsMatchAnother")
        Try
            Return Me.HeadData.IsMatchBytes(OtherItem.HeadData.Main)
        Catch ex As Exception
            Me.PrintDebugLog(LOG.SubName, ex.Message.ToString)
            Return False
        End Try
    End Function

    Public Function CopyToMe(ByRef OtherItem As PigKeyValue) As String
        Dim LOG As New PigStepLog("CopyToMe")
        Try
            If Me.KeyName <> OtherItem.KeyName Then
                LOG.StepName = "Check KeyName"
                Throw New Exception(Me.KeyName & " not match " & OtherItem.KeyName)
            End If
            If Me.mIsNew = True Then
                With Me
                    .ValueType = OtherItem.ValueType
                    .SaveType = OtherItem.SaveType
                    .TextType = OtherItem.TextType
                    .ChkMD5Type = OtherItem.ChkMD5Type
                    .ExpTime = OtherItem.ExpTime
                    .BodyLen = OtherItem.BodyLen
                    .BodyMD5 = OtherItem.BodyMD5
                    .mstrValue = ""
                    .mpbPbValue = OtherItem.PbValue
                End With
            Else
                LOG.StepName = "LoadHead"
                LOG.Ret = Me.LoadHead(OtherItem.HeadData)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.KeyName)
                    Throw New Exception(LOG.Ret)
                End If
                LOG.StepName = "LoadBody"
                LOG.Ret = Me.LoadBody(OtherItem.BodyData.Main)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.KeyName)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
