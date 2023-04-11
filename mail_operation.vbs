Option Explicit


Function EscapeStrForCSV(ByVal CSVStr)
  'CSV�ł̓_�u���N�I�[�e�[�V�����́u�J���}�𖳎�����v�Ƃ������}�ł���̂�
  '������A�u�_�u���N�I�[�e�[�V�������̂��̂ł���v���Ƃ������ɂ�,�_�u���N�I�[�e�[�V����2���K�v
  EscapeStrForCSV=Replace(CSVStr,"""","""""")
  If InStr(EscapeStrForCSV,",") > 0 Then
    '�܂�������Ɂu,(�J���})�v���܂܂�Ă�����,�����,�u������̃J���}�v�ł���Ƃ������Ƃ��������߂ɋt�ɂ��̕�������_�u���N�I�[�e�[�V�����ł�����
    EscapeStrForCSV=""""&EscapeStrForCSV&""""
  End If
End Function

'CSV�p�̃G�X�P�[�v������
Function DeEscapeStrForCSV(ByVal CSVStr)
  '�_�u���N�I�[�e�[�V�����������Ƃ���,1�̓_�u���N�I�[�e�[�V�������̂��̂ł���1�̓_�u���N�I�[�e�[�V�����ɕ�����ł���Ɩ��߂������
  DeEscapeStrForCSV=Replace(CSVStr,"""""","""")
  If InStr(DeEscapeStrForCSV,",") > 0 Then
     DeEscapeStrForCSV=Left(DeEscapeStrForCSV,Len(EscapeStrForCSV)-Len(""""))
    DeEscapeStrForCSV=Right(DeEscapeStrForCSV,Len(EscapeStrForCSV)-Len(""""))
  End If
End Function

'�������̕������,�������̕��������؂蕶���Ƃ��Ĕz��ɕ������邪(�����܂ł͕W����split�֐��Ɠ���)
'��O�����̕�����ɋ��܂ꂽ�Ƃ��͑������̕����񂪗��Ă���؂蕶���Ƃ݂Ȃ��Ȃ��悤�ɂ�����
'������ł����s�蒷�̌�ǂݐ�ǂ݂����AVBS�ɂ͂���ȋ@�\�͂Ȃ��̂Ŏ��O�Ŏ���
'�Ⴆ��,CSV�ł́u,�v���f�[�^�Ԃ̋�؂蕶��(�Z��)�Ƃ��Ă��邪�A�u"(�_�u���N�I�[�e�[�V����)�v�̒��ɋ��܂ꂽ�J���}�͋�؂蕶���Ƃ��Ĉ��������Ȃ��B����Ȏ��Ɏg��
Function SplitExceptEscapeChar(ByVal OneStr,ByVal SplitChar,ByVal EscapeChar)
  Dim Result ()
  Dim Last
  Last=0
  Dim IsInsideOfEscapeChar
  IsInsideOfEscapeChar=False
  
  Dim SplitStartIndex
  SplitStartIndex=1
  
  Dim i
  For i=1 To Len(OneStr)
    Dim CurrentChar
    CurrentChar=Mid(OneStr,i,1)
    If CurrentChar = SplitChar Then
      If Not IsInsideOfEscapeChar Then
          ReDim Preserve Result(Last)
          Result(Last)=Mid(OneStr,SplitStartIndex,i-SplitStartIndex)
          SplitStartIndex=i+1
          Last=Last+1
      End If
    ElseIf CurrentChar = EscapeChar Then
       IsInsideOfEscapeChar=Not IsInsideOfEscapeChar
    End If
  Next
  
  ReDim Preserve Result(Last)
  Result(Last)=Mid(OneStr,SplitStartIndex,Len(OneStr)-SplitStartIndex+1)
  SplitExceptEscapeChar=Result
  
End Function       

Class AddrNumSet

  Public DstEmailAddress
  Public DstName
  Public FirstDate
  Public CumulativeMailNum
  Public CurrentSavedMailNum
  Public CurrentTmpMailNum
  Public State
  
  Public Sub Class_Initialize()
    DstEmailAddress=""
    DstName=""
    CurrentSavedMailNum=0
    CurrentTmpMailNum=0
    State="�폜�ς݂ֈړ�"
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  Public Function SetValue(ByVal Addr,ByVal Name,ByVal NumStr,ByVal MailState)
  
   DstEmailAddress=Addr
   DstName=Name
   CumulativeMailNum=CLng(NumStr)
   State=MailState
   
  End Function
  
  Public Function AddDate(ByVal FDate)
    If IsEmpty(FirstDate) Then
      FirstDate=FDate
    ElseIf FDate < FirstDate Then
      FirstDate=FDate
    End If
  End Function
  
  
  Public Function GetMailNum(ByVal ItemName)
    Select Case ItemName
      Case "ReceivedFolder"
        GetMailNum=CurrentSavedMailNum
        Exit Function
      Case "DeletedFolder"
        GetMailNum=CurrentTmpMailNum
        Exit Function
      Case "Exists"
        GetMailNum=CurrentSavedMailNum+CurrentTmpMailNum
        Exit Function
      Case Else
        GetMailNum=CumulativeMailNum
    End Select
    
  End Function
  
  '�V�����J�E���g
  '�������͈ȑO�ɂ��̃��[���𐔂������Ƃ����邩�ǂ���(���̃��[�����O��J�E���g���̃��[���̍ŐV�̓�������ł��邱��)�������t���O���ݐσ��[�����𐔂���ۂɕK�v(�������[�����ɂ͕K�v�Ȃ�)
  '��������,���݂��̃��[�����ǂ��ɂ��邩�������i���[���̃t�H���_�̏ꏊ)(���S�폜���ꂽ���̂Ɋւ��Ă̓J�E���g���Ȃ�)
  Public Function NumIncrement(ByVal HasUnCounted,ByVal CurrentFolderName)
  
   If HasUnCounted Then
     CumulativeMailNum=CumulativeMailNum+1
   End If
     
   Select Case CurrentFolderName
     Case "ReceivedFolder"
       CurrentSavedMailNum=CurrentSavedMailNum+1
     Case "DeletedFolder"
       CurrentTmpMailNum=CurrentTmpMailNum+1
     Case Else
   End Select
     
   
  End Function
  
  Public Function GetAddress()
   GetAddress=DstEmailAddress
  End Function
  
  Public Function GetName()
   GetName=DstName
  End Function
  
  Public Function GetMailState()
   GetMailState=State
  End Function
  
  Public Function GetFirstDate()
   GetFirstDate=FirstDate
  End Function
  
  Public Function ToStr()
   Dim Former
   Dim Letter
   Former=EscapeStrForCSV(DstEmailAddress)&","&EscapeStrForCSV(DstName)&","&FirstDate&","&CumulativeMailNum
   Letter=","&(CurrentSavedMailNum+CurrentTmpMailNum)&","&CurrentSavedMailNum&","&CurrentTmpMailNum&","&State
   ToStr=Former&Letter
  End Function

End Class


Class DataManager

   Public AddrNumLists()
   
   Public Sub Class_Initialize()
    '�Ǘ��z��i���̒��Ɋe���[���A�h���X�i�����j����̃��[���f�[�^�Z�b�g���i�[�j
    ReDim AddrNumLists(0)
   End Sub
   
   Public Sub Class_Terminate()
    Dim i
    For i=LBound(AddrNumLists) To UBound(AddrNumLists)
      Set AddrNumLists(i)=Nothing
    Next
   End Sub
   
   '�t�@�C���̓��e����f�[�^�Z�b�g(���[���A�h���X�A�����A�ŏ��̓����A���̓����ȍ~�̂��̈��悩��̃��[�����A��舵��)�����o��
   Public Function ParseDataFromFileContent(OneData)
    ReDim Preserve AddrNumLists(UBound(AddrNumLists)+1)
    Dim NewObj
    Set NewObj=New AddrNumSet
    If UBound(OneData) = 7 Then
       Select Case OneData(7)
        Case "�ۑ�","�폜�ς݂ֈړ�","���S�폜"
          NewObj.SetValue DeEscapeStrForCSV(OneData(0)),DeEscapeStrForCSV(oneData(1)),OneData(3),OneData(7)
        Case Else
          NewObj.SetValue DeEscapeStrForCSV(OneData(0)),DeEscapeStrForCSV(oneData(1)),OneData(3),"�폜�ς݂ֈړ�"
       End Select
    ElseIf UBound(OneData) >= 3 Then
       NewObj.SetValue DeEscapeStrForCSV(OneData(0)),DeEscapeStrForCSV(OneData(1)),OneData(3),"�폜�ς݂ֈړ�"
    End If
    On Error Resume Next
     NewObj.AddDate CDate(OneData(2))
     If Err.Number <> 0 Then
       NewObj.AddDate CDate("1970/1/1")
     End If
     Err.Clear
    On Error GoTo 0
    Set AddrNumLists(UBound(AddrNumLists))=NewObj
    
   End Function
    
   '���[���A�h���X�ƈ������L�[�Ƃ���,���̃��[���A�h���X�i�����j�̏�񂪊Ǘ��z��̂ǂ��̃C���f�b�N�X�ɂ��邩������
   Public Function DataIndex(ByVal Address,ByVal Name)
    Dim i
    For i=1 To UBound(AddrNumLists)
      If AddrNumLists(i).GetAddress = Address And AddrNumLists(i).GetName = Name Then
        DataIndex=i
        Exit Function
      End If
    Next
    DataIndex=-1
   End Function
   
   '������͗^����ꂽ���[���A�h���X�ƈ�������̃��[�������݂��邩�ǂ�����Ԃ�
   '�܂�,��̃��\�b�h�̖߂�l��-1�̎��͂��̈��悩��̃��[���͂Ȃ�.����ȊO�̎��͂���Ƃ�������
   Public Function DataExists(ByVal Address,ByVal Name)
     If DataIndex(Address,Name) <> -1 Then
       DataExists=True
       Exit Function
     End If
     DataExists=False
   End Function
   
   '���[�����J�E���g����(�J�E���g��0�A�܂�܂����̈��悩��̃��[�������݂��Ȃ��ꍇ��,�V�����I�u�W�F�N�g�����)
   Public Function Count(ByVal Address,ByVal Name,ByVal MailDate,ByVal HasUnCounted,ByVal CurrentFolderName)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      AddrNumLists(index).NumIncrement HasUnCounted,CurrentFolderName
      AddrNumLists(index).AddDate MailDate
      Exit Function
    End If
    
    If HasUnCounted Then
      ReDim Preserve AddrNumLists(UBound(AddrNumLists)+1)
      Dim NewObj
      Set NewObj=new AddrNumSet
      NewObj.SetValue Address,Name,"1","�폜�ς݂ֈړ�"
      NewObj.AddDate MailDate
      NewObj.NumIncrement False,CurrentFolderName
      Set AddrNumLists(UBound(AddrNumLists))=NewObj
    End If
   End Function
   
   '�����_�ł̍��v���[����
   Public Function GetSumMailNum(ByVal ItemName)
    GetSumMailNum=0
    Dim i
    For i= 1 To UBound(AddrNumLists)
      GetSumMailNum=GetSumMailNum+AddrNumLists(i).GetMailNum(ItemName)
    Next
  End Function
  
 
  '�^����ꂽ���[���A�h���X�i�����j����̃��[���̐���Ԃ�
  Public Function GetNum(ByVal Address,ByVal Name,ByVal ItemName)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      GetNum=AddrNumLists(index).GetMailNum(ItemName)
      Exit Function
    End If
    
    GetNum=0
  End Function
  
  '���̈���̃��[�����ǂ̂悤�Ɉ�����(�폜�ς݂ֈړ��i�f�t�H���g)���A�ۑ���,���S�폜��)
  Public Function GetState(ByVal Address,ByVal Name)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      GetState=AddrNumLists(index).GetMailState()
      Exit Function
    End If
    GetState="�폜�ς݂ֈړ�"
  End Function
  
  
  '���߂ăt�@�C���ɏ������ލ�,�\���������ԏ��ɂ���K�v������
  Public Function DateSort()  
    Dim i
    Dim j
    Dim tmp
    Dim head
    head=1
    
   For i=2 To UBound(AddrNumLists)   
      j=i
      'VBS�ł͒Z���]�����Ă���Ȃ��̂�,��ԍŏ��̃C���f�b�N�X���O�̃C���f�b�N�X�ɃA�N�Z�X���Ă��܂��G���[���o��̂�
      '�����ł�,��ԍŏ��̃C���f�b�N�X��1�O�܂ł̃C���f�b�N�X�ɂ��ă\�[�g��
      Do While (j > head+1) And(AddrNumLists(j-1).GetFirstDate() > AddrNumLists(j).GetFirstDate())
        Set tmp=AddrNumLists(j)
        Set AddrNumLists(j)=AddrNumLists(j-1)
        Set AddrNumLists(j-1)=tmp
        j=j-1
      Loop
      
      '�Ō�̃C���f�b�N�X���������ŕʓr�s��
      If (AddrNumLists(head).GetFirstDate() > AddrNumLists(head+1).GetFirstDate()) Then
        Set tmp=AddrNumLists(head)
        Set AddrNumLists(head)=AddrNumLists(head+1)
        Set AddrNumLists(head+1)=tmp
      End If
      
   Next
        
  End Function
  
  '���[���̏W�v���J�n�������̓��t���擾
  Public Function GetCountStartDate()
    DateSort
    GetCountStartDate=AddrNumLists(1).GetFirstDate
  End Function
  
  
  '���ۂɃt�@�C���ɏ������ނƂ��̕����񂪗v�f�ƂȂ����z��𐶐�����
  Public Function ToFileWriteStr()
   
    DateSort()
    Dim HeaderFormer
    Dim Header
    Dim Today
    Dim FWriter
    Today=Date()
    Dim FirstDate
    FirstDate=GetCountStartDate
    
    HeaderFormer="���[���A�h���X,����,"&FirstDate&"�ȍ~�ōł��������̈��悩�烁�[�����͂������t,"&FirstDate&"����"&Today&"�܂łɓ͂����ݐσ��[����(���S�폜���ꂽ���̂��܂�),"
    Header=HeaderFormer&Today&"���_�ő��݂��Ă��郁�[����(�����S�폜�E�܍폜�ς݃t�H���_),"&Today&"���_�ł̎�M�t�H���_�̃��[����,"&Today&"���_�̍폜�ς݃t�H���_�̃��[����,���[���̎�舵��"
    
    Dim Content()
    ReDim Preserve Content(0)
    Content(0)=Header
    Dim i
    For i=1 To UBound(AddrNumLists)
     ReDim Preserve Content(UBound(Content)+1)
     Content(UBound(Content))=AddrNumLists(i).ToStr
    Next
    
    ReDim Preserve Content(UBound(Content)+1)
    Dim Footer
    Dim CumulativeMailSum
    CumulativeMailSum=GetSumMailNum("Cumulative")
    Dim ExistsMailSum
    ExistsMailSum=GetSumMailNum("Exists")
    Dim ReceivedMailSum
    ReceivedMailSum=GetSumMailNum("ReceivedFolder")
    Dim DeletedMailSum
    DeletedMailSum=GetSumMailNum("DeletedFolder")
    Footer="���v,,,"&CumulativeMailSum&","&ExistsMailSum&","&ReceivedMailSum&","&DeletedMailSum
    Content(UBound(Content))=Footer
    
    ToFileWriteStr=Content
    
  End Function

End Class
    
    

'�t�@�C���̓ǂݏ���������  
Class FileOperator

  Public FIOOperator
  
  Public Sub  Class_Initialize()
   Set FIOOperator=Nothing
  End Sub
  
  Public Sub Class_Terminate()
   Set FIOOperator=Nothing
  End Sub
  
  Public Function SetFSObj()
   Set FIOOperator=Wscript.CreateObject("ADODB.Stream")
   FIOOperator.Type=2
   FIOOperator.Charset="UTF-8"
   FIOOperator.LineSeparator=10
  End Function
  
  Public Function FRead(FilePath)
   SetFSObj
   
   FIOOperator.Open
   
   Dim FileOpen
   FileOpen=True
   
   '�t�@�C������f�[�^���擾����ۂ�,���̃t�@�C�����̂��J����Ă������̓G���[���o��̂�,File��������܂ŉi�v���[�v����
   Do While FileOpen
     On Error Resume Next
      FIOOperator.LoadFromFile FilePath
      
      '�G���[���Ȃ������i�t�@�C���������Ă�����),�����ɂ��ǂ蒅���̂�,FileOpen�t���O��������,���[�v�𔲂���
      If Err.Number = 0 Then
        FileOpen=False
      End If
      Err.Clear
     On Error GoTo 0 
     
   Loop
   
   Dim Result()
   Dim LineNum
   LineNum=0
   Dim OneLine
   
   Do While FIOOperator.EOS = False
     ReDim Preserve Result(LineNum)
     OneLine=FIOOperator.ReadText(-2)
     Result(LineNum)=OneLine
     LineNum=LineNum+1
   Loop
   
   FRead=Result
     
   FIOOperator.Close
   Set FIOOperator=Nothing
  
  End Function 
  
  '�t�@�C���̏�������(Mode��"w"�Ȃ�㏑��,"a"�Ȃ�ǋL,WriteType��"Array"�Ȃ�z��Ƃ���,�e�v�f��1�s�������Ă䂭,"Str"�Ȃ當����Ƃ���1�s��������)
  Public Function FWrite(ByVal FilePath,Contents,ByVal Mode,ByVal WriteType)
    Dim FileOpen
    
    SetFSObj
    FIOOperator.Open
    Select Case Mode
      Case "a","A"
       FileOpen=True
       Do While FileOpen
        On Error Resume Next
          FIOOperator.LoadFromFile FilePath
          If Err.Number = 0 Then
            FileOpen=False
          End If
          Err.Clear
        On Error GoTo 0
       Loop
       FIOOperator.Position=FIOOperator.Size
    End Select
    
    Select Case WriteType
     Case "Str"
       FIOOperator.WriteText Contents,1
     Case Else
       Dim i
       For i= LBound(Contents) To UBound(Contents)
         FIOOperator.WriteText Contents(i),1
       Next
    End Select
  
    FileOpen=True
    
    '�G���[�`�F�b�N(�����������Ƃ��Ă���t�@�C�����J����Ă����ꍇ,�G���[���o�Ă��܂��̂�,�G���[���Ȃ��Ȃ�i�t�@�C����������܂Łj�҂�)
    Do While FileOpen
      On Error Resume Next
        FIOOperator.SaveToFile FilePath,2
        If Err.Number = 0 Then
          FileOpen=False
        End If
        Err.Clear
      On Error GoTo 0 
    Loop
    
    FIOOperator.Close
    Set FIOOperator=Nothing
    
  End Function

End Class


Class FileManager

  Public FIO
  Public WshObj
  Public FSObj
  Public OriginalFileName
  Public BackupFileName
  Public BackupFolder
  Public TimeLogFile
  Public NumLogFile
  Public BackupLogFolder
  
  
  Public Sub Class_Initialize()
   Set WshObj=Wscript.CreateObject("Wscript.Shell")
   Set FSObj=Wscript.CreateObject("Scripting.FileSystemObject")
   Set FIO=New FileOperator
   Dim DesktopFolder
   DesktopFolder=WshObj.SpecialFolders(4)
   OriginalFileName=DesktopFolder&"\outlook_mail_dest_list.csv"
   BackupFolder=WshObj.SpecialFolders(5)
   BackupFileName=BackupFolder&"\outlook_mail_dest_list.csv"
   BackupLogFolder=BackupFolder&"\backup"
   '�o�b�N�A�b�v�t�H���_���Ȃ��Ȃ���
   If Not FSObj.FolderExists(BackupLogFolder) Then
      FSObj.CreateFolder(BackupLogFolder)
   End If
   TimeLogFile=BackupFolder&"\datelog.log"
   NumLogFile=BackupFolder&"\mail_num.log"
  End Sub
  
  Public Sub Class_Terminate()
   Set WshObj=Nothing
   Set FSObj=Nothing
   Set FIO=Nothing
  End Sub
  
  Public Function GetDataManageObj()
   Set GetDataManageObj=new DataManager
   Dim FileContents
   If FSObj.FileExists(OriginalFileName) Then
     FileContents=FIO.FRead(OriginalFileName)
   ElseIf FSobj.FileExists(BackupFileName) Then
     Dim FObj
     Set FObj=FSObj.GetFile(BackupFileName)
     FObj.attributes=0
     FileContents=FIO.FRead(BackupFileName)
     FObj.attributes=1
     Set FObj=Nothing
   Else
     Exit Function
   End If
   
   Dim i
   Dim AllDataWithoutBr
   Dim OneData
   '�w�b�_�͂���Ȃ��̂�1�Ԗڂ���擾����B�����čŌ�̍s�͍��v�Ȃ̂ł��������Ȃ�
   For i=1 To UBound(FileContents)-1
     AllDataWithoutBr=Replace(FileContents(i),VbCr,"")
     OneData=SplitExceptEscapeChar(AllDataWithoutBr,",","""")
     GetDataManageObj.ParseDataFromFileContent OneData
   Next
   
  End Function
  
  '���[���𒲂ׂ�ɂ�����,�ۑ����[���̏d���J�E���g������邽��,�O��A���[���̐����J�E���g�����̂͂��Ȃ̂��𓾂�
  '�ۑ����[���ɓ����Ă���,���̓��t���O�̃��[���Ɋւ��Ă͂��łɃJ�E���g���Ă���̂ŃJ�E���g���Ȃ�
  Public Function GetLastMailDate()
   '���O�t�@�C������Ō�Ƀ��[�����`�F�b�N�������t���𓾂�
   If FSObj.FileExists(TimeLogFile) Then
    FSObj.GetFile(TimeLogFile).attributes=0
    Dim Contents
    Contents=FIO.FRead(TimeLogFile)
    FSObj.GetFile(TimeLogFile).attributes=1
    GetLastMailDate=CDate(Replace(Contents(0),VbCr,""))
   ElseIf FSObj.FileExists(OriginalFileName) Then
    'Log�t�@�C�����Ȃ������ꍇ���̓��t�ő�p����
    GetLastMailDate=FSObj.GetFile(OriginalFileName).DateLastModified
   ElseIf FSObj.FileExists(BackupFileName) Then
    GetLastMailDate=FSObj.GetFile(BackupFileName).DateLastModified
   Else
    GetLastMailDate=CDate("1970/1/1")
   End If
   
  End Function
  
  '���ۂɌ��ʂ̏�������
  Public Function WriteResultDataManageObj(DataManageObj)
   FIO.FWrite OriginalFileName,DataManageObj.ToFileWriteStr(),"w","Array"
   Dim FObj
   '�t�@�C�����Ȃ��Ƃ��i�͂��߂ăo�b�N�A�b�v�t�@�C�����쐬�����,�t�@�C�����̂����݂��Ȃ��̂�,������ς���ɂ��ς����Ȃ�����
   '���̎��̓G���[������Ԃ�
   On Error Resume Next
     FSObj.GetFile(BackupFileName).attributes=0
     Err.Clear
   On Error GoTo 0
   FSObj.CopyFile OriginalFileName,BackUpFileName
   Set FObj=FSObj.GetFile(BackupFileName)
   FObj.attributes=1
  End Function
  
  '���񃁁[���𒲂ׂăJ�E���g����ɂ�����,���񉽎������̃��[���܂ł��J�E���g�ς݂Ȃ̂����L�^���Ă���
  '��L�̂悤�Ɏ���A�ǂ̃��[������J�E���g����΂悢�̂����������߁i�d���J�E���g������邽�߁j
  Public Function WriteRenewLastMailDate(ByVal CountStartDate,ByVal LastMailDate,ByVal CumulativeMailNum,ByVal ExistsMailNum,ByVal SaveMailNum,ByVal DeletedMailNum)
   On Error Resume Next
    FSObj.GetFile(TimeLogFile).attributes=0
    Err.Clear
   On Error GoTo 0
   FIO.FWrite TimeLogFile,""&LastMailDate,"w","Str"
   FSObj.GetFile(TimeLogFile).attributes=1
   
   Dim NowDate
   NowDate=Now
   
   If FSObj.FileExists(NumLogFile) Then
     Dim Content
     Content=""&NowDate&","&LastMailDate&","&CumulativeMailNum&","&ExistsMailNum&","&SaveMailNum&","&DeletedMailNum
     FSObj.GetFile(NumLogFile).attributes=0
     FIO.FWrite NumLogFile,Content,"a","Str"
     FSObj.GetFile(NumLogFile).attributes=1
   Else
    Dim Header
    Header="�v���O�������s����,���̎��_�ł̍ŐV�̃��[������,"&CountStartDate&"���炻�̎��_�܂ł̗ݐσ��[����(�܊��S�폜),���̎��_�ő��݂��Ă��郁�[����(�����S�폜�E�܍폜�ς݃t�H���_),��M�ς݃t�H���_�̃��[����,�폜�ς݃t�H���_�̃��[����"
    Dim Body
    Body=""&NowDate&","&LastMailDate&","&CumulativeMailNum&","&ExistsMailNum&","&SaveMailNum&","&DeletedMailNum
    Dim Contents(1)
    Contents(0)=Header
    Contents(1)=Body
    FIO.FWrite NumLogFile,Contents,"w","Array"
    FSObj.GetFile(NumLogFile).attributes=1
   End If
   
   Dim yymmddhhmmssStr
   yymmddhhmmssStr=ToyymmddhhmmssStr(NowDate)
   Dim BackupSaveLogFile
   BackupSaveLogFile=BackupLogFolder&"\outlook_mail_dest_list_"&yymmddhhmmssStr&"_backup.csv"
   FSObj.CopyFile OriginalFileName,BackupSaveLogFile
   FSObj.GetFile(BackupSaveLogFile).attributes=1
  End Function
  
  
  Public Function ToyymmddhhmmssStr(NowDate)
   ToyymmddhhmmssStr=Year(NowDate)&PadZero(Month(NowDate),2)&PadZero(Day(NowDate),2)&PadZero(Hour(NowDate),2)&PadZero(Minute(NowDate),2)&PadZero(Second(NowDate),2)
  End Function
  
  Public Function PadZero(ByVal Before,ByVal Num)
    Dim BeforeInt
    BeforeInt=CLng(Before)
    Dim Result
    Result=""
    Dim Digit
    Digit=1
    Dim DigitNum
    Dim i
    i=0
    Do While i < Num Or Digit < BeforeInt
      DigitNum=((BeforeInt\Digit) Mod 10)
      Result=""&DigitNum&Result
      Digit=Digit*10
      i=i+1
    Loop
    PadZero=Result
  End Function
     
       
End Class

'���[���A�C�e�������S�폜����ہA��M�ς݃t�H���_�����C�ɍ폜���邱�Ƃ��ł��Ȃ�
'�䂦�Ɋ��S�폜�̍ۂ�,�u��M�ς݂��炢������폜�ς݂Ɉړ����Ă���A��������폜�v,�܂�Q��̈ړ��葱�������Ȃ���΂Ȃ�Ȃ���
'���S�폜�����,���̊��S�폜�������Ǝv���Ă��郁�[�����폜�ς݃t�H���_�̂ǂ��ɂ���̂��𑗐M�ҏ��⑗�M�����Ȃǂ��q���g�ɒT���K�v������
'�ɂ�������炸��M�ς݂���폜�ς݂Ɉړ������Ƃ�(1��ڂ̈ړ��̎��j�ɁA���S�폜���悤�Ǝv���Ă��郁�[���̑��M�ҏ��⑗�M�����Ȃǂ��Q�Ƃł��Ȃ��Ȃ��Ă��܂�.
'�䂦�ɂ����ł�,��M�ς݂���̈ړ��̑O��,���M�ҏ�񓙂��c���Ă����ꎞ�ϐ��I�N���X�������Łu�킴�킴�v����Ă���

Class TmpMailItem
  Public m_SenderEmailAddress
  Public m_SenderName
  Public m_ReceivedTime
  Public m_Subject
  Public m_Body
  
  Public Sub Class_Initialize()
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  
  Public Function SetMailInfo(Email,Name,Time,Subj,Bd)
   m_SenderEmailAddress=Email
   m_SenderName=Name
   m_ReceivedTime=Time
   m_Subject=Subj
   m_Body=Bd
  End Function
  
  '�o�O�̉����ƂȂ邽��,�t�B�[���h�ɒ��ڃA�N�Z�X����,�v���p�e�B��ʂ��悤�ɂ���
  Public Property Get SenderEmailAddress()
    SenderEmailAddress=m_SenderEmailAddress
  End Property
  
  Public Property Get SenderName()
    SenderName=m_SenderName
  End Property
  
  Public Property Get ReceivedTime()
    ReceivedTime=m_ReceivedTime
  End Property
  
  Public Property Get Subject()
    Subject=m_Subject
  End Property
  
  Public Property Get Body()
    Body=m_Body
  End Property
  
 
End Class



'���[���A�C�e�����������̂Ȃ̂��𔻒肷��(�^��(�ꎞ�ۑ���������:�ڍׂ͌�q)�ł��{���̃��[���A�C�e���I�u�W�F�N�g�ł������Ƃ������悤�ɔ��肷��)
Function MailItemEquals(One,Other)
    If One.SenderEmailAddress <> Other.SenderEmailAddress Then
      MailItemEquals=False
      Exit Function
    End If
    
    If One.SenderName <> Other.SenderName Then
      MailItemEquals=False
      Exit Function
    End If
    
    If One.Subject <> Other.Subject Then
      MailItemEquals=False
      Exit Function
    End If
    
    If One.Body = Other.Body Then
      MailItemEquals=True
      Exit Function
    End If
    
    '����Đ��b����ē���̓��e�̃��[���������Ă��܂��Ƃ������Ƃ�����\�����l����
    '�����Ɋւ��Ă͑S���������蓯���łȂ��Ă�5�b�����̌덷�Ȃ瓯�����m��������
    If 5 < Abs(DateDiff("s",One.ReceivedTime,Other.ReceivedTime)) Then
      MailItemEquals=False
      Exit Function
    End If
     
    
    MailItemEquals=True
    
End Function
 

'�{����ReceivedFolder��DeletedFolder��,���ԏ��ɕ���łȂ��ďd�����[���̍폜���������,��������1����ׂȂ���΂����Ȃ�
'�Ȃ̂ł����ł͍폜�ς݃t�H���_�Ǝ�M�t�H���_�̃A�C�e�����R�s�[��������(��L�ŏq�ׂ��^�����[�����)�������Ă���)
Class TimeOrderedFolder

   Public TimeOrderedMailItem()
   Public Num
   
   Public Sub Class_Initialize()
     ReDim Preserve TimeOrderedMailItem(0)
     Num=0
   End Sub
   
  Public Sub Class_Terminate()
     ReDim Preserve TimeOrderedMailItem(0)
     Set TimeOrderedMailItem(0)=Nothing
     Num=0
  End Sub
  
  '����͂����T�����邾���̃��\�b�h
  '�������,���̃N���X�̊Ǘ��z��ɂ����邻�̗v�f�̃C���f�b�N�X��Ԃ�
  '������Ȃ������ꍇ�͐V�����}���\��̕��̃C���f�b�N�X������
  Public Function Search(OneMailInfo)
    Dim Start
    Dim Goal
    
    Start=1
    Goal=UBound(TimeOrderedMailItem)
    Do While Start <= Goal
      Dim CurrentIndex
      CurrentIndex=((Start+Goal)\2)
      Dim CurrentItem
      Set CurrentItem=TimeOrderedMailItem(CurrentIndex)
      If MailItemEquals(OneMailInfo,CurrentItem) Then
         Search=CurrentIndex
         Exit Function
      ElseIf OneMailInfo.ReceivedTime < CurrentItem.ReceivedTime Then
         Goal=CurrentIndex-1
      Else
         Start=CurrentIndex+1
      End If
      Set CurrentItem=Nothing
    Loop
      
    Search=(-1)*Start
  End Function
  
  Public Function Find(OneMailInfo)
     '�A�C�e����1���Ȃ��Ƃ�
     If UBound(TimeOrderedMailItem) = 0 Then
       Find=-1
       Exit Function
     '�P�ԍŏ��̃A�C�e����莞�Ԃ������ꍇ
     ElseIf OneMailInfo.ReceivedTime < TimeOrderedMailItem(1).ReceivedTime Then
       Find=-1
       Exit Function
     '1�ԍŌ�̃A�C�e����莞�Ԃ��x���ꍇ������3�̏����̂����ǂꂩ1�ɓ��Ă͂܂�Ƃ���,�T�����Ȃ��Ă悢
     ElseIf TimeOrderedMailItem(UBound(TimeOrderedMailItem)).ReceivedTime < OneMailInfo.ReceivedTime Then
       Find=-1
       Exit Function
     End If
     
     
     Dim Result
     Result=Search(OneMailInfo)
     If Result < 0 Then
       Find=-1
       Exit Function
     End If
     Find=Result
  End Function
  
  '������͑}��(������Ȃ������ꍇ�ɂ̂ݍs��,Search���\�b�h���Ԃ������̃C���f�b�N�X�𐳂ɖ߂��āA���̎��ӂɓ����
  Public Function AddData(OneMailInfo)
    
    '�v�f������Ȃ��Ƃ��͒��ׂ���1�Ԗڂɑ}��
    If UBound(TimeOrderedMailItem) = 0 Then
       Insert 1,OneMailInfo
       Exit Function
    '1�Ԗڂ̗v�f��葁�����Ԃ̂��̂�1�Ԗڂɑ}��
    ElseIf OneMailInfo.ReceivedTime < TimeOrderedMailItem(1).ReceivedTime Then
       Insert 1,OneMailInfo
       Exit Function
    '1�ԍŌ�̗v�f��莞�Ԃ��x�����͈̂�ԍŌ�ɑ}��
    ElseIf TimeOrderedMailItem(UBound(TimeOrderedMailItem)).ReceivedTime < OneMailInfo.ReceivedTime Then
       Insert UBound(TimeOrderedMailItem)+1,OneMailInfo
       Exit Function
    End If
    
    '��L�̂悤�ȏꍇ�͒��ׂȂ��Ă��}������ꏊ�͎����ł�������,����ȊO�̏ꍇ�͑}������ꏊ�𒲂ׂȂ��Ă͂Ȃ�Ȃ�
    Dim InsertIndex
    InsertIndex=Search(OneMailInfo)
    '���łɑ��݂����ꍇ�͑}�����Ă͂����Ȃ�
    If InsertIndex > 0 Then
       Exit Function
    End If
    
    InsertIndex=(-1)*InsertIndex
    Insert InsertIndex,OneMailInfo
    
  End Function
  
  Public Function Insert(ByVal Index,NewItem) 
  
    Num=Num+1
    ReDim Preserve TimeOrderedMailItem (UBound(TimeOrderedMailItem)+1)
    Dim i
    
    For i=UBound(TimeOrderedMailItem) To Index+1 Step -1
       Set TimeOrderedMailItem(i)=TimeOrderedMailItem(i-1)
    Next
       
    Set TimeOrderedMailItem(Index)=NewItem
    
  End Function
  
  Public Function CheckPrint()
    Dim i
    For i=1 To UBound(TimeOrderedMailItem)
     Wscript.Echo i&"�Ԗ� ���[���A�h���X:"&TimeOrderedMailItem(i).SenderEmailAddress&" ����:"&TimeOrderedMailItem(i).SenderName&" ����:"&TimeOrderedMailItem(i).ReceivedTime&" ����"&TimeOrderedMailItem(i).Subject
    Next
  End Function
  
  Public Function  Item(index)
    Items=TimeOrderedMailItem(index)
  End Function
  
  Public Property Get Count()
    Count=UBound(TimeOrderedMailItem)
  End Property
    
End Class

'�����ł͍폜�ς݃t�H���_�̃��[���̏d�����폜����N���X�����
'�ŏ��Ɂu�폜�ς݃t�H���_�v���̏d�����[�����폜���Ă��܂��B
'���̌�,�u��M�t�H���_�v���������Ƃ��Ɂu��M�t�H���_�v�Ɓu�폜�ς݃t�H���_�v�ɂ܂������Ă������,���邢��,�u��M�t�H���_�v���m�̏d�����폜����
'���ۂ͎�M�t�H���_�ƍ폜�ς݃A�C�e���͕ʁX�ɂȂ��Ă��邪�A����TimeOrderedReceivedAndDeletedFolder�ɂ͎�M�t�H���_�ƍ폜�ς݃t�H���_�̗������킹�����[���A�C�e��(��ŏq�ׂ��^��)�������ł͊i�[��,�d����r������
'����Łu��M�t�H���_�v�݂̂̏d���폜,�u�폜�ς݃t�H���_�v�݂̂̏d���폜�ɂƂǂ܂炸,���t�H���_�ɂ܂������ďd������A�C�e�����ꊇ�ō폜�ł���
Class TimeOrderedRemoveDuplicateInFolder
  Public TimeOrderedReceivedAndDeletedFolder
  
  Public Sub Class_Initialize()
    Set TimeOrderedReceivedAndDeletedFolder=New TimeOrderedFolder
  End Sub
  
  Public Sub Class_Terminate()
    Set TimeOrderedReceivedAndDeletedFolder=Nothing
  End Sub

  '�����ł͍ŏ��Ɂu�폜�ς݃t�H���_�v�ɓ����Ă�����݂̂̂��폜
  Public Function  RemoveDuplicateInDeletedFolder(DeletedFolder)
   Dim MailItemIndex
   MailItemIndex=1
  
   '�e�X�g�p���S�폜�ς݃t�H���_�\(�e�X�g�̎��͂����Ȃ芮�S�폜���Ȃ�)
   Dim testCompDeleteFolder
   Set testCompDeleteFolder=DeletedFolder.Folders("test_comp_delete")
  
   '�e���[���A�C�e���𒲂ׂ�
   Do While MailItemIndex <= DeletedFolder.Items.Count
    Dim CurrentCheckingMailItem
    Set CurrentCheckingMailItem=DeletedFolder.Items.Item(MailItemIndex)
    Dim CurrentCheckingTmpMailItem
    Set CurrentCheckingTmpMailItem=New TmpMailItem
    Dim Addr
    Dim Name
    Dim Time
    Dim Subj
    Dim Body
    Addr=CurrentCheckingMailItem.SenderEmailAddress
    Name=CurrentCheckingMailItem.SenderName
    Time=CurrentCheckingMailItem.ReceivedTime
    Subj=CurrentCheckingMailItem.Subject
    Body=CurrentCheckingMailItem.Body
    CurrentCheckingTmpMailItem.SetMailInfo Addr,Name,Time,Subj,Body
    If TimeOrderedReceivedAndDeletedFolder.Find(CurrentCheckingTmpMailItem) = -1 Then
        TimeOrderedReceivedAndDeletedFolder.AddData CurrentCheckingTmpMailItem
        '�C���f�b�N�X�̉��Z�͌�����Ȃ������Ƃ��̂�
        MailItemIndex=MailItemIndex+1 
    Else
      '�V�݂������̂Ȃ̂łƂ肠�����e�X�g���S�폜�t�H���_�[�ֈړ�
      Dim TestFolderMailNum
      TestFolderMailNum=testCompDeleteFolder.Items.Count
      CurrentCheckingMailItem.Move testCompDeleteFolder
      Do While TestFolderMailNum = testCompDeleteFolder.Items.Count
        Wscript.Sleep 500
      Loop
        
      '�ȍ~�̓e�X�g�I����(���̓R�����g�A�E�g)
      'Dim CurrentDeletedFolderMailNum
      'Dim CurrentFolderMailNum
      'CurrentFolderMailNum=Folder.Items.Count
      'Dim CurrenDeletedFolderMailNum
      'CurrentDeletedFolderMailNum=DeletedFolder.Items.Count
      'Do
        'On Error Resume Next
        'Err.Clear
        '�u��M�ς݁v�ɂ��郁�[���A�C�e���ɑ΂���Delete���\�b�h��p�����ꍇ,���̂܂܂��Ɓu�폜�ς݁v�Ɉړ������ŏI����Ă��܂�
        '������,2�x�ڂ̊֐��Ăяo�����́u�폜�ς݃t�H���_�v�𒲂ׂ��,���R�Ȃ��炱�̃��[�����͏d�����Ă���Ɣ��肳���
        '�i1�x�ڂ̌Ăяo���̍ۂɏd���Ɣ��肳����2�x�ڂ̍ۂ̓f�[�^�͌���Ȃ��̂ŕK���d���Ɣ��肳���j�̂ŁA2�x�ڂ̌Ăяo���̍ۂɓ����悤��Delete����Ί��S�폜�����̂ŐS�z�Ȃ�
        'CurrentCheckingMailItem.Delete
        'On Error GoTo 0
      'Loop While Err.Number <> 0
       
      '�u�폜���\�b�h���Ăяo���O�̃��[���A�C�e�����ƌ��݂̃��[���A�C�e�����v(�폜�ς݃t�H���_�̃A�C�e���������l)�̕ω�������ׂ�,�����ς������ړ�(�폜)�����������Ƃ������ƂɂȂ�
      '���������폜�ς݃t�H���_�������Ƃ��͑S������������2��ɂȂ�̂Ŗ��ʂȏ����ɂȂ邩������Ȃ���,�����֐���2�x��`�������M�ς݂����ʂ̏����������肩�͂͂邩�Ɍ������ǂ�)
      'Do While CurrentFolderMailNum = Folder.Items.Count And  CurrentDeletedFolderMailNum = DeletedFolder.Items.Count 
        'Wscript.Sleep 500
      'Loop
      '�R�����g�A�E�g�I��
    End If
    Set CurrentCheckingMailItem=Nothing
    
  Loop
 End Function
 
 '�u��M�t�H���_�v�̃��[���A�C�e���𒲂ׂč폜�����,���̃A�C�e�������łɁu(�^��)�폜�ς݃t�H���_�v�ɑ��݂��Ȃ����ǂ������ׂ�
 '�����Ƃ��ė^����̂́u��M�t�H���_�v�̒������̃��[���A�C�e��(���[���A�C�e�����̂��͈̂ڂ��Ȃ��̂ŋ^���̂��́j
 '�u���̋^���폜�ς݃t�H���_�v�ɁA�����ׂĂ����M�t�H���_�̈ꃁ�[���A�C�e�������݂����ꍇ�͏d������Ƃ������ƂȂ̂�,True��Ԃ�
 '�u�^���폜�ς݃t�H���_�v�ɍ����ׂĂ��郁�[���A�C�e�������݂��Ȃ����,���̋^���폜�ς݃t�H���_�Ƀ��[���A�C�e����ǉ���,False��Ԃ�.
 '���̂��Ƃɂ����,���Ƃ��Ɓu��M�ς݃t�H���_�v�ɂ������A�C�e�������̏d�������ׂ���
 Public Function CheckDuplicate(OneTmpMailItem)
    CheckDuplicate=True
    If TimeOrderedReceivedAndDeletedFolder.Find(OneTmpMailItem) = -1 Then
       TimeOrderedReceivedAndDeletedFolder.AddData OneTmpMailItem
       CheckDuplicate=False
    End If
 End Function
 
 
 Public Function CheckPrint()
   TimeOrderedReceivedAndDeletedFolder.CheckPrint
 End Function
 
 Public Property Get Items()
   Items=TimeOrderedReceivedAndDeletedFolder
 End Property
 

End Class
   

'�ꎞ�ۑ��������̃��[����񂪃t�H���_�̂ǂ̃C���f�b�N�X�ɑ��݂��邩�𒲂ׂ郁�\�b�h(���`�T��)
'���݂���ꍇ�͑��݂���C���f�b�N�X,���݂��Ȃ��ꍇ��-1��ԋp����
'�{���̃��[���A�C�e���I�u�W�F�N�g�ł�,�^���I�ɍ쐬�����ꎞ���[�����ۑ��I�u�W�F�N�g�ł��ǂ���ł��悢
Function MailIndexInFolder(OneMailItem,ItemsFolder)
 Dim CompMailItem
 Dim i

 '�}�����ꂽ�΂���̃f�[�^�͌��̂ق��ɗ��₷��(?)�̂Ō�납�璲�ׂ��ق�������
 For i=ItemsFolder.Items.Count To 1 Step -1
  Set CompMailItem=ItemsFolder.Items.Item(i)
  If MailItemEquals(OneMailItem,CompMailItem) Then
    MailIndexInFolder=i
    Exit Function
  End If
 Next
 MailIndexInFolder=-1
End Function

    
Function Main()

  Dim outlook
  Dim namespace
  Dim receiveFolder
  Dim mailItems
  Dim Wsh

  Dim addressInfo()

  'outlook�ɃA�N�Z�X
  Set outlook=Wscript.CreateObject("Outlook.Application")

  'outlook�̃��[���ɃA�N�Z�X���邽�߂̃C���^�[�t�F�[�X
  Set namespace=outlook.GetNameSpace("MAPI")
  

  'outlook�̎�M�t�H���_
  Set receiveFolder=namespace.GetDefaultFolder(6)
  
  '�폜�ς݃t�H���_
  Dim deletedFolder
  Set deletedFolder=namespace.GetDefaultFolder(3)
  
  
  
  'outlook���ŏ������ċN��(�����outlook���N�����Ă��邱�Ƃ��킩��Ȃ��悤�ɂ���)
  Set Wsh=Wscript.CreateObject("Wscript.shell")
  Wsh.Run "outlook.exe",7,False

  
  '�t�@�C���̓ǂݏ����Ɋւ���Ǘ�������N���X
  Dim FManager
  Set FManager=new FileManager
  
  '���[���̐����舵�����Ǘ�����N���X
  Dim DManager
  Set DManager=FManager.GetDataManageObj()
  
 
  
  '���[���̈�����(���[���̈���ɂ���ă��[���̏����̎d����ς���)
  '�폜�ς݂ֈړ��i�f�t�H���g)��,���S�폜��,�ۑ���
  Dim MailOperation
 
  '���[��
  Dim OneMailItem
   
  '�������Ƃ��Ă����������[�����o���炱�̕ϐ���p���ĂƂ��Ă����������[���̐���ۑ����Ă���
  '�܂胁�[���𑀍삷��ۂ̃C���f�b�N�X
  Dim SaveMailNum
  SaveMailNum=0
  
  Dim CurrentMailNum
  
  Dim TestDeleteMailNum
  
  
  '���S�폜���郁�[�����ꎞ�ۑ��������
  Dim TmpOneMailItem
 
 '�ǂ̎����܂ł̃��[�����J�E���g�����̂��i�O��J�E���g���s���_�ł̍ŐV�̃��[���̎���)�𓾂�(�ۑ��t�H���_����d���J�E���g�����Ȃ��悤��)
  Dim LastCountMailDate
  LastCountMailDate=FManager.GetLastMailDate()
  
  '�{�J�E���g�̒��ň�ԐV��������(���񕪂�LastCountMailDate)
  '���̕ϐ��͍폜�ς݃t�H���_��������ɏ��������Ȃ�
  Dim LastMailTime
  LastMailTime=LastCountMailDate
  
  '���[�����
  '�ォ�烁�[���A�h���X�A�����A��M����,�����A�{��
  Dim Address
  Dim Name
  Dim Time
  Dim Subject
  Dim Body
  
  '�폜�ς݃t�H���_�̏d���폜
  '�e�t�H���_�̏d���폜�͐��ɂ���邪,���`�T���I�ɍs���ꍇ�v�Z�ʂ�n^2�ƂȂ��Ă��܂��̂Ŕ��Ɍ���������
  '����ă��[���̎������Â��ق����珇�Ԃɕ��Ԃ悤�ȏ����t���̍폜�ς݃t�H���_�[(�{���̍폜�ς݃t�H���_�[�͕K���������[���̎������ɕ���ł���킯�ł͂Ȃ�����)���^���I�ɒ�`����
  '�񕪒T���ŏd���𔭌����폜����
  
  '�����t����M+�폜�ς݃t�H���_(���t�H���_�̃��[���A�C�e��(�^��)��������,���ԏ��ɕ��ד񕪒T���ō����ŏd���𔭌����폜����)
  Dim TimeOrderedFolder
  Set TimeOrderedFolder=New TimeOrderedRemoveDuplicateInFolder
  
  TimeOrderedFolder.RemoveDuplicateInDeletedFolder deletedFolder
  
  
  '�܂��͍폜�ς݃A�C�e���ɂ��Ă݂Ă䂭
  Do While SaveMailNum < deletedFolder.Items.Count
     CurrentMailNum=deletedFolder.Items.Count
     Set OneMailItem=deletedFolder.Items.Item(SaveMailNum+1)
     Address=OneMailItem.SenderEmailAddress
     Name=OneMailItem.SenderName
     Time=OneMailItem.ReceivedTime
     Subject=OneMailItem.Subject
     Body=OneMailItem.Body
     MailOperation=DManager.GetState(Address,Name)
     
     If LastMailTime < Time Then
       LastMailTime=Time
     End If
          
     Select Case MailOperation
       Case "�ۑ�"
        OneMailItem.Move receiveFolder
        Do While  CurrentMailNum = deletedFolder.Items.Count 
            Wscript.Sleep 500
        Loop
       Case "���S�폜"
       
        DManager.Count Address,Name,Time,(LastCountMailDate < Time),"CompDeletedFolder"
        
        OneMailItem.Delete
        
        Do While  CurrentMailNum = deletedFolder.Items.Count 
            Wscript.Sleep 500
        Loop
        
       Case Else
        DManager.Count Address,Name,Time,(LastCountMailDate < OneMailItem.ReceivedTime),"DeletedFolder"
        SaveMailNum=SaveMailNum+1
       End Select
  Loop
   
    
  
  
  
  
  '�폜�ς݃A�C�e���̕��������̂Ŏ��͎�M�ς݂̕�������
  Set OneMailItem=Nothing
  Set TmpOneMailItem=Nothing
  
  '�ۑ��i�t�H���_�L�[�v�j�̃��[������0�ɖ߂�
  SaveMailNum=0
  
  '���[���̈����������ɖ߂�
  MailOperation=""
  
  '���[�����̏�����
  Address=""
  Name=""
  Time=""
  Subject=""
  Body=""
 
  '���͎�M�A�C�e���̒������Ă䂭��,��M�A�C�e���̓ǂݍ��݂Ƃ��̏����͔񓯊��ł��邱�Ƃ��班���^�C�����O��݂���K�v������
  '�^�C�����O�p�̕ϐ�
  Dim NormSec
  NormSec=120
  
  Dim AllMailNum
  AllMailNum=receiveFolder.Items.Count+deletedFolder.Items.Count
  
  '���[���̌����������ꍇ,�������̂��̂Ɏ��Ԃ������邽��,�^�C�����O�̕b���͌��炷
  '��������炳�Ȃ��ƃ��[�U�[��҂����鎞�Ԃ�������
  Dim NormSecBias
  NormSecBias=(AllMailNum\100)+1
  NormSec=NormSec\(Sqr(NormSecBias))
  
  Dim CurrentNormSec
  CurrentNormSec=NormSec
 
 
  Dim CountTimes
  CountTimes=0
 
  Dim CountWaitSec
  CountWaitSec=0
 
  Dim EnterFlag
  EnterFlag=False
  
  Dim EnterTime
  EnterTime=0
 
  
  
  '�J�E���^�ϐ�
  Dim i
 
 
  CurrentMailNum=receiveFolder.Items.Count

  Dim HasError
  
  Dim CurrentDeletedFolderMailNum
  
  Do While CountWaitSec < CurrentNormSec 
    Do While SaveMailNum < receiveFolder.Items.Count
      '���[��������(�폜�ς݃t�H���_�j�Ɉړ��������,��M�g���[�̃��[��������̂�,������,�����C���f�b�N�X���A�N�Z�X����
      
      '�^�C�����O�p(��x�ł��폜���������,�^�C�����O�J�E���g�ϐ���0�ɂ���
      EnterFlag=True
      
      HasError=0
      
      On Error Resume Next
      Set OneMailItem=receiveFolder.Items.Item(SaveMailNum+1)
      HasError=Err.Number
      If HasError <> 0 Then
        Wscript.Echo HasError
        Wscript.Echo "�G���["
      End If
      Err.Clear
      On Error GoTo 0
      
      If HasError = 0 Then
      
        Name=OneMailItem.SenderName
        Address=OneMailItem.SenderEmailAddress
        Time=OneMailItem.ReceivedTime
        Subject=OneMailItem.Subject
        Body=OneMailItem.Body
        
        '���[���̏d������
        Set TmpOneMailItem=New TmpMailItem
        TmpOneMailItem.SetMailInfo Address,Name,Time,Subject,Body
        
        '�d�����邩�ǂ����̃t���O
        Dim HasDuplicated
        HasDuplicated=TimeOrderedFolder.CheckDuplicate(TmpOneMailItem)
           
        
        '������͗ݐς̃J�E���g�ɂ�����J�E���g�ς݂��ǂ����������t���O
        '���Ƃ��A���J�E���g�^���Œ��̃��[����,�O��J�E���g�������[���̒��ōł��V��������(�ǂ̃��[���܂ł��J�E���g���������������߂̂���)���Â����̂ł������ꍇ,
        '����̓J�E���g���Ă��邱�ƂɂȂ�̂�,�d���ŃJ�E���g���Ȃ��悤�ɂ���
        Dim HasUnCounted
        HasUnCounted=False
        
         
        '�O��̃J�E���g���̍ŐV�̃��[���̓����i���̎������O�̓J�E���g���Ă��邱�Ƃ��������́j��荡�̃��[������̎����Ȃ�,
        '�悤�₭���J�E���g�Ƃ݂Ȃ����
        If LastCountMailDate < Time Then
          HasUnCounted=True
        End If
              
        If LastMailTime < Time Then
          LastMailTime=Time
        End If
        
            
        CurrentMailNum=receiveFolder.Items.Count
        CurrentDeletedFolderMailNum=deletedFolder.Items.Count
        MailOperation=DManager.GetState(Address,Name)
        
        '���̃��[���A�C�e����,�u�폜�ς݃t�H���_�v���邢�́u���łɃJ�E���g�ς�(���̃��[�v�Ō���)�̎�M�ς݃t�H���_�v���Ɋ��ɂ���΂���͏d��
        '�����͊��S�폜�Ɠ����ɂ���
        If HasDuplicated Then
          If MailOperation <> "�ۑ�" Then
            MailOperation="���S�폜"
          End If
          '���̃��[�����d���ł���Ȃ�������[���͐����Ă���Ƃ������ƂɂȂ�
          HasUnCounted=False
        End If
              
        '���[���̈������i����ɂ���ă��[�����ǂ�������)
           
        '�ۑ�����ꍇ
        Select  Case MailOperation
          Case  "�ۑ�"
           '���̃��[���͏������ɎQ�Ƃ��郁�[���̃C���f�b�N�X��1�i�߂�
            SaveMailNum=SaveMailNum+1
            DManager.Count Address,Name,Time,HasUnCounted,"ReceivedFolder"
              
          Case"���S�폜"
          
            ''VBS�̏ꍇ,Delete���\�b�h�͎�M�ς݃t�H���_�̃A�C�e�����폜����ƍ폜�ς݃t�H���_�ցA�폜�ς݃t�H���_�̃A�C�e�����폜����Ɗ��S�폜�����
            '�Ȃ̂ň�C�ɂ����Ȃ�,��M�ς݂̂��̂����S�폜���邱�Ƃ͂ł��Ȃ�
            '�䂦��,�܂��͂�������,�폜�ς݂ֈړ�����
            
            '��M�ς݂���폜�ς݂Ɉړ���,OneMailItem�Ƃ������[���A�C�e���I�u�W�F�N�g�ϐ��͎Q�Ƃł��Ȃ��Ȃ�(Null�ɂȂ��Ă��܂�)�̂�,
            '�����Ŏ���̃��[�����u����I�u�W�F�N�g������Ă���(�폜�ς݂���,���S�폜�����,���Ԗڂ̃A�C�e�����폜����΂悢�̂���m��Ȃ���΂Ȃ�Ȃ�����)
            Set TmpOneMailItem=New TmpMailItem
            TmpOneMailItem.SetMailInfo Address,Name,Time,Subject,Body
            DManager.Count Address,Name,Time,HasUnCounted,"CompDeleted"
            
            
            '�܂��͍폜�ς݂ֈړ�(�����OneMailItem�͎Q�Ƃł��Ȃ��Ȃ�)
            
            OneMailItem.Delete
            
            '�ړ�����������܂ő҂�
            Do While  CurrentMailNum = receiveFolder.Items.Count And  CurrentDeletedFolderMailNum = deletedFolder.Items.Count 
              Wscript.Sleep 500
            Loop
            
            
            '��M�ς݂���폜�ς݂ֈړ�����OneMailItem�Ƃ����ϐ��͂����g���Ȃ��̂�,��ňꎞ�I�ɕۑ����Ă��������[���������Ƃ�,
            '���ړ������A�C�e�����폜�ς݃t�H���_�̉��Ԗڂɂ��邩��T��(���`�T��)
            '������͏d������ł͂Ȃ��P�Ȃ�C���f�b�N�X�T���Ȃ̂Ő��`�ł悢
            Dim DeletedMailIndex
            DeletedMailIndex=MailIndexInFolder(TmpOneMailItem,deletedFolder)
            
            CurrentDeletedFolderMailNum=deletedFolder.Items.Count
            
            '�����Ŋ��S�폜(�폜�ς݃t�H���_������폜����)�A�C�e�������Ԗڂɂ��邩���������̂ł��̏������ɍēx���[���A�C�e���I�u�W�F�N�g�ϐ�������
            '���S�폜����Delete���\�b�h�̓��[���A�C�e���I�u�W�F�N�g�ϐ��ɕR�Â��Ă��郁�\�b�h�ł��邽��,���R�Ȃ����̋^�����[�����ł͍폜�ł��Ȃ�
            Dim CompDeleteMailItem
            Set CompDeleteMailItem=deletedFolder.Items.Item(DeletedMailIndex)
            
            '�폜���ł���悤�ɂȂ�܂ő҂�
            Do
             On Error Resume Next
              Err.Clear
              CompDeleteMailItem.Delete
             On Error GoTo 0
            Loop While Err.Number <> 0
               
            '�폜����������܂ő҂�
            Do While   CurrentDeletedFolderMailNum = deletedFolder.Items.Count
              Wscript.Sleep 500
            Loop
             
            
            Set TmpOneMailItem=Nothing             
         
             
          Case Else
            DManager.Count Address,Name,Time,HasUnCounted,"DeletedFolder"
            CurrentDeletedFolderMailNum=deletedFolder.Items.Count
            OneMailItem.Move deletedFolder
            '�ړ�����������܂ő҂�
            Do While  CurrentMailNum = receiveFolder.Items.Count And  CurrentDeletedFolderMailNum = deletedFolder.Items.Count 
              Wscript.Sleep 500
            Loop
        End Select
      End If
      
      Set OneMailItem=Nothing
        
    Loop
      
    If EnterFlag Then
      CountWaitSec=0
      EnterTime=EnterTime+1
      EnterFlag=False
      Dim Bias
      Bias=(EnterTime+1)\2
      CurrentNormSec=NormSec\Bias
      If CurrentNormSec < 1 Then
        CurrentNormSec=1
      End If
    End If
      
    Wscript.Sleep 500
    CountWaitSec=CountWaitSec+1
    
  Loop
  
  
  FManager.WriteResultDataManageObj DManager
  FManager.WriteRenewLastMailDate DManager.GetCountStartDate(),LastMailTime,DManager.GetSumMailNum("Cumulative"),DManager.GetSumMailNum("Exists"),DManager.GetSumMailNum("ReceivedFolder"),DManager.GetSumMailNum("DeletedFolder")
   
 
  Set OneMailItem=Nothing
  
  Set DManager=Nothing
  Set FManager=Nothing
  
  Dim Fold
  Set Fold=namespace.GetDefaultFolder(20)
  
  Do While Fold.Items.Count <> 0
    On Error Resume Next
     Fold.Items(1).Delete
     Wscript.Sleep 1000
     Err.Clear
    On Error GoTo 0
    
  Loop
  
  Dim Explorer
  Set Explorer=outlook.ActiveExplorer
  
  'Explorer���擾�ł��Ȃ�������,�����I��
  If Explorer Is Nothing Then
    'Outlook�̏I��
    Dim Locator
    Dim Service
    Dim oProc
    Dim oProcs
    
    Set Locator = WScript.CreateObject("WbemScripting.SWbemLocator")
    Set Service = Locator.ConnectServer
    Set oProcs = Service.ExecQuery("Select * From Win32_Process Where Description=""OUTLOOK.EXE""")
    
    For Each oProc In oProcs
       oProc.Terminate
    Next
    
  Else
    Explorer.Close
  End If
  
  
 
  
End Function

Main