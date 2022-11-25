Option Explicit

Class AddrNumSet

  Public DstAddress
  Public DstName
  Public FirstDate
  Public mailNum
  Public State
  
  Public Sub Class_Initialize()
    DstAddress=""
    DstName=""
    mailNum=0
    State="�폜�ς݂ֈړ�"
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  Public Function SetValue(Addr,Name,NumStr,MailState)
  
   DstAddress=Addr
   DstName=Name
   mailNum=CLng(NumStr)
   State=MailState
   
  End Function
  
  Public Function AddDate(FDate)
    If IsEmpty(FirstDate) Then
      FirstDate=FDate
    ElseIf FDate < FirstDate Then
      FirstDate=FDate
    End If
  End Function
  
  Public Function NumIncrement()
   mailNum=mailNum+1
  End Function
  
  Public Function GetAddress()
   GetAddress=DstAddress
  End Function
  
  Public Function GetNum()
   GetNum=mailNum
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
   ToStr=DstAddress&","&DstName&","&FirstDate&","&mailNum&","&State
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
    ReDim Preserve AddrNumLists(UBound(AddrNumList+1))
    Dim NewObj
    Set NewObj=new AddrNumSet
    If UBound(OneData) = 4 Then
       Select Case OneData(4)
        Case "�ۑ�","�폜�ς݂ֈړ�","���S�폜"
          NumObj.SetValue OneData(0),oneData(1),OneData(3),OneData(4)
        Case Else
           NumObj.SetValue OneData(0),oneData(1),OneData(3),"�폜�ς݂ֈړ�"
       End Select
    ElseIf UBound(OneData) >= 3 Then
       NumObj.SetValue OneData(0),OneData(1),OneData(3),"�폜�ς݂ֈړ�"
    End If
    On Error Resume Next
     NewObj.AddDate CDate(OneData(2))
    On Error GoTo 0
    AddrNumLists(UBound(AddrNumLists))=NewObj
    
   End Function
    
   '���[���A�h���X�ƈ������L�[�Ƃ���,���̃��[���A�h���X�i�����j�̏�񂪊Ǘ��z��̂ǂ��̃C���f�b�N�X�ɂ��邩������
   Public Function DataIndex(Address,Name)
    Dim i
    For i=1 To UBound(AddrNumLists)
      If And AddrNumLists(i).GetAddress = Address AddrNumLists(i).GetName = Name Then
        DataIndex=i
        Exit Function
      End If
    Next
    DataIndex=-1
   End Function
   
   '������͗^����ꂽ���[���A�h���X�ƈ�������̃��[�������݂��邩�ǂ�����Ԃ�
   '�܂�,��̃��\�b�h�̖߂�l��-1�̎��͂��̈��悩��̃��[���͂Ȃ�.����ȊO�̎��͂���Ƃ�������
   Public Function DataExists(Address,Name)
     If DataIndex(Address,Name) <> -1 Then
       DataExists=True
       Exit Function
     End If
     DataExists=False
   End Function
   
   '���[�����J�E���g����(�J�E���g��0�A�܂�܂����̈��悩��̃��[�������݂��Ȃ��ꍇ��,�V�����I�u�W�F�N�g�����)
   Public Function Count(Address,Name,Date)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      AddrNumLists(index).NumIncrement()
      AddrNumLists(index).AddDate Date
      Exit Function
    End If
    
    ReDim Preserve AddrNumLists(UBound(AddrNumLists)+1)
    Dim NewObj
    Set NewObj=New AddrNumSet
    NumObj.SetValue Addr,Name,"1","�폜�ς݂ֈړ�"
    NumObj.AddDate Date
    Set AddrNumLists(UBound(AddrNumLists))=NumObj
   End Function
   
   '�����_�ł̍��v���[����
   Public Function GetSumMailNum()
    GetSumMailNum=0
    Dim i
    For i= 1 To UBound(AddrNumLists)
      GetSumMailNum=GetSumMailNum+AddrNumLists(i).GetNum()
    Next
  End Function
  
  '���[���̏W�v���J�n�������̓��t���擾
  Public GetCountStartDate()
    DateSort
    GetCountStartDate=AddrNumLists(1).GetFirstDate
  End Function
  
  '�^����ꂽ���[���A�h���X�i�����j����̃��[���̐���Ԃ�
  Public Function GetNum(Address,Name)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      GetNum=AddrNumLists(index).getNum()
      Exit Function
    End If
    
    GetNum=0
  End Function
  
  '���̈���̃��[�����ǂ̂悤�Ɉ�����(�폜�ς݂ֈړ��i�f�t�H���g)���A�ۑ���,���S�폜��)
  Public Function GetState(Address,Name)
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
  
  '���ۂɃt�@�C���ɏ������ނƂ��̕����񂪗v�f�ƂȂ����z��𐶐�����
  Public Function ToFileWriteStr()
   
    DateSort()
    Dim Header
    Dim Today
    Dim FWriter
    Today=Date()
    Dim FirstDate
    FirstDate=GetCountStartDate
    
    Dim Header
    Header="���[���A�h���X,����,"&FirstDate&"�ȍ~�ōł��������̈��悩�烁�[�����͂������t,"&FirstDate&"����"&Today&"�܂łɓ͂������[���̐�,���[���̎�舵��"
    
    Dim Content
    ReDim Preserve Content(0)
    Content(0)=Header
    Dim i
    For i=1 To UBound(AddrNumLists)
     ReDim Preserve Content(UBound(Content)+1)
     Content(UBound(Content))=AddrNumLists(i).ToStr
    Next
    
    ReDim Preserve Content(UBound(Content)+1)
    Dim Footer
    Dim MailSum
    MailSum=GetSumMailNum()
    Footer="���v,,,"&MailSum&","
    Content(UBound(Content))=Footer
    
    Set ToFileWriteStr=Content
    
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
     On Error GoTo 0 
     
   Loop
   
   Dim Result
   ReDim Result()
   Dim LineNum
   LineNum=0
   Dim OneLine
   
   Do While FIOOperator.EOS = False
     ReDim Preserve Result(LineNum)
     OneLine=FIOOperator.ReadText(-2)
     Result(LineNum)=OneLine
     LineNum=LineNum+1
   Loop
   
   Set FRead=Result
     
   FIOOperator.Close
   Set FIOOperator=Nothing
  
  End Function 
  
  '�t�@�C���̏�������(Mode��"w"�Ȃ�㏑��,"a"�Ȃ�ǋL,WriteType��"Array"�Ȃ�z��Ƃ���,�e�v�f��1�s�������Ă䂭,"Str"�Ȃ當����Ƃ���1�s��������)
  Public Function FWrite(FilePath,Contents,Mode,WriteType)
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
        On Error GoTo 0
       Loop
       FIOOperator.Position=FWriter.Size
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
   Dim BackupFolder
   BackupFolder=WshObj.SpecialFolders(5)
   BackupFileName=BackupFolder&"\outlook_mail_dest_list.csv"
   BackupLogFolder=BackupFolder&"\backup"
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
   ElseIf FSobj.FileExists(BackupFolder) Then
     Dim FObj
     Set FObj=FSObj.GetFile(BackupFileName)
     ChangeFileAttributes(BackupFileName)
     FileContents=FIO.FRead(BackuplFileName)
     ChangeFileAttributes(BackupFileName)
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
     OneData=Split(AllDataWithoutBr,",")
     GetDataManageObj.ParseDataFromFileContent OneData
   Next
   
  End Function
  
  '���[���𒲂ׂ�ɂ�����,�ۑ����[���̏d���J�E���g������邽��,�O��A���[���̐����J�E���g�����̂͂��Ȃ̂��𓾂�
  '�ۑ����[���ɓ����Ă���,���̓��t���O�̃��[���Ɋւ��Ă͂��łɃJ�E���g���Ă���̂ŃJ�E���g���Ȃ�
  Public Function GetLastMailDate()
   '���O�t�@�C������Ō�Ƀ��[�����`�F�b�N�������t���𓾂�
   If FSObj.FileExists(TimeLogFile) Then
    ChangeFileAttributes(TimeLogFile)
    Dim Contents
    Contents=FIO.FRead(TimeLogFile)
    ChangeFileAttributes(TimeLogFile)
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
   Copy OriginalFileName,BackUpFileName,True
  End Function
  
  '���񃁁[���𒲂ׂăJ�E���g����ɂ�����,���񉽎������̃��[���܂ł��J�E���g�ς݂Ȃ̂����L�^���Ă���
  '��L�̂悤�Ɏ���A�ǂ̃��[������J�E���g����΂悢�̂����������߁i�d���J�E���g������邽�߁j
  Public Function WriteRenewLastMailDate(CountStartDate,LastMailDate,MailNum)
   ChangeFileAttributes(TimeLogFile)
   FIO.FWrite TimeLogFile,""&LastMailDate,"w","Str"
   ChangeFileAttributes(TimeLogFile)
   
   Dim NowDate
   NowDate=Now
   
   If FSObj.FileExists(NumLogFile) Then
     Dim Content=""
     Content=""&NowDate&","&LastMailDate&","&MailNum
     ChangeFileAttributes(NumLogFile)
     FIO.FWrite NumLogFile,Content,"a","Str"
     ChangeFileAttributes(NumLogFile)
   Else
    Dim Header
    Header="�v���O�������s����,���̎��_�ł̍ŐV�̃��[������,"&CountStartDate&"���炻�̎��_�܂ł̗ݐσ��[����"
    Dim Body
    Body=""&NowDate&","&LastMailDate&","&MailNum
    Dim Contents(1)
    Contents(0)=Header
    Contents(1)=Content
    ChangeFileAttributes(NumLogFile)
    FIO.FWrite NumLogFile,Contents,"w","Array"
    ChangeFileAttributes(NumLogFile)
   End If
   
   Dim yymmddhhmmssStr
   yymmddhhmmssStr=ToyymmddhhmmssStr(NowDate)
   Dim BackupSaveLogFile
   BackupSaveLogFile=BackupLogFolder&"\outlook_mail_dest_list_"&yymmddhhmmssStr&"_backup.csv"
   Copy OriginalFileName,BackupSaveLogFile,True
  End Function
    
  
  Public Copy(Src,Dst,Block)
    If Block Then
      ChangeFileAttributes(Dst)
    End If
    
    FSObj.CopyFile Src,Dst,True
    
    If Block Then
      ChangeFileAttributes(Dst)
    End If
  End Function
  
  Public Function ToyymmddhhmmssStr(NowDate)
   ToyymmddhhmmssStr=Year(NowDate)&PadZero(Month(NowDate),2)&PadZero(Date(NowDate),2)&PadZero(Hour(NowDate),2)&PadZero(Minute(NowDate),2)&PadZero(Second(NowDate),2)
  End Function
  
  Public PadZero(Before,Num)
    Dim BeforeInt
    BeforeInt=CLng(Num)
    Dim Result
    Result=""
    Dim Digit
    Digit=1
    Dim DigitNum
    Dim i
    For i=1 To Num
      DigitNum=((BeforeInt\Digit) Mod 10)
      Result=""&DigitNum&Result
      Digit=Digit*10
    Next
    PadZero=Result
  End Function
     
  Public ChangeFileAttributes(FName)
    Dim FObj
    If FSObj.FileExists(FName) Then
     Set FObj=FSObj.GetFile(FName)
     Select Case FObj.attributes
       Case 0
         FObj.attributes=1
       Case Else
         FObj.attributes=0
     End Select
     Set FObj=Nothing
    End If
  End Function
       
End Class
  
Function MailEquals(One,Other)
 
 On Error Resume Next
  If One.SenderEMailAddress <> Other.SenderEMailAddress Then
    MailEquals=False
    Exit Function
  End If
  
  If Err.Number <> 0 Then
    Exit Function
  End If
  
  If One.SenderName <> Other.SenderName Then
    MailEquals=False
    Exit Function
  End If
  
  If One.ReceivedTime <> Other.ReceivedTime Then
    MailEquals=False
    Exit Function
  End If
  
  If One.Subject <> Other.Subject Then
    MailEquals=False
    Exit Function
  End If
  
 On Error GoTo 0
 
 MailEquals=True
 
End Function

Function HasMailInFolder(MailItem,Folder)
  Dim i
  Dim Result
  For i=1 To Folder.Items.Count
   Result=MailEquals(MailItem,Folder.Items(i))
   If Result Then
     HasMailInFolder=True
     Exit Function
   ElseIf IsEmpty(Result) Then
     Exit Function
   End If
  Next
  
  HasMailInFolder=False

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
  
  
  
  '�e�X�g�p���S�폜�ς݃t�H���_�\(�e�X�g�̎��͂����Ȃ芮�S�폜���Ȃ�)
  Dim testCompDeleteFolder
  Set testCompDeleteFolder=deletedFolder.Folders("test_comp_delete")
  
  
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
 
 
   
  '�܂��͍폜�ς݃A�C�e���ɂ��Ă݂Ă䂭
  Do While SaveMailNum < deletedFolder.Items.Count
     CurrentMailNum=deletedFolder.Items.Count
     Set OneMailItem=deletedFolder.Items.Item(SaveMailNum+1)
     MailOperation=CountManager.GetState(OneMailItem.SenderEMailAddress,OneMailItem.SenderName)
     Select Case MailOperation
       Case "�ۑ�"
        OneMailItem.Move receiveFolder
        Do While  CurrentMailNum = deletedFolder.Items.Count 
            Wscript.Sleep 1000
        Loop
       Case "���S�폜"
        'OneMailItem.Delete
        TestDeleteMailNum=testCompDeleteFolder.Items.Count
        OneMailItem.Move testCompDeleteFolder
        'Do While  CurrentMailNum = deletedFolder.Items.Count 
        Do While TestDeleteMailNum = testCompDeleteFolder.Items.Count
            Wscript.Sleep 1000
        Loop
       Case Else
        SaveMailNum=SaveMailNum+1
    End Select
  Loop
   
  
  '�폜�ς݃A�C�e���̕��������̂Ŏ��͎�M�ς݂̕�������
  Set OneMailItem=Nothing
  '�ۑ��i�t�H���_�L�[�v�j�̃��[������0�ɖ߂�
  SaveMailNum=0
  MailOperation=""
 
  '���͎�M�A�C�e���̒������Ă䂭��,��M�A�C�e���̓ǂݍ��݂Ƃ��̏����͔񓯊��ł��邱�Ƃ��班���^�C�����O��݂���K�v������
  '�^�C�����O�p�̕ϐ�
  Dim NormSec
  NormSec=60
  
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
 
  '�Ō�ɂ����[���̃J�E���g���s�����̂��𓾂�(�ۑ��t�H���_����d���J�E���g�����Ȃ��悤��)
  Dim LastCountMailDate
  LastCountMailDate=DManager.GetLastMailDate()
  
  
  '�J�E���^�ϐ�
  Dim i
  
  '���[���̏��
  Dim Addr
  Dim Name
  Dim Time
  CurrentMailNum=receiveFolder.Items.Count
  
  Dim TestFolderMailNum
  
  
  Dim LastMailTime
  LastMailTime=LastCountMailDate
  
  Dim HasError
  HasError=0
  
  Dim CurrentDeletedFolderMailNum
 
  Do While CountWaitSec < CurrentNormSec 
    Do While SaveMailNum < receiveFolder.Items.Count
      '���[��������(�폜�ς݃t�H���_�j�Ɉړ��������,��M�g���[�̃��[��������̂�,������,�����C���f�b�N�X���A�N�Z�X����
      EnterFlag=True
      
      On Error Resume Next
       Set OneMailItem=receiveFolder.Items.Item(SaveMailNum+1)
       HasError=Err.Number
       If HasError <> 0 Then
         Wscript.Echo HasError
         Wscript.Echo "�G���["
       End If
      On Error GoTo 0
      
      If HasError = 0 Then
      
        Name=OneMailItem.SenderName
        Addr=OneMailItem.SenderEmailAddress
        Time=OneMailItem.ReceivedTime
        
        Dim  MailAlreadyHas
        MailAlreadyHas=HasMailInFolder(OneMailItem,deletedFolder)
          
        If MailAlreadyHas Then
          On Error Resume Next
          OneMailItem.Delete
          'Do While  CurrentMailNum = receiveFolder.Items.Count 
            'Wscript.Sleep 1000
          'Loop
          On Error GoTo 0
        Else If Not IsEmpty(MailAlreadyHas) Then
          If LastCountMailDate < Time Then
            DManager.Count Addr,Name,Time
          End If
            
          If LastMailTime < Time Then
            LastMailTime=Time
          End If
            
          CurrentMailNum=receiveFolder.Items.Count
          CurrentDeletedFolderMailNum=deletedFolder.Items.Count
          TestFolderMailNum=testCompDeleteFolder.Items.Count
          MailOperation=DManager.GetState(Addr,Name)
              
          '���[���̈������i����ɂ���ă��[�����ǂ�������)
           
          '�ۑ�����ꍇ
          Select  Case MailOperation
            Case  "�ۑ�"
             '���̃��[���͏������ɎQ�Ƃ��郁�[���̃C���f�b�N�X��1�i�߂�
             SaveMailNum=SaveMailNum+1
              
            Case"���S�폜"
             'On Error Resume Next
             'OneMailItem.Delete
             'On Error GoTo 0
             OneMailItem.Move testCompDeleteFolder
             '�폜����������܂ő҂�
             Do While  CurrentMailNum = receiveFolder.Items.Count  And TestFolderMailNum = testCompDeleteFolder.Items.Count
               Wscript.Sleep 1000
             Loop
             
            Case Else
              CurrentDeletedFolderMailNum=deletedFolder.Items.Count
              OneMailItem.Move deletedFolder
              '�ړ�����������܂ő҂�
              Do While  CurrentMailNum = receiveFolder.Items.Count And  CurrentDeletedFolderMailNum = deletedFolder.Items.Count 
                Wscript.Sleep 1000
              Loop
            End Select
        End If
      End If
      
     
       
    Loop
      
    If EnterFlag Then
      CountWaitSec=0
      EnterTime=EnterTime+1
      EnterFlag=False
      Dim Bias
      Bias=(EnterTime+1)\2
      CurrentNormSec=NormSec\Bias
    End If
      
    Wscript.Sleep 1000
    
      
    CountWaitSec=CountWaitSec+1
    
  Loop
  

  FManager.WriteResultDataManageObj DManager
  FManager.WriteRenewLastMailDate DManager.GetrCountStartDate(),LastMailDate,DManager.GetSumMailNum()
   
 
  Set OneMailItem=Nothing
  
  Set DManager=Nothing
  Set FManager=Nothing
  
  Dim Fold
  Set Fold=namespace.GetDefaultFolder(20)
  
  Do While Fold.Items.Count <> 0
    Fold.Items(1).Delete
    Wscript.Sleep 1000
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