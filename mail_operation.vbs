Option Explicit

Class AddrNumSet

  Public DstAddress
  Public DstName
  Public FirstDate
  Public mailNum
  Public State
  Public IsFirst
  Public DataPool ()
  
  Public Sub Class_Initialize()
    DstAddress=""
    DstName=""
    mailNum=0
    State="�폜�ς݂ֈړ�"
    ReDim DataPool(0)
    IsFirst=True
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  Public Function SetValue(Addr,Name,NumStr,MailState)
  
   DstAddress=Addr
   DstName=Name
   mailNum=CLng(NumStr)
   State=MailState
   
  End Function
  
  Public Function SetFirstDate(FDate)
    FirstDate=CDate(FDate)
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
   If Not IsEmpty(FirstDate) Then
    GetFirstDate=CDate(FirstDate)
    Exit Function
   End If
   
   GetFirstDate=DataPool(1)
   FirstDate=DataPool(1)
   Dim i
   For i=2 To UBound(DataPool)
     If DataPool(i) < GetFirstDate Then
       GetFirstDate=DataPool(i)
       FirstDate=DataPool(i)
     End If
   Next
   
  End Function
  
  Public Function AddDataPool(NewDate)
    If IsEmpty(FirstDate) Then
     ReDim Preserve DataPool(UBound(DataPool)+1)
     DataPool(UBound(DataPool))=NewDate
    End If
  End Function
  
  Public Function ToStr()
   ToStr=DstAddress&","&DstName&","&FirstDate&","&mailNum&","&State
  End Function

End Class

  
Class Manager

  Public AddrNumLists ()
  Public FileName
  Public BackupFileName
  Public LogFile
  Public WshObj
  Public FDate
  
  Public Sub Class_Initialize()
   ReDim AddrNumLists(0)
   FileName="outlook_mail_dest_list.csv"
   Set WshObj=Wscript.CreateObject("Wscript.Shell")
   Dim BackupFolder
   BackupFolder=WshObj.SpecialFolders(5)
   BackupFileName=BackupFolder&"\"&FileName
   LogFile=BackupFolder&"\datelog.log"
   SetFileDate
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  '���[�����J�E���g����(�J�E���g��0�A�܂�܂����̈��悩��̃��[�������݂��Ȃ��ꍇ��,�V�����I�u�W�F�N�g�����)
  Public Function Count(addr,name,date)
   Dim i
   For i=1 To UBound(AddrNumLists)
     If addr = AddrNumLists(i).GetAddress And name = AddrNumLists(i).GetName Then
       AddrNumLists(i).NumIncrement()
       AddrNumLists(i).AddDataPool date
       Exit Function
     End If
   Next
   
   ReDim Preserve AddrNumLists (UBound(AddrNumLists)+1)
   Dim NumObj
   Set NumObj=new AddrNumSet
   NumObj.SetValue addr,name,"1","�폜�ς݂ֈړ�"
   NumObj.AddDataPool date
   Set AddrNumLists(UBound(AddrNumLists))=NumObj
   
  End Function
  
  '���̈��悩��̃��[���̐���Ԃ�
  Public Function getNum(addr,name)
   For i=LBound(AddrNumLists) To UBound(AddrNumLists)
     If addr = AddrNumLists(i).GetAddress And AddrNumLists(i).getName = name Then
       getNum=AddrNumLists(i).getNum()
       Exit Function
     End If
   Next
   getNum=0
  End Function
  
  '���̈���̃��[�����ǂ̂悤�Ɉ�����(�폜�ς݂ֈړ��i�f�t�H���g)���A�ۑ���,���S�폜��)
  Public Function GetState(addr,name)
    Dim i
    For i=1 To UBound(AddrNumLists)
     If addr = AddrNumLists(i).GetAddress And name = AddrNumLists(i).GetName Then
       GetState=AddrNumLists(i).GetMailState()
       Exit Function
     End If
   Next
   GetState="�폜�ς݂ֈړ�"
  End Function
  
  '���[���𒲂ׂ�ɂ�����,�ۑ����[���̏d���J�E���g������邽��,�O��A���[���̐����J�E���g�����̂͂��Ȃ̂��𓾂�
  '�ۑ����[���ɓ����Ă���,���̓��t���O�̃��[���Ɋւ��Ă͂��łɃJ�E���g���Ă���̂ŃJ�E���g���Ȃ�
  Public Function GetModifiedFileDate()
    GetModifiedFileDate=FDate
  End Function
 
  Public Function SetFileDate()
   Dim FObj
   Dim FSObj
   Set FSObj=CreateObject("Scripting.FileSystemObject")
   '���O�t�@�C������Ō�Ƀ��[�����`�F�b�N�������t���𓾂�
   If FSObj.FileExists(LogFile) Then
     FSObj.GetFile(LogFile).attributes=0
     FDate=GetRealModifiedDate
   ElseIf FSObj.FileExists(FileName) Then
     'Log�t�@�C�����Ȃ������ꍇ���̓��t�ő�p����
     FDate=FSObj.GetFile(FileName).DateLastModified
   ElseIf FSObj.FileExists(BackupFileName) Then
     FDate=FSObj.GetFile(BackupFileName).DateLastModified
   Else
     FDate=CDate("1970/1/1")
   End If
  
  End Function
     
   
  Public Function FRead()
    Dim FObj
    Dim FSObj
    Set FSObj=CreateObject("Scripting.FileSystemObject")
    If FSObj.FileExists(FileName) Then
      Set FObj=FSObj.GetFile(FileName)
      FReading FileName
    ElseIf FSObj.FileExists(BackupFileName) Then
      Set FObj=FSObj.GetFile(BackupFileName)
      FSObj.attributes=0
      FReading BackupFileName
    End If
    
   
  End Function
  
  Public Function  FReading(FName)
   Dim FReader
    
   Set FReader=Wscript.CreateObject("ADODB.Stream")
   FReader.Type=2
   FReader.Charset="UTF-8"
   FReader.LineSeparator=10
   FReader.Open
   FReader.LoadFromFile FName
     
   '1�s�ڂ̓w�b�_�Ȃ̂ŏ��Ƃ��ĕK�v�Ȃ��B
   '�Ƃ肠����,1��S���̍s�ɂ��ăf�[�^���擾����
   Dim AllData ()
   Dim LineNum
   Dim OneLine
   LineNum=0
   Do While FReader.EOS = False
     ReDim Preserve AllData(LineNum)
     OneLine=FReader.ReadText(-2)
     AllData(LineNum)=OneLine
     LineNum=LineNum+1
   Loop
     
   FReader.Close
   Set FReader=Nothing
   Dim OneData
   Dim NumObj
   Dim i
   Dim AllDataWithoutBr
     
   '�w�b�_�͂���Ȃ��̂�1�Ԗڂ���擾����B�����čŌ�̍s�͍��v�Ȃ̂ł��������Ȃ�
   For i=1 To UBound(AllData)-1
     AllDataWithoutBr=Replace(AllData(i),VbCr,"")
     OneData=Split(AllDataWithoutBr,",")
     ReDim Preserve AddrNumLists(UBound(AddrNumLists)+1)
     Set NumObj=new AddrNumSet
     If UBound(OneData) = 4 Then
       NumObj.SetValue OneData(0),oneData(1),OneData(3),OneData(4)
     ElseIf UBound(OneData) >= 3 Then
       NumObj.SetValue OneData(0),oneData(1),OneData(3),"�폜�ς݂ֈړ�"
     End If
     NumObj.SetFirstDate OneData(2)
     Set AddrNumLists(UBound(AddrNumLists))=NumObj
   Next
      
  End Function
  
  Public Function GetRealModifiedDate()
    Dim DateDataStr
    Dim FReader
    Set FReader=Wscript.CreateObject("ADODB.Stream")
    FReader.Type=2
    FReader.Charset="UTF-8"
    FReader.LineSeparator=10
    FReader.Open
    FReader.LoadFromFile LogFile
    Do While FReader.EOS = False
      DateDataStr=FReader.ReadText(-2)
      Exit Do
    Loop
    
    FReader.Close
    Set FReader=Nothing
    getRealModifiedDate=CDate(Replace(DateDataStr,VbCr,""))
    
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
  
  Public Function FWrite()
  
    DateSort()

    
    Dim Header
    Dim Today
    Dim FWriter
    Today=Date()
    Dim FirstDate
    FirstDate=AddrNumLists(1).GetFirstDate
    
    Header="���[���A�h���X,����,"&FirstDate&"�ȍ~�ōł��������̈��悩�烁�[�����͂������t,"&FirstDate&"����"&Today&"�܂łɓ͂������[���̐�,���[���̎�舵��"
    Set FWriter=Wscript.CreateObject("ADODB.Stream")
    FWriter.Type=2
    FWriter.Charset="UTF-8"
    FWriter.Open
    FWriter.WriteText Header,1
    
    Dim i
    
    '���v���[����(1�ԍŌ�̍s�ɏ����Ă���)
    Dim MailSum
    MailSum=0
    For i= 1 To UBound(AddrNumLists)
      MailSum=MailSum+AddrNumLists(i).GetNum()
      FWriter.WriteText AddrNumLists(i).ToStr(),1
    Next
    
    FWriter.WriteText "���v,,,"&MailSum&",",1
    
    FWriter.SaveToFile FileName,2
    FWriter.Close
    
    Set FWriter=Nothing
     
    Dim CpyFso
    Set CpyFso=Wscript.CreateObject("Scripting.FileSystemObject")
    '�o�b�N�A�b�v�t�@�C�����Ȃ��ꍇ�G���[���o��
    On Error Resume Next
      CpyFso.GetFile(BackupFileName).attributes=0
    On Error Goto 0
    
    CpyFso.CopyFile FileName,BackupFileName,True
    CpyFso.GetFile(BackupFileName).attributes=1
    Set CpyFso=Nothing
   
    For i=LBound(AddrNumLists) To UBound(AddrNumLists)
      Set AddrNumLists(i)=Nothing
    Next
    
  End Function
  
  '�`�F�b�N�������t�����O�ɋL�^
  Public Function LogWrite(LastMailDate)
    Dim FWriter
    Set FWriter=Wscript.CreateObject("ADODB.Stream")
    FWriter.Type=2
    FWriter.Charset="UTF-8"
    FWriter.Open
    FWriter.WriteText ""&LastMailDate,1
    FWriter.SaveToFile LogFile,2
    FWriter.Close
    Set FWriter=Nothing
    Wscript.CreateObject("Scripting.FileSystemObject").GetFile(LogFile).attributes=1
  End Function
  
  
      

End Class

Function MailEquals(One,Other)
 
 MailEquals=True
 If One.SenderEMailAddress <> Other.SenderEMailAddress Then
   MailEquals=False
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
 
End Function

Function HasMailInFolder(MailItem,Folder)
  Dim i
  For i=1 To Folder.Items.Count
   If MailEquals(MailItem,Folder.Items(i)) Then
     HasMailInFolder=True
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
  
  
  'outlook���ŏ������ċN��(�����outlook���N�����Ă��邱�Ƃ��킩��Ȃ��悤�ɂ���)
  Set Wsh=Wscript.CreateObject("Wscript.shell")
  Wsh.Run "outlook.exe",7,False


  '���[���̃J�E���g�Ȃǂ��s���Ǘ��N���X
  Dim CountManager

  Set CountManager=new Manager
  
  
  '�t�@�C���Ǎ�
  CountManager.FRead()
  
 
  
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
 
 
   
  '�܂��͍폜�ς݃A�C�e���ɂ��Ă݂Ă䂭
  Do While SaveMailNum < deletedFolder.Items.Count
     CurrentMailNum=deletedFolder.Items.Count
     Set OneMailItem=deletedFolder.Items.Item(SaveMailNum+1)
     MailOperation=CountManager.GetState(OneMailItem.SenderEMailAddress,OneMailItem.SenderName)
     If MailOperation = "�ۑ�" Then
       OneMailItem.Move receiveFolder
       Do While  CurrentMailNum = deletedFolder.Items.Count 
           Wscript.Sleep 1000
       Loop
     ElseIf MailOperation = "���S�폜" Then
       OneMailItem.Delete
       Do While  CurrentMailNum = deletedFolder.Items.Count 
           Wscript.Sleep 1000
       Loop
     Else
       SaveMailNum=SaveMailNum+1
     End If
  Loop
   
  
  '�폜�ς݃A�C�e���̕��������̂Ŏ��͎�M�ς݂̕�������
  Set OneMailItem=Nothing
  '�ۑ��i�t�H���_�L�[�v�j�̃��[������0�ɖ߂�
  SaveMailNum=0
  MailOperation=""
 
  '���͎�M�A�C�e���̒������Ă䂭��,��M�A�C�e���̓ǂݍ��݂Ƃ��̏����͔񓯊��ł��邱�Ƃ��班���^�C�����O��݂���K�v������
  '�^�C�����O�p�̕ϐ�
  Dim NormSec
  NormSec=30
 
 
  Dim CountTimes
  CountTimes=0
 
  Dim CountWaitSec
  CountWaitSec=0
 
  Dim EnterFlag
  EnterFlag=False
 
  '�Ō�ɂ����[���̃J�E���g���s�����̂��𓾂�(�ۑ��t�H���_����d���J�E���g�����Ȃ��悤��)
  Dim LastCountMailDate
  LastCountMailDate=CountManager.GetModifiedFileDate()
  
  Dim TestFolder
  Set TestFolder=deletedFolder.Folders("test")
  
  
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
 
  Do While CountWaitSec < NormSec 
    Do While SaveMailNum < receiveFolder.Items.Count
      '���[��������(�폜�ς݃t�H���_�j�Ɉړ��������,��M�g���[�̃��[��������̂�,������,�����C���f�b�N�X���A�N�Z�X����
      EnterFlag=True
      CurrentMailNum=receiveFolder.Items.Count
      Set OneMailItem=receiveFolder.Items.Item(SaveMailNum+1)
      Name=OneMailItem.SenderName
      Addr=OneMailItem.SenderEmailAddress
      Time=OneMailItem.ReceivedTime
      
      If HasMailInFolder(OneMailItem,deletedFolder) Then
         TestFolderMailNum=TestFolder.Items.Count
         OneMailItem.Move TestFolder
         Do While  TestFolderMailNum = TestFolder.Items.Count
           Wscript.Sleep 1000
           Wscript.Echo "�ҋ@��"
         Loop
      Else
        If LastCountMailDate < Time Then
          CountManager.Count Addr,Name,Time
        End If
        
        If LastMailTime < Time Then
          LastMailTime=Time
        End If
        
       
        MailOperation=CountManager.GetState(Addr,Name)
          
        '���[���̈������i����ɂ���ă��[�����ǂ�������)
        '�ۑ�����ꍇ
        If MailOperation = "�ۑ�" Then
          '���̃��[���͏������ɎQ�Ƃ��郁�[���̃C���f�b�N�X��1�i�߂�
          SaveMailNum=SaveMailNum+1
        ElseIf MailOperation = "���S�폜" Then
          OneMailItem.Delete
          '�폜����������܂ő҂�
          Do While  CurrentMailNum = receiveFolder.Items.Count 
            Wscript.Sleep 1000
          Loop
        Else
          OneMailItem.Move deletedFolder
          '�ړ�����������܂ő҂�
          Do While  CurrentMailNum = receiveFolder.Items.Count
            Wscript.Sleep 1000
          Loop
        End If
      End If 
       
        
    Loop
      
    If EnterFlag Then
      CountWaitSec=0
      If CountTimes <= 100 Then
        CountTimes=CountTimes+1
      End If
      Dim Bias
      Bias=(CountTimes\5)+1
      NormSec=30\Bias
      EnterFlag=False
    End If
      
    Wscript.Sleep 1000
      
    CountWaitSec=CountWaitSec+1
    
  Loop
  
  '���ʂ��t�@�C���ɋL��
  CountManager.FWrite()
  CountManager.LogWrite(LastMailTime)
  
  Set OneMailItem=Nothing
  
  Set CountManager=Nothing
  
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