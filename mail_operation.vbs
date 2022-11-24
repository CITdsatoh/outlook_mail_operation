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
  Public NumLogFile
  Public WshObj
  Public FDate
  
  Public Sub Class_Initialize()
   ReDim AddrNumLists(0)
   Set WshObj=Wscript.CreateObject("Wscript.Shell")
   Dim DesktopFolder
   DesktopFolder=WshObj.SpecialFolders(4)
   
   FileName=DesktopFolder&"\outlook_mail_dest_list.csv"
   Dim BackupFolder
   BackupFolder=WshObj.SpecialFolders(5)
   BackupFileName=BackupFolder&"\outlook_mail_dest_list.csv"
   LogFile=BackupFolder&"\datelog.log"
   NumLogFile=BackupFolder&"\mail_num.log"
   SetFileDate
  End Sub
  
  Public Sub Class_Terminate()
    Dim i
    For i=LBound(AddrNumLists) To UBound(AddrNumLists)
      Set AddrNumLists(i)=Nothing
    Next
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
      FObj.attributes=0
      FReading BackupFileName
    End If
    
   
  End Function
  
  Public Function  FReading(FName)
   Dim FReader
   
   Dim FileOpen
   FileOpen=True
   
   Set FReader=Wscript.CreateObject("ADODB.Stream")
   FReader.Type=2
   FReader.Charset="UTF-8"
   FReader.LineSeparator=10
   FReader.Open
  
   '�t�@�C������f�[�^���擾����ۂ�,���̃t�@�C�����̂��J����Ă������̓G���[���o��̂�,File��������܂ŉi�v���[�v����
   Do While FileOpen
     On Error Resume Next
      FReader.LoadFromFile FName
      
      '�G���[���Ȃ������i�t�@�C���������Ă�����),�����ɂ��ǂ蒅���̂�,FileOpen�t���O��������,���[�v�𔲂���
      If Err.Number = 0 Then
        FileOpen=False
      End If
     On Error GoTo 0 
     
   Loop
       
     
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
    Dim FileOpen
    FileOpen=True
    
    Set FReader=Wscript.CreateObject("ADODB.Stream")
    FReader.Type=2
    FReader.Charset="UTF-8"
    FReader.LineSeparator=10
    FReader.Open
    
    Do While FileOpen
      On Error Resume Next
        FReader.LoadFromFile LogFile
        If Err.Number = 0 Then
          FileOpen=False
        End If
      On Error GoTo 0 
    Loop
    
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
  
  Public Function GetSumMailNum()
   GetSumMailNum=0
   Dim i
   For i= 1 To UBound(AddrNumLists)
      GetSumMailNum=GetSumMailNum+AddrNumLists(i).GetNum()
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
    MailSum=GetSumMailNum()
    
    For i= 1 To UBound(AddrNumLists)
      FWriter.WriteText AddrNumLists(i).ToStr(),1
   Next
   
    
    FWriter.WriteText "���v,,,"&MailSum&",",1
    
    Dim FileOpen
    FileOpen=True
    
    Do While FileOpen
      On Error Resume Next
        FWriter.SaveToFile FileName,2
        If Err.Number = 0 Then
          FileOpen=False
        End If
      On Error GoTo 0 
    Loop
    
    FWriter.Close
    
    Set FWriter=Nothing
    
     
    Dim CpyFso
    Set CpyFso=Wscript.CreateObject("Scripting.FileSystemObject")
    '�o�b�N�A�b�v�t�@�C�����Ȃ��ꍇ�G���[���o��
    On Error Resume Next
      CpyFso.GetFile(BackupFileName).attributes=0
    On Error GoTo 0 
    
    FileOpen=True
    Do While FileOpen
      
      On Error Resume Next
        CpyFso.CopyFile FileName,BackupFileName,True
        If Err.Number = 0 Then
          FileOpen=False
        End If
        CpyFso.GetFile(BackupFileName).attributes=1
      On Error GoTo 0
    Loop
      
    Set CpyFso=Nothing
   
    
    
  End Function
  
  '�`�F�b�N�������t�����O�ɋL�^
  Public Function LogWrite(LastMailDate)
    Dim FWriter
    Set FWriter=Wscript.CreateObject("ADODB.Stream")
    Dim FileOpen
    FileOpen=True
    
    FWriter.Type=2
    FWriter.Charset="UTF-8"
    FWriter.Open
    FWriter.WriteText ""&LastMailDate,1
    
    If Wscript.CreateObject("Scripting.FileSystemObject").FileExists(LogFile) Then
      Wscript.CreateObject("Scripting.FileSystemObject").GetFile(LogFile).attributes=0
    End If
    
    Do While FileOpen
      On Error Resume Next
        FWriter.SaveToFile LogFile,2
        If Err.Number = 0 Then
          FileOpen=False
        End If
      On Error GoTo 0 
    Loop
    
    FWriter.Close
    Set FWriter=Nothing
    Wscript.CreateObject("Scripting.FileSystemObject").GetFile(LogFile).attributes=1
  End Function
  
  Public Function NumLogWrite(LastMailDate)
    Dim NowDate
    NowDate=Now
    
    Dim MailNum
    MailNum=GetSumMailNum()
    
    Dim FirstDate
    FirstDate=AddrNumLists(1).GetFirstDate
    
    Dim FWriter
    Set FWriter=Wscript.CreateObject("ADODB.Stream")
    
    Dim FSObj
    Set FSObj=Wscript.CreateObject("Scripting.FileSystemObject")
    
    Dim FileOpen
    FileOpen=True
    
    FWriter.Type=2
    FWriter.Charset="UTF-8"
    FWriter.Open
    FSobj.GetFile(NumLogFile).attributes=0
    
    If FSObj.FileExists(NumLogFile) Then
      FWriter.LoadFromFile NumLogFile
      FWriter.Position=FWriter.Size
      FWriter.WriteText Now&","&LastMailDate&","&MailNum,1
      Do While FileOpen
        On Error Resume Next
          FWriter.SaveToFile NumLogFile,2
          If Err.Number = 0 Then
            FileOpen=False
          End If
        On Error GoTo 0 
      Loop
      
    Else
      FWriter.WriteText "�v���O�������s����,���̎��_�ł̍ŐV�̃��[������,"&FirstDate&"���炻�̎��_�܂ł̗ݐσ��[����",1
      FWriter.WriteText Now&","&LastMailDate&","&MailNum,1
      Do While FileOpen
       On Error Resume Next
          FWriter.SaveToFile NumLogFile,2
          If Err.Number = 0 Then
            FileOpen=False
          End If
          
       On Error GoTo 0 
      Loop
    End If
    
    FWriter.Close
    Set FWriter=Nothing
    FSobj.GetFile(NumLogFile).attributes=1
    Set FSobj=Nothing
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
  
  
  
  '�e�X�g�p���S�폜�ς݃t�H���_�\(�e�X�g�̎��͂����Ȃ芮�S�폜���Ȃ�)
  Dim testCompDeleteFolder
  Set testCompDeleteFolder=deletedFolder.Folders("test_comp_delete")
  
  
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
  LastCountMailDate=CountManager.GetModifiedFileDate()
  
  
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
          
        If HasMailInFolder(OneMailItem,deletedFolder) Then
          On Error Resume Next
          OneMailItem.Delete
          'Do While  CurrentMailNum = receiveFolder.Items.Count 
            'Wscript.Sleep 1000
          'Loop
          On Error GoTo 0
        Else
          If LastCountMailDate < Time Then
            CountManager.Count Addr,Name,Time
          End If
            
          If LastMailTime < Time Then
            LastMailTime=Time
          End If
            
          CurrentMailNum=receiveFolder.Items.Count
          CurrentDeletedFolderMailNum=deletedFolder.Items.Count
          TestFolderMailNum=testCompDeleteFolder.Items.Count
          MailOperation=CountManager.GetState(Addr,Name)
              
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
  

  CountManager.FWrite()
 
  CountManager.LogWrite(LastMailTime)
  
  CountManager.NumLogWrite(LastMailTime)
     
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