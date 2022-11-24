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
    State="削除済みへ移動"
    ReDim DataPool(0)
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
    ReDim AddrNumLists(0)
   End Sub
   
   Public Sub Class_Terminate()
    Dim i
    For i=LBound(AddrNumLists) To UBound(AddrNumLists)
      Set AddrNumLists(i)=Nothing
    Next
   End Sub
   
   Public Function ParseDataFromFileContent(OneData)
    ReDim Preserve AddrNumLists(UBound(AddrNumList+1))
    Dim NewObj
    Set NewObj=new AddrNumSet
    If UBound(OneData) = 4 Then
       Select Case OneData(4)
        Case "保存","削除済みへ移動","完全削除"
          NumObj.SetValue OneData(0),oneData(1),OneData(3),OneData(4)
        Case Else
           NumObj.SetValue OneData(0),oneData(1),OneData(3),"削除済みへ移動"
       End Select
    ElseIf UBound(OneData) >= 3 Then
       NumObj.SetValue OneData(0),OneData(1),OneData(3),"削除済みへ移動"
    End If
    On Error Resume Next
     NewObj.AddDate CDate(OneData(2))
    On Error GoTo 0
    AddrNumLists(UBound(AddrNumLists))=NewObj
    
   End Function
    
    
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
   
   Public Function DataExists(Address,Name)
     If DataIndex(Address,Name) <> -1 Then
       DataExists=True
       Exit Function
     End If
     DataExists=False
   End Function
   
   Public Function Count(Address,Name,Date)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      AddrNumLists(index).NumIncrement()
      AddrNumLists(index).AddData Date
      Exit Function
    End If
    
    ReDim Preserve AddrNumLists(UBound(AddrNumLists)+1)
    Dim NewObj
    Set NewObj=New AddrNumSet
    NumObj.SetValue Addr,Name,"1","削除済みへ移動"
    NumObj.AddData Date
    Set AddrNumLists(UBound(AddrNumLists))=NumObj
   End Function
   
   Public Function GetSumMailNum()
    GetSumMailNum=0
    Dim i
    For i= 1 To UBound(AddrNumLists)
      GetSumMailNum=GetSumMailNum+AddrNumLists(i).GetNum()
    Next
  End Function
  
  Public GetCountStartDate()
    DateSort
    GetCountStartDate=AddrNumLists(1).GetFirstDate
  End Function
  
   'その宛先からのメールの数を返す
  Public Function GetNum(Address,Name)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      GetNum=AddrNumLists(index).getNum()
      Exit Function
    End If
    
    GetNum=0
  End Function
  
  'その宛先のメールをどのように扱うか(削除済みへ移動（デフォルト)か、保存か,完全削除か)
  Public Function GetState(Address,Name)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      GetState=AddrNumLists(index).GetMailState()
      Exit Function
    End If
    GetState="削除済みへ移動"
  End Function
  
  
  '初めてファイルに書き込む際,表示順を時間順にする必要がある
  Public Function DateSort()  
    Dim i
    Dim j
    Dim tmp
    Dim head
    head=1
    
   For i=2 To UBound(AddrNumLists)   
      j=i
      'VBSでは短絡評価してくれないので,一番最初のインデックスより前のインデックスにアクセスしてしまいエラーが出るので
      'ここでは,一番最初のインデックスの1つ前までのインデックスについてソートし
      Do While (j > head+1) And(AddrNumLists(j-1).GetFirstDate() > AddrNumLists(j).GetFirstDate())
        Set tmp=AddrNumLists(j)
        Set AddrNumLists(j)=AddrNumLists(j-1)
        Set AddrNumLists(j-1)=tmp
        j=j-1
      Loop
      
      '最後のインデックスだけここで別途行う
      If (AddrNumLists(head).GetFirstDate() > AddrNumLists(head+1).GetFirstDate()) Then
        Set tmp=AddrNumLists(head)
        Set AddrNumLists(head)=AddrNumLists(head+1)
        Set AddrNumLists(head+1)=tmp
      End If
      
   Next
        
  End Function
  
 
  
  Public Function ToFileWriteStr()
   
   
    Dim Header
    Dim Today
    Dim FWriter
    Today=Date()
    Dim FirstDate
    FirstDate=GetCountStartDate
    
    Dim Header
    Header="メールアドレス,宛名,"&FirstDate&"以降で最も早くその宛先からメールが届いた日付,"&FirstDate&"から"&Today&"までに届いたメールの数,メールの取り扱い"
    
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
    Footer="合計,,,"&MailSum&","
    Content(UBound(Content))=Footer
    
    Set ToFileWriteStr=Content
    
  End Function

End Class
    
    
   
Class FileOperator

  Public FIOOperator
  
  Public Sub  Class_Initialize()
   Set FIOOperator=Wscript.CreateObject("ADODB.Stream")
   FIOOperator.Type=2
   FIOOperator.Charset="UTF-8"
   FIOOperator.LineSeparator=10
  End Sub
  
  Public Sub Class_Terminate()
   Set FIOOperator=Nothing
  End Sub
  
  Public Function FRead(FilePath)
  
   FIOOperator.Open
   
   Dim FileOpen
   FileOpen=True
   
   'ファイルからデータを取得する際に,そのファイル自体が開かれていた時はエラーが出るので,Fileが閉じられるまで永久ループする
   Do While FileOpen
     On Error Resume Next
      FIOOperator.LoadFromFile FilePath
      
      'エラーがなかった（ファイルが閉じられていたら),ここにたどり着くので,FileOpenフラグを卸して,ループを抜ける
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
  
  End Function 
  
  Public Function FWrite(FilePath,Contents,Mode,WriteType)
    Dim FileOpen
    
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
    
    Do While FileOpen
      On Error Resume Next
        FIOOperator.SaveToFile FilePath,2
        If Err.Number = 0 Then
          FileOpen=False
        End If
      On Error GoTo 0 
    Loop
    
    FIOOperator.Close
    
  End Function

End Class

Class DateManager
  Public FirstMailDate
  Public LastMailDate
  Public ExeDate
  

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
   'ヘッダはいらないので1番目から取得する。そして最後の行は合計なのでそれもいらない
   For i=1 To UBound(FileContents)-1
     AllDataWithoutBr=Replace(FileContents(i),VbCr,"")
     OneData=Split(AllDataWithoutBr,",")
     GetDataManageObj.ParseDataFromFileContent OneData
   Next
   
  End Function
  
  Public Function GetLastMailDate()
   'ログファイルから最後にメールをチェックした日付情報を得る
   If FSObj.FileExists(TimeLogFile) Then
    ChangeFileAttributes(TimeLogFile)
    Dim Contents
    Contents=FIO.FRead(TimeLogFile)
    ChangeFileAttributes(TimeLogFile)
    GetLastMailDate=CDate(Replace(Contents(0),VbCr,""))
   ElseIf FSObj.FileExists(OriginalFileName) Then
    'Logファイルがなかった場合この日付で代用する
    GetLastMailDate=FSObj.GetFile(OriginalFileName).DateLastModified
   ElseIf FSObj.FileExists(BackupFileName) Then
    GetLastMailDate=FSObj.GetFile(BackupFileName).DateLastModified
   Else
    GetLastMailDate=CDate("1970/1/1")
   End If
   
  End Function
  
  Public Function WriteResultDataManageObj(DataManageObj)
   FIO.FWrite OriginalFileName,DataManageObj.ToFileWriteStr(),"w","Array"
   Copy OriginalFileName,BackUpFileName,True
  End Function
  
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
    Header="プログラム実行時刻,その時点での最新のメール時刻,"&FirstDate&"からその時点までの累積メール数"
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
  
  'メールをカウントする(カウントが0、つまりまだその宛先からのメールが存在しない場合は,新しくオブジェクトを作る)
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
   NumObj.SetValue addr,name,"1","削除済みへ移動"
   NumObj.AddDataPool date
   Set AddrNumLists(UBound(AddrNumLists))=NumObj
   
  End Function
  
  'その宛先からのメールの数を返す
  Public Function getNum(addr,name)
   For i=LBound(AddrNumLists) To UBound(AddrNumLists)
     If addr = AddrNumLists(i).GetAddress And AddrNumLists(i).getName = name Then
       getNum=AddrNumLists(i).getNum()
       Exit Function
     End If
   Next
   getNum=0
  End Function
  
  'その宛先のメールをどのように扱うか(削除済みへ移動（デフォルト)か、保存か,完全削除か)
  Public Function GetState(addr,name)
    Dim i
    For i=1 To UBound(AddrNumLists)
     If addr = AddrNumLists(i).GetAddress And name = AddrNumLists(i).GetName Then
       GetState=AddrNumLists(i).GetMailState()
       Exit Function
     End If
   Next
   GetState="削除済みへ移動"
  End Function
  
  'メールを調べるにあたり,保存メールの重複カウントを避けるため,前回、メールの数をカウントしたのはいつなのかを得る
  '保存メールに入っている,この日付より前のメールに関してはすでにカウントしているのでカウントしない
  Public Function GetModifiedFileDate()
    GetModifiedFileDate=FDate
  End Function
 
  Public Function SetFileDate()
   Dim FObj
   Dim FSObj
   Set FSObj=CreateObject("Scripting.FileSystemObject")
   'ログファイルから最後にメールをチェックした日付情報を得る
   If FSObj.FileExists(LogFile) Then
     FSObj.GetFile(LogFile).attributes=0
     FDate=GetRealModifiedDate
   ElseIf FSObj.FileExists(FileName) Then
     'Logファイルがなかった場合この日付で代用する
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
  
   'ファイルからデータを取得する際に,そのファイル自体が開かれていた時はエラーが出るので,Fileが閉じられるまで永久ループする
   Do While FileOpen
     On Error Resume Next
      FReader.LoadFromFile FName
      
      'エラーがなかった（ファイルが閉じられていたら),ここにたどり着くので,FileOpenフラグを卸して,ループを抜ける
      If Err.Number = 0 Then
        FileOpen=False
      End If
     On Error GoTo 0 
     
   Loop
       
     
   '1行目はヘッダなので情報として必要ない。
   'とりあえず,1回全部の行についてデータを取得する
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
     
   'ヘッダはいらないので1番目から取得する。そして最後の行は合計なのでそれもいらない
   For i=1 To UBound(AllData)-1
     AllDataWithoutBr=Replace(AllData(i),VbCr,"")
     OneData=Split(AllDataWithoutBr,",")
     ReDim Preserve AddrNumLists(UBound(AddrNumLists)+1)
     Set NumObj=new AddrNumSet
     If UBound(OneData) = 4 Then
       NumObj.SetValue OneData(0),oneData(1),OneData(3),OneData(4)
     ElseIf UBound(OneData) >= 3 Then
       NumObj.SetValue OneData(0),oneData(1),OneData(3),"削除済みへ移動"
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
  
  '初めてファイルに書き込む際,表示順を時間順にする必要がある
  Public Function DateSort()  
    Dim i
    Dim j
    Dim tmp
    Dim head
    head=1
    
    
   For i=2 To UBound(AddrNumLists)
      
      j=i
      'VBSでは短絡評価してくれないので,一番最初のインデックスより前のインデックスにアクセスしてしまいエラーが出るので
      'ここでは,一番最初のインデックスの1つ前までのインデックスについてソートし
      Do While (j > head+1) And(AddrNumLists(j-1).GetFirstDate() > AddrNumLists(j).GetFirstDate())
        Set tmp=AddrNumLists(j)
        Set AddrNumLists(j)=AddrNumLists(j-1)
        Set AddrNumLists(j-1)=tmp
        j=j-1
      Loop
      
      '最後のインデックスだけここで別途行う
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
    
    Header="メールアドレス,宛名,"&FirstDate&"以降で最も早くその宛先からメールが届いた日付,"&FirstDate&"から"&Today&"までに届いたメールの数,メールの取り扱い"
    Set FWriter=Wscript.CreateObject("ADODB.Stream")
    FWriter.Type=2
    FWriter.Charset="UTF-8"
    FWriter.Open
    FWriter.WriteText Header,1
    
    Dim i
    
    '合計メール数(1番最後の行に書いておく)
    Dim MailSum
    MailSum=GetSumMailNum()
    
    For i= 1 To UBound(AddrNumLists)
      FWriter.WriteText AddrNumLists(i).ToStr(),1
   Next
   
    
    FWriter.WriteText "合計,,,"&MailSum&",",1
    
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
    'バックアップファイルがない場合エラーが出る
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
  
  'チェックした日付をログに記録
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
      FWriter.WriteText "プログラム実行時刻,その時点での最新のメール時刻,"&FirstDate&"からその時点までの累積メール数",1
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

  'outlookにアクセス
  Set outlook=Wscript.CreateObject("Outlook.Application")

  'outlookのメールにアクセスするためのインターフェース
  Set namespace=outlook.GetNameSpace("MAPI")
  

  'outlookの受信フォルダ
  Set receiveFolder=namespace.GetDefaultFolder(6)
  
  '削除済みフォルダ
  Dim deletedFolder
  Set deletedFolder=namespace.GetDefaultFolder(3)
  
  
  
  'テスト用完全削除済みフォルダ―(テストの時はいきなり完全削除しない)
  Dim testCompDeleteFolder
  Set testCompDeleteFolder=deletedFolder.Folders("test_comp_delete")
  
  
  'outlookを最小化して起動(相手にoutlookが起動していることがわからないようにする)
  Set Wsh=Wscript.CreateObject("Wscript.shell")
  Wsh.Run "outlook.exe",7,False


  'メールのカウントなどを行う管理クラス
  Dim CountManager

  Set CountManager=new Manager
  
  
  'ファイル読込
  CountManager.FRead()
  
 
  
  'メールの扱い方(メールの宛先によってメールの処理の仕方を変える)
  '削除済みへ移動（デフォルト)か,完全削除か,保存か
  Dim MailOperation
 
  'メール
  Dim OneMailItem
   
  'もしもとっておきたいメールが出たらこの変数を用いてとっておきたいメールの数を保存しておく
  'つまりメールを操作する際のインデックス
  Dim SaveMailNum
  SaveMailNum=0
  
  Dim CurrentMailNum
  
  Dim TestDeleteMailNum
 
 
   
  'まずは削除済みアイテムについてみてゆく
  Do While SaveMailNum < deletedFolder.Items.Count
     CurrentMailNum=deletedFolder.Items.Count
     Set OneMailItem=deletedFolder.Items.Item(SaveMailNum+1)
     MailOperation=CountManager.GetState(OneMailItem.SenderEMailAddress,OneMailItem.SenderName)
     Select Case MailOperation
       Case "保存"
        OneMailItem.Move receiveFolder
        Do While  CurrentMailNum = deletedFolder.Items.Count 
            Wscript.Sleep 1000
        Loop
       Case "完全削除"
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
   
  
  '削除済みアイテムの方を見たので次は受信済みの方を見る
  Set OneMailItem=Nothing
  '保存（フォルダキープ）のメール数を0に戻す
  SaveMailNum=0
  MailOperation=""
 
  '次は受信アイテムの中を見てゆくが,受信アイテムの読み込みとこの処理は非同期であることから少しタイムラグを設ける必要がある
  'タイムラグ用の変数
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
 
  '最後にいつメールのカウントを行ったのかを得る(保存フォルダから重複カウントをしないように)
  Dim LastCountMailDate
  LastCountMailDate=CountManager.GetModifiedFileDate()
  
  
  'カウンタ変数
  Dim i
  
  'メールの情報
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
      'メールを消去(削除済みフォルダ）に移動させると,受信トレーのメールが減るので,ずっと,同じインデックスをアクセスする
      EnterFlag=True
      
      On Error Resume Next
       Set OneMailItem=receiveFolder.Items.Item(SaveMailNum+1)
       HasError=Err.Number
       If HasError <> 0 Then
         Wscript.Echo HasError
         Wscript.Echo "エラー"
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
            CountManager.Count Addr,Name,Time
          End If
            
          If LastMailTime < Time Then
            LastMailTime=Time
          End If
            
          CurrentMailNum=receiveFolder.Items.Count
          CurrentDeletedFolderMailNum=deletedFolder.Items.Count
          TestFolderMailNum=testCompDeleteFolder.Items.Count
          MailOperation=CountManager.GetState(Addr,Name)
              
          'メールの扱い方（宛先によってメールをどう扱うか)
           
          '保存する場合
          Select  Case MailOperation
            Case  "保存"
             'このメールは消さずに参照するメールのインデックスを1進める
             SaveMailNum=SaveMailNum+1
              
            Case"完全削除"
             'On Error Resume Next
             'OneMailItem.Delete
             'On Error GoTo 0
             OneMailItem.Move testCompDeleteFolder
             '削除が完了するまで待つ
             Do While  CurrentMailNum = receiveFolder.Items.Count  And TestFolderMailNum = testCompDeleteFolder.Items.Count
               Wscript.Sleep 1000
             Loop
             
            Case Else
              CurrentDeletedFolderMailNum=deletedFolder.Items.Count
              OneMailItem.Move deletedFolder
              '移動が完了するまで待つ
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
  
  'Explorerが取得できなかったら,強制終了
  If Explorer Is Nothing Then
    'Outlookの終了
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