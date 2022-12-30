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
    '管理配列（この中に各メールアドレス（宛名）からのメールデータセットを格納）
    ReDim AddrNumLists(0)
   End Sub
   
   Public Sub Class_Terminate()
    Dim i
    For i=LBound(AddrNumLists) To UBound(AddrNumLists)
      Set AddrNumLists(i)=Nothing
    Next
   End Sub
   
   'ファイルの内容からデータセット(メールアドレス、宛名、最初の日時、その日時以降のその宛先からのメール数、取り扱い)を作り出す
   Public Function ParseDataFromFileContent(OneData)
    ReDim Preserve AddrNumLists(UBound(AddrNumLists)+1)
    Dim NewObj
    Set NewObj=new AddrNumSet
    If UBound(OneData) = 4 Then
       Select Case OneData(4)
        Case "保存","削除済みへ移動","完全削除"
          NewObj.SetValue OneData(0),oneData(1),OneData(3),OneData(4)
        Case Else
          NewObj.SetValue OneData(0),oneData(1),OneData(3),"削除済みへ移動"
       End Select
    ElseIf UBound(OneData) >= 3 Then
       NewObj.SetValue OneData(0),OneData(1),OneData(3),"削除済みへ移動"
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
    
   'メールアドレスと宛名をキーとして,そのメールアドレス（宛名）の情報が管理配列のどこのインデックスにあるかを示す
   Public Function DataIndex(Address,Name)
    Dim i
    For i=1 To UBound(AddrNumLists)
      If AddrNumLists(i).GetAddress = Address And AddrNumLists(i).GetName = Name Then
        DataIndex=i
        Exit Function
      End If
    Next
    DataIndex=-1
   End Function
   
   'こちらは与えられたメールアドレスと宛名からのメールが存在するかどうかを返す
   'つまり,上のメソッドの戻り値が-1の時はその宛先からのメールはない.それ以外の時はあるということ
   Public Function DataExists(Address,Name)
     If DataIndex(Address,Name) <> -1 Then
       DataExists=True
       Exit Function
     End If
     DataExists=False
   End Function
   
   'メールをカウントする(カウントが0、つまりまだその宛先からのメールが存在しない場合は,新しくオブジェクトを作る)
   Public Function Count(Address,Name,MailDate)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      AddrNumLists(index).NumIncrement()
      AddrNumLists(index).AddDate MailDate
      Exit Function
    End If
    
    ReDim Preserve AddrNumLists(UBound(AddrNumLists)+1)
    Dim NewObj
    Set NewObj=new AddrNumSet
    NewObj.SetValue Address,Name,"1","削除済みへ移動"
    NewObj.AddDate MailDate
    Set AddrNumLists(UBound(AddrNumLists))=NewObj
   End Function
   
   '現時点での合計メール数
   Public Function GetSumMailNum()
    GetSumMailNum=0
    Dim i
    For i= 1 To UBound(AddrNumLists)
      GetSumMailNum=GetSumMailNum+AddrNumLists(i).GetNum()
    Next
  End Function
  
 
  '与えられたメールアドレス（宛名）からのメールの数を返す
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
  
  'メールの集計を開始した日の日付を取得
  Public Function GetCountStartDate()
    DateSort
    GetCountStartDate=AddrNumLists(1).GetFirstDate
  End Function
  
  
  '実際にファイルに書き込むときの文字列が要素となった配列を生成する
  Public Function ToFileWriteStr()
   
    DateSort()
    Dim Header
    Dim Today
    Dim FWriter
    Today=Date()
    Dim FirstDate
    FirstDate=GetCountStartDate
    
    Header="メールアドレス,宛名,"&FirstDate&"以降で最も早くその宛先からメールが届いた日付,"&FirstDate&"から"&Today&"までに届いたメールの数,メールの取り扱い"
    
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
    Dim MailSum
    MailSum=GetSumMailNum()
    Footer="合計,,,"&MailSum&","
    Content(UBound(Content))=Footer
    
    ToFileWriteStr=Content
    
  End Function

End Class
    
    

'ファイルの読み書きをする  
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
   
   'ファイルからデータを取得する際に,そのファイル自体が開かれていた時はエラーが出るので,Fileが閉じられるまで永久ループする
   Do While FileOpen
     On Error Resume Next
      FIOOperator.LoadFromFile FilePath
      
      'エラーがなかった（ファイルが閉じられていたら),ここにたどり着くので,FileOpenフラグを卸して,ループを抜ける
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
  
  'ファイルの書き込み(Modeは"w"なら上書き,"a"なら追記,WriteTypeは"Array"なら配列として,各要素を1行ずつ書いてゆく,"Str"なら文字列として1行だけ書く)
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
    
    'エラーチェック(書き込もうとしているファイルが開かれていた場合,エラーが出てしまうので,エラーがなくなる（ファイルが閉じられるまで）待つ)
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
   'ヘッダはいらないので1番目から取得する。そして最後の行は合計なのでそれもいらない
   For i=1 To UBound(FileContents)-1
     AllDataWithoutBr=Replace(FileContents(i),VbCr,"")
     OneData=Split(AllDataWithoutBr,",")
     GetDataManageObj.ParseDataFromFileContent OneData
   Next
   
  End Function
  
  'メールを調べるにあたり,保存メールの重複カウントを避けるため,前回、メールの数をカウントしたのはいつなのかを得る
  '保存メールに入っている,この日付より前のメールに関してはすでにカウントしているのでカウントしない
  Public Function GetLastMailDate()
   'ログファイルから最後にメールをチェックした日付情報を得る
   If FSObj.FileExists(TimeLogFile) Then
    FSObj.GetFile(TimeLogFile).attributes=0
    Dim Contents
    Contents=FIO.FRead(TimeLogFile)
    FSObj.GetFile(TimeLogFile).attributes=1
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
  
  '実際に結果の書き込み
  Public Function WriteResultDataManageObj(DataManageObj)
   FIO.FWrite OriginalFileName,DataManageObj.ToFileWriteStr(),"w","Array"
   Dim FObj
   'ファイルがないとき（はじめてバックアップファイルを作成する際,ファイル自体が存在しないので,属性を変えるにも変えられないため
   'その時はエラーを握りつぶす
   On Error Resume Next
     FSObj.GetFile(BackupFileName).attributes=0
   On Error GoTo 0
   FSObj.CopyFile OriginalFileName,BackUpFileName
   Set FObj=FSObj.GetFile(BackupFileName)
   FObj.attributes=1
  End Function
  
  '次回メールを調べてカウントするにあたり,今回何時何分のメールまでがカウント済みなのかを記録しておく
  '上記のように次回、どのメールからカウントすればよいのかを書くため（重複カウントを避けるため）
  Public Function WriteRenewLastMailDate(CountStartDate,LastMailDate,MailNum)
   On Error Resume Next
    FSObj.GetFile(TimeLogFile).attributes=0
   On Error GoTo 0
   FIO.FWrite TimeLogFile,""&LastMailDate,"w","Str"
   FSObj.GetFile(TimeLogFile).attributes=1
   
   Dim NowDate
   NowDate=Now
   
   If FSObj.FileExists(NumLogFile) Then
     Dim Content
     Content=""&NowDate&","&LastMailDate&","&MailNum
     FSObj.GetFile(NumLogFile).attributes=0
     FIO.FWrite NumLogFile,Content,"a","Str"
     FSObj.GetFile(NumLogFile).attributes=1
   Else
    Dim Header
    Header="プログラム実行時刻,その時点での最新のメール時刻,"&CountStartDate&"からその時点までの累積メール数"
    Dim Body
    Body=""&NowDate&","&LastMailDate&","&MailNum
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
  
  Public Function PadZero(Before,Num)
    Dim BeforeInt
    BeforeInt=CLng(Before)
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
     
       
End Class
  
Function MailEquals(One,Other)
 
 On Error Resume Next
  If One.SenderEMailAddress <> Other.SenderEMailAddress Then
    MailEquals=False
    Exit Function
  End If
  
  If Err.Number <> 0 Then
    Err.Clear
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

'第一引数として指定したメールオブジェクトが第二引数のフォルダの何番目にあるかを返す関数
'存在すればそのインデックス,存在しなかった場合,-1を返し,第一引数が空であれば-2を返す
Function MailIndexInFolder(MailItem,Folder)
  Dim i
  Dim Result
  For i=1 To Folder.Items.Count
   Result=MailEquals(MailItem,Folder.Items(i))
   If Result Then
     MailIndexInFolder=i
     Exit Function
   ElseIf IsEmpty(Result) Then
     MailIndexInFolder=-2
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
  
  'ファイルの読み書きに関する管理をするクラス
  Dim FManager
  Set FManager=new FileManager
  
  'メールの数や取り扱いを管理するクラス
  Dim DManager
  Set DManager=FManager.GetDataManageObj()
  
 
  
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
     MailOperation=DManager.GetState(OneMailItem.SenderEMailAddress,OneMailItem.SenderName)
     Select Case MailOperation
       Case "保存"
        OneMailItem.Move receiveFolder
        Do While  CurrentMailNum = deletedFolder.Items.Count 
            Wscript.Sleep 500
        Loop
       Case "完全削除"
        'OneMailItem.Delete
        TestDeleteMailNum=testCompDeleteFolder.Items.Count
        OneMailItem.Move testCompDeleteFolder
        'Do While  CurrentMailNum = deletedFolder.Items.Count 
        Do While TestDeleteMailNum = testCompDeleteFolder.Items.Count
            Wscript.Sleep 500
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
  NormSec=120
  
  Dim AllMailNum
  AllMailNum=receiveFolder.Items.Count+deletedFolder.Items.Count
  
  'メールの件数が多い場合,処理そのものに時間がかかるため,タイムラグの秒数は減らす
  'これを減らさないとユーザーを待たせる時間が増える
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
 
  '最後にいつメールのカウントを行ったのかを得る(保存フォルダから重複カウントをしないように)
  Dim LastCountMailDate
  LastCountMailDate=FManager.GetLastMailDate()
  
  
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
       Err.Clear
      On Error GoTo 0
      
      If HasError = 0 Then
      
        Name=OneMailItem.SenderName
        Addr=OneMailItem.SenderEmailAddress
        Time=OneMailItem.ReceivedTime
        
        Dim  MailIndex
        MailIndex=MailIndexInFolder(OneMailItem,deletedFolder)
          
        If MailIndex  > -1 Then
         On Error Resume Next
          deletedFolder.Items.Item(MailIndex).Delete
          Err.Clear
         On Error GoTo 0
        End If
        
        If MailIndex <> -2 Then 
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
              
          'メールの扱い方（宛先によってメールをどう扱うか)
           
          '保存する場合
          Select  Case MailOperation
            Case  "保存"
             'このメールは消さずに参照するメールのインデックスを1進める
             SaveMailNum=SaveMailNum+1
              
            Case"完全削除"
             OneMailItem.Delete
             '移動が完了するまで待つ
             Do While  CurrentMailNum = receiveFolder.Items.Count And  CurrentDeletedFolderMailNum = deletedFolder.Items.Count 
               Wscript.Sleep 500
             Loop
             Dim DeletedMailIndex
             DeletedMailIndex=MailIndexInFolder(OneMailItem,deletedFolder)
             'On Error Resume Next
             'deletedFolder.Items.Item(DeletedMailIndex).Delete
             'Err.Clear
             'On Error GoTo 0
             
             'とりあえず,テスト段階では完全削除フォルダーに移動する
             deletedFolder.Items.Item(DeletedMailIndex).Move testCompDeleteFolder
             '削除が完了するまで待つ
             Do While  CurrentMailNum = receiveFolder.Items.Count  And TestFolderMailNum = testCompDeleteFolder.Items.Count
               Wscript.Sleep 500
             Loop
             
             
              'On Error Resume Next
              'deletedFolder.Items.Item(MailIndex).Delete
              'Err.Clear
              'On Error GoTo 0
             'On Error GoTo 0
             OneMailItem.Delete
             
             OneMailItem.Move testCompDeleteFolder
             '削除が完了するまで待つ
             Do While  CurrentMailNum = receiveFolder.Items.Count  And TestFolderMailNum = testCompDeleteFolder.Items.Count
               Wscript.Sleep 500
             Loop
             
            Case Else
              CurrentDeletedFolderMailNum=deletedFolder.Items.Count
              OneMailItem.Move deletedFolder
              '移動が完了するまで待つ
              Do While  CurrentMailNum = receiveFolder.Items.Count And  CurrentDeletedFolderMailNum = deletedFolder.Items.Count 
                Wscript.Sleep 500
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
      If CurrentNormSec < 1 Then
        CurrentNormSec=1
      End If
    End If
      
    Wscript.Sleep 500
    CountWaitSec=CountWaitSec+1
    
  Loop
  

  FManager.WriteResultDataManageObj DManager
  FManager.WriteRenewLastMailDate DManager.GetCountStartDate(),LastMailTime,DManager.GetSumMailNum()
   
 
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