Option Explicit


Function EscapeStrForCSV(ByVal CSVStr)
  'CSVではダブルクオーテーションは「カンマを無視する」という合図であるので
  'これを、「ダブルクオーテーションそのものである」ことを示すには,ダブルクオーテーション2つ分必要
  EscapeStrForCSV=Replace(CSVStr,"""","""""")
  If InStr(EscapeStrForCSV,",") > 0 Then
    'また文字列に「,(カンマ)」が含まれていたら,それは,「文字列のカンマ」であるということを示すために逆にその文字列をダブルクオーテーションでくくる
    EscapeStrForCSV=""""&EscapeStrForCSV&""""
  End If
End Function

'CSV用のエスケープを解除
Function DeEscapeStrForCSV(ByVal CSVStr)
  'ダブルクオーテーションが二つあるときは,1つはダブルクオーテーションそのものでもう1つはダブルクオーテーションに文字列であると命令するもの
  DeEscapeStrForCSV=Replace(CSVStr,"""""","""")
  If InStr(DeEscapeStrForCSV,",") > 0 Then
     DeEscapeStrForCSV=Left(DeEscapeStrForCSV,Len(EscapeStrForCSV)-Len(""""))
    DeEscapeStrForCSV=Right(DeEscapeStrForCSV,Len(EscapeStrForCSV)-Len(""""))
  End If
End Function

'第一引数の文字列を,第二引数の文字列を区切り文字として配列に分割するが(ここまでは標準のsplit関数と同じ)
'第三引数の文字列に挟まれたときは第二引数の文字列が来ても区切り文字とみなさないようにしたい
'多言語でいう不定長の後読み先読みだが、VBSにはそんな機能はないので自前で実装
'例えば,CSVでは「,」をデータ間の区切り文字(セル)としているが、「"(ダブルクオーテーション)」の中に挟まれたカンマは区切り文字として扱いたくない。そんな時に使う
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
    State="削除済みへ移動"
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
  
  '新しくカウント
  '第一引数は以前にそのメールを数えたことがあるかどうか(このメールが前回カウント時のメールの最新の日時より後であること)を示すフラグ→累積メール数を数える際に必要(現存メール数には必要ない)
  '第二引数は,現在このメールがどこにあるかを示す（メールのフォルダの場所)(完全削除されたものに関してはカウントしない)
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
    Set NewObj=New AddrNumSet
    If UBound(OneData) = 7 Then
       Select Case OneData(7)
        Case "保存","削除済みへ移動","完全削除"
          NewObj.SetValue DeEscapeStrForCSV(OneData(0)),DeEscapeStrForCSV(oneData(1)),OneData(3),OneData(7)
        Case Else
          NewObj.SetValue DeEscapeStrForCSV(OneData(0)),DeEscapeStrForCSV(oneData(1)),OneData(3),"削除済みへ移動"
       End Select
    ElseIf UBound(OneData) >= 3 Then
       NewObj.SetValue DeEscapeStrForCSV(OneData(0)),DeEscapeStrForCSV(OneData(1)),OneData(3),"削除済みへ移動"
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
   
   'こちらは与えられたメールアドレスと宛名からのメールが存在するかどうかを返す
   'つまり,上のメソッドの戻り値が-1の時はその宛先からのメールはない.それ以外の時はあるということ
   Public Function DataExists(ByVal Address,ByVal Name)
     If DataIndex(Address,Name) <> -1 Then
       DataExists=True
       Exit Function
     End If
     DataExists=False
   End Function
   
   'メールをカウントする(カウントが0、つまりまだその宛先からのメールが存在しない場合は,新しくオブジェクトを作る)
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
      NewObj.SetValue Address,Name,"1","削除済みへ移動"
      NewObj.AddDate MailDate
      NewObj.NumIncrement False,CurrentFolderName
      Set AddrNumLists(UBound(AddrNumLists))=NewObj
    End If
   End Function
   
   '現時点での合計メール数
   Public Function GetSumMailNum(ByVal ItemName)
    GetSumMailNum=0
    Dim i
    For i= 1 To UBound(AddrNumLists)
      GetSumMailNum=GetSumMailNum+AddrNumLists(i).GetMailNum(ItemName)
    Next
  End Function
  
 
  '与えられたメールアドレス（宛名）からのメールの数を返す
  Public Function GetNum(ByVal Address,ByVal Name,ByVal ItemName)
    Dim index
    index=DataIndex(Address,Name)
    If index <> -1 Then
      GetNum=AddrNumLists(index).GetMailNum(ItemName)
      Exit Function
    End If
    
    GetNum=0
  End Function
  
  'その宛先のメールをどのように扱うか(削除済みへ移動（デフォルト)か、保存か,完全削除か)
  Public Function GetState(ByVal Address,ByVal Name)
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
    Dim HeaderFormer
    Dim Header
    Dim Today
    Dim FWriter
    Today=Date()
    Dim FirstDate
    FirstDate=GetCountStartDate
    
    HeaderFormer="メールアドレス,宛名,"&FirstDate&"以降で最も早くその宛先からメールが届いた日付,"&FirstDate&"から"&Today&"までに届いた累積メール数(完全削除されたものも含む),"
    Header=HeaderFormer&Today&"時点で存在しているメール数(除完全削除・含削除済みフォルダ),"&Today&"時点での受信フォルダのメール数,"&Today&"時点の削除済みフォルダのメール数,メールの取り扱い"
    
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
    Footer="合計,,,"&CumulativeMailSum&","&ExistsMailSum&","&ReceivedMailSum&","&DeletedMailSum
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
   'バックアップフォルダがないなら作る
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
   'ヘッダはいらないので1番目から取得する。そして最後の行は合計なのでそれもいらない
   For i=1 To UBound(FileContents)-1
     AllDataWithoutBr=Replace(FileContents(i),VbCr,"")
     OneData=SplitExceptEscapeChar(AllDataWithoutBr,",","""")
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
     Err.Clear
   On Error GoTo 0
   FSObj.CopyFile OriginalFileName,BackUpFileName
   Set FObj=FSObj.GetFile(BackupFileName)
   FObj.attributes=1
  End Function
  
  '次回メールを調べてカウントするにあたり,今回何時何分のメールまでがカウント済みなのかを記録しておく
  '上記のように次回、どのメールからカウントすればよいのかを書くため（重複カウントを避けるため）
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
    Header="プログラム実行時刻,その時点での最新のメール時刻,"&CountStartDate&"からその時点までの累積メール数(含完全削除),その時点で存在しているメール数(除完全削除・含削除済みフォルダ),受信済みフォルダのメール数,削除済みフォルダのメール数"
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

'メールアイテムを完全削除する際、受信済みフォルダから一気に削除することができない
'ゆえに完全削除の際は,「受信済みからいったん削除済みに移動してから、そこから削除」,つまり２回の移動手続きをしなければならないが
'完全削除する際,その完全削除したいと思っているメールが削除済みフォルダのどこにあるのかを送信者情報や送信時刻などをヒントに探す必要がある
'にもかかわらず受信済みから削除済みに移動したとき(1回目の移動の時）に、完全削除しようと思っているメールの送信者情報や送信時刻などが参照できなくなってしまう.
'ゆえにここでは,受信済みからの移動の前に,送信者情報等を残しておく一時変数的クラスをここで「わざわざ」作っておく

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
  
  'バグの温床となるため,フィールドに直接アクセスせず,プロパティを通すようにする
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



'メールアイテムが同じものなのかを判定する(疑似(一時保存したもの:詳細は後述)でも本物のメールアイテムオブジェクトでも両方とも同じように判定する)
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
    
    '誤って数秒ずれて同一の内容のメールが送られてしまうということがある可能性を考えて
    '時刻に関しては全くきっかり同じでなくても5秒未満の誤差なら同じモノ扱いする
    If 5 < Abs(DateDiff("s",One.ReceivedTime,Other.ReceivedTime)) Then
      MailItemEquals=False
      Exit Function
    End If
     
    
    MailItemEquals=True
    
End Function
 

'本来のReceivedFolderとDeletedFolderは,時間順に並んでなくて重複メールの削除等をする際,いちいち1つずつ比べなければいけない
'なのでここでは削除済みフォルダと受信フォルダのアイテムをコピーしたもの(上記で述べた疑似メール情報)をおいておく)
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
  
  'これはただ探索するだけのメソッド
  '見つかれば,このクラスの管理配列におけるその要素のインデックスを返し
  '見つからなかった場合は新しく挿入予定の負のインデックスを入れる
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
     'アイテムが1つもないとき
     If UBound(TimeOrderedMailItem) = 0 Then
       Find=-1
       Exit Function
     '１番最初のアイテムより時間が早い場合
     ElseIf OneMailInfo.ReceivedTime < TimeOrderedMailItem(1).ReceivedTime Then
       Find=-1
       Exit Function
     '1番最後のアイテムより時間が遅い場合→この3つの条件のうちどれか1つに当てはまるときは,探索しなくてよい
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
  
  'こちらは挿入(見つからなかった場合にのみ行い,Searchメソッドが返した負のインデックスを正に戻して、その周辺に入れる
  Public Function AddData(OneMailInfo)
    
    '要素が一つもないときは調べずに1番目に挿入
    If UBound(TimeOrderedMailItem) = 0 Then
       Insert 1,OneMailInfo
       Exit Function
    '1番目の要素より早い時間のものは1番目に挿入
    ElseIf OneMailInfo.ReceivedTime < TimeOrderedMailItem(1).ReceivedTime Then
       Insert 1,OneMailInfo
       Exit Function
    '1番最後の要素より時間が遅いものは一番最後に挿入
    ElseIf TimeOrderedMailItem(UBound(TimeOrderedMailItem)).ReceivedTime < OneMailInfo.ReceivedTime Then
       Insert UBound(TimeOrderedMailItem)+1,OneMailInfo
       Exit Function
    End If
    
    '上記のような場合は調べなくても挿入する場所は自明であったが,それ以外の場合は挿入する場所を調べなくてはならない
    Dim InsertIndex
    InsertIndex=Search(OneMailInfo)
    'すでに存在した場合は挿入してはいけない
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
     Wscript.Echo i&"番目 メールアドレス:"&TimeOrderedMailItem(i).SenderEmailAddress&" 宛名:"&TimeOrderedMailItem(i).SenderName&" 時刻:"&TimeOrderedMailItem(i).ReceivedTime&" 件名"&TimeOrderedMailItem(i).Subject
    Next
  End Function
  
  Public Function  Item(index)
    Items=TimeOrderedMailItem(index)
  End Function
  
  Public Property Get Count()
    Count=UBound(TimeOrderedMailItem)
  End Property
    
End Class

'ここでは削除済みフォルダのメールの重複を削除するクラスを作る
'最初に「削除済みフォルダ」内の重複メールを削除してしまう。
'その後,「受信フォルダ」を一つずつ見るときに「受信フォルダ」と「削除済みフォルダ」にまたがっているもの,あるいは,「受信フォルダ」同士の重複を削除する
'実際は受信フォルダと削除済みアイテムは別々になっているが、このTimeOrderedReceivedAndDeletedFolderには受信フォルダと削除済みフォルダの両方合わせたメールアイテム(上で述べた疑似)をここでは格納し,重複を排除する
'これで「受信フォルダ」のみの重複削除,「削除済みフォルダ」のみの重複削除にとどまらず,両フォルダにまたがって重複するアイテムも一括で削除できる
Class TimeOrderedRemoveDuplicateInFolder
  Public TimeOrderedReceivedAndDeletedFolder
  
  Public Sub Class_Initialize()
    Set TimeOrderedReceivedAndDeletedFolder=New TimeOrderedFolder
  End Sub
  
  Public Sub Class_Terminate()
    Set TimeOrderedReceivedAndDeletedFolder=Nothing
  End Sub

  'ここでは最初に「削除済みフォルダ」に入っているもののみを削除
  Public Function  RemoveDuplicateInDeletedFolder(DeletedFolder)
   Dim MailItemIndex
   MailItemIndex=1
  
   'テスト用完全削除済みフォルダ―(テストの時はいきなり完全削除しない)
   Dim testCompDeleteFolder
   Set testCompDeleteFolder=DeletedFolder.Folders("test_comp_delete")
  
   '各メールアイテムを調べる
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
        'インデックスの加算は見つからなかったときのみ
        MailItemIndex=MailItemIndex+1 
    Else
      '新設したものなのでとりあえずテスト完全削除フォルダーへ移動
      Dim TestFolderMailNum
      TestFolderMailNum=testCompDeleteFolder.Items.Count
      CurrentCheckingMailItem.Move testCompDeleteFolder
      Do While TestFolderMailNum = testCompDeleteFolder.Items.Count
        Wscript.Sleep 500
      Loop
        
      '以降はテスト終了後(今はコメントアウト)
      'Dim CurrentDeletedFolderMailNum
      'Dim CurrentFolderMailNum
      'CurrentFolderMailNum=Folder.Items.Count
      'Dim CurrenDeletedFolderMailNum
      'CurrentDeletedFolderMailNum=DeletedFolder.Items.Count
      'Do
        'On Error Resume Next
        'Err.Clear
        '「受信済み」にあるメールアイテムに対してDeleteメソッドを用いた場合,このままだと「削除済み」に移動だけで終わってしまう
        'しかし,2度目の関数呼び出し時の「削除済みフォルダ」を調べる際,当然ながらこのメール情報は重複していると判定される
        '（1度目の呼び出しの際に重複と判定されれば2度目の際はデータは減らないので必ず重複と判定される）ので、2度目の呼び出しの際に同じようにDeleteすれば完全削除されるので心配ない
        'CurrentCheckingMailItem.Delete
        'On Error GoTo 0
      'Loop While Err.Number <> 0
       
      '「削除メソッドを呼び出す前のメールアイテム数と現在のメールアイテム数」(削除済みフォルダのアイテム数も同様)の変化を見比べて,数が変わったら移動(削除)が完了したということになる
      '第一引数が削除済みフォルダだったときは全く同じ処理が2回になるので無駄な処理になるかもしれないが,同じ関数を2度定義したり受信済みだけ別の処理をするよりかははるかに効率が良い)
      'Do While CurrentFolderMailNum = Folder.Items.Count And  CurrentDeletedFolderMailNum = DeletedFolder.Items.Count 
        'Wscript.Sleep 500
      'Loop
      'コメントアウト終了
    End If
    Set CurrentCheckingMailItem=Nothing
    
  Loop
 End Function
 
 '「受信フォルダ」のメールアイテムを調べて削除する際,そのアイテムがすでに「(疑似)削除済みフォルダ」に存在しないかどうか調べる
 '引数として与えるのは「受信フォルダ」の調査中のメールアイテム(メールアイテムそのものは移せないので疑似のもの）
 '「この疑似削除済みフォルダ」に、今調べている受信フォルダの一メールアイテムが存在した場合は重複ありということなので,Trueを返す
 '「疑似削除済みフォルダ」に今調べているメールアイテムが存在しなければ,この疑似削除済みフォルダにメールアイテムを追加し,Falseを返す.
 'このことによって,もともと「受信済みフォルダ」にあったアイテム同氏の重複が調べられる
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
   

'一時保存したこのメール情報がフォルダのどのインデックスに存在するかを調べるメソッド(線形探索)
'存在する場合は存在するインデックス,存在しない場合は-1を返却する
'本物のメールアイテムオブジェクトでも,疑似的に作成した一時メール情報保存オブジェクトでもどちらでもよい
Function MailIndexInFolder(OneMailItem,ItemsFolder)
 Dim CompMailItem
 Dim i

 '挿入されたばかりのデータは後ろのほうに来やすい(?)ので後ろから調べたほうが早い
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

  'outlookにアクセス
  Set outlook=Wscript.CreateObject("Outlook.Application")

  'outlookのメールにアクセスするためのインターフェース
  Set namespace=outlook.GetNameSpace("MAPI")
  

  'outlookの受信フォルダ
  Set receiveFolder=namespace.GetDefaultFolder(6)
  
  '削除済みフォルダ
  Dim deletedFolder
  Set deletedFolder=namespace.GetDefaultFolder(3)
  
  
  
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
  
  
  '完全削除するメールを一時保存するもの
  Dim TmpOneMailItem
 
 'どの時刻までのメールをカウントしたのか（前回カウント実行時点での最新のメールの時刻)を得る(保存フォルダから重複カウントをしないように)
  Dim LastCountMailDate
  LastCountMailDate=FManager.GetLastMailDate()
  
  '本カウントの中で一番新しい時刻(次回分のLastCountMailDate)
  'この変数は削除済みフォルダを見た後に初期化しない
  Dim LastMailTime
  LastMailTime=LastCountMailDate
  
  'メール情報
  '上からメールアドレス、宛名、受信時刻,件名、本文
  Dim Address
  Dim Name
  Dim Time
  Dim Subject
  Dim Body
  
  '削除済みフォルダの重複削除
  '各フォルダの重複削除は数にもよるが,線形探索的に行う場合計算量がn^2となってしまうので非常に効率が悪い
  'よってメールの時刻が古いほうから順番に並ぶような順序付きの削除済みフォルダー(本物の削除済みフォルダーは必ずしもメールの時刻順に並んでいるわけではないため)を疑似的に定義して
  '二分探索で重複を発見し削除する
  
  '順序付き受信+削除済みフォルダ(両フォルダのメールアイテム(疑似)を混ぜて,時間順に並べ二分探索で高速で重複を発見し削除する)
  Dim TimeOrderedFolder
  Set TimeOrderedFolder=New TimeOrderedRemoveDuplicateInFolder
  
  TimeOrderedFolder.RemoveDuplicateInDeletedFolder deletedFolder
  
  
  'まずは削除済みアイテムについてみてゆく
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
       Case "保存"
        OneMailItem.Move receiveFolder
        Do While  CurrentMailNum = deletedFolder.Items.Count 
            Wscript.Sleep 500
        Loop
       Case "完全削除"
       
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
   
    
  
  
  
  
  '削除済みアイテムの方を見たので次は受信済みの方を見る
  Set OneMailItem=Nothing
  Set TmpOneMailItem=Nothing
  
  '保存（フォルダキープ）のメール数を0に戻す
  SaveMailNum=0
  
  'メールの扱い方も元に戻す
  MailOperation=""
  
  'メール情報の初期化
  Address=""
  Name=""
  Time=""
  Subject=""
  Body=""
 
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
 
  
  
  'カウンタ変数
  Dim i
 
 
  CurrentMailNum=receiveFolder.Items.Count

  Dim HasError
  
  Dim CurrentDeletedFolderMailNum
  
  Do While CountWaitSec < CurrentNormSec 
    Do While SaveMailNum < receiveFolder.Items.Count
      'メールを消去(削除済みフォルダ）に移動させると,受信トレーのメールが減るので,ずっと,同じインデックスをアクセスする
      
      'タイムラグ用(一度でも削除等をすれば,タイムラグカウント変数を0にする
      EnterFlag=True
      
      HasError=0
      
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
        Address=OneMailItem.SenderEmailAddress
        Time=OneMailItem.ReceivedTime
        Subject=OneMailItem.Subject
        Body=OneMailItem.Body
        
        'メールの重複調査
        Set TmpOneMailItem=New TmpMailItem
        TmpOneMailItem.SetMailInfo Address,Name,Time,Subject,Body
        
        '重複あるかどうかのフラグ
        Dim HasDuplicated
        HasDuplicated=TimeOrderedFolder.CheckDuplicate(TmpOneMailItem)
           
        
        'こちらは累積のカウントにおけるカウント済みかどうかを示すフラグ
        'たとえ、今カウント真っ最中のメールが,前回カウントしたメールの中で最も新しい日時(どのメールまでをカウントしたかを示すためのもの)より古いものであった場合,
        'それはカウントしていることになるので,重複でカウントしないようにする
        Dim HasUnCounted
        HasUnCounted=False
        
         
        '前回のカウント時の最新のメールの日時（この時刻より前はカウントしてあることを示すもの）より今のメールが後の時刻なら,
        'ようやく未カウントとみなされる
        If LastCountMailDate < Time Then
          HasUnCounted=True
        End If
              
        If LastMailTime < Time Then
          LastMailTime=Time
        End If
        
            
        CurrentMailNum=receiveFolder.Items.Count
        CurrentDeletedFolderMailNum=deletedFolder.Items.Count
        MailOperation=DManager.GetState(Address,Name)
        
        '今のメールアイテムが,「削除済みフォルダ」あるいは「すでにカウント済み(このループで見た)の受信済みフォルダ」内に既にあればそれは重複
        '扱いは完全削除と同じにする
        If HasDuplicated Then
          If MailOperation <> "保存" Then
            MailOperation="完全削除"
          End If
          'そのメールが重複であるならもうメールは数えているということになる
          HasUnCounted=False
        End If
              
        'メールの扱い方（宛先によってメールをどう扱うか)
           
        '保存する場合
        Select  Case MailOperation
          Case  "保存"
           'このメールは消さずに参照するメールのインデックスを1進める
            SaveMailNum=SaveMailNum+1
            DManager.Count Address,Name,Time,HasUnCounted,"ReceivedFolder"
              
          Case"完全削除"
          
            ''VBSの場合,Deleteメソッドは受信済みフォルダのアイテムを削除すると削除済みフォルダへ、削除済みフォルダのアイテムを削除すると完全削除される
            'なので一気にいきなり,受信済みのものを完全削除することはできない
            'ゆえに,まずはいったん,削除済みへ移動する
            
            '受信済みから削除済みに移動と,OneMailItemというメールアイテムオブジェクト変数は参照できなくなる(Nullになってしまう)ので,
            'ここで自作のメール情報置き場オブジェクトを作っておく(削除済みから,完全削除する際,何番目のアイテムを削除すればよいのかを知らなければならないため)
            Set TmpOneMailItem=New TmpMailItem
            TmpOneMailItem.SetMailInfo Address,Name,Time,Subject,Body
            DManager.Count Address,Name,Time,HasUnCounted,"CompDeleted"
            
            
            'まずは削除済みへ移動(これでOneMailItemは参照できなくなる)
            
            OneMailItem.Delete
            
            '移動が完了するまで待つ
            Do While  CurrentMailNum = receiveFolder.Items.Count And  CurrentDeletedFolderMailNum = deletedFolder.Items.Count 
              Wscript.Sleep 500
            Loop
            
            
            '受信済みから削除済みへ移動したOneMailItemという変数はもう使えないので,上で一時的に保存しておいたメール情報をもとに,
            '今移動したアイテムが削除済みフォルダの何番目にあるかを探す(線形探索)
            'こちらは重複判定ではなく単なるインデックス探索なので線形でよい
            Dim DeletedMailIndex
            DeletedMailIndex=MailIndexInFolder(TmpOneMailItem,deletedFolder)
            
            CurrentDeletedFolderMailNum=deletedFolder.Items.Count
            
            'ここで完全削除(削除済みフォルダからも削除する)アイテムが何番目にあるか分かったのでその情報を元に再度メールアイテムオブジェクト変数をつくる
            '完全削除するDeleteメソッドはメールアイテムオブジェクト変数に紐づいているメソッドであるため,当然ながら上の疑似メール情報では削除できない
            Dim CompDeleteMailItem
            Set CompDeleteMailItem=deletedFolder.Items.Item(DeletedMailIndex)
            
            '削除ができるようになるまで待つ
            Do
             On Error Resume Next
              Err.Clear
              CompDeleteMailItem.Delete
             On Error GoTo 0
            Loop While Err.Number <> 0
               
            '削除が完了するまで待つ
            Do While   CurrentDeletedFolderMailNum = deletedFolder.Items.Count
              Wscript.Sleep 500
            Loop
             
            
            Set TmpOneMailItem=Nothing             
         
             
          Case Else
            DManager.Count Address,Name,Time,HasUnCounted,"DeletedFolder"
            CurrentDeletedFolderMailNum=deletedFolder.Items.Count
            OneMailItem.Move deletedFolder
            '移動が完了するまで待つ
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