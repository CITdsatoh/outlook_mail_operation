 
  Option Explicit
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
  'メール
  Dim OneMailItem
  
  Dim CurrentMailNum
 
 'まずは削除済みアイテムについてみてゆく
  Do While 0 < deletedFolder.Items.Count
     CurrentMailNum=deletedFolder.Items.Count
     Set OneMailItem=deletedFolder.Items.Item(1)
     OneMailItem.Move receiveFolder
     '移動は非同期処理なので移動が終わるまで待つ
     Do While  CurrentMailNum = deletedFolder.Items.Count 
           Wscript.Sleep 500
       Loop

  Loop