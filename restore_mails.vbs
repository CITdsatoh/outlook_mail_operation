 
  Option Explicit
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
  '���[��
  Dim OneMailItem
  
  Dim CurrentMailNum
 
 '�܂��͍폜�ς݃A�C�e���ɂ��Ă݂Ă䂭
  Do While 0 < deletedFolder.Items.Count
     CurrentMailNum=deletedFolder.Items.Count
     Set OneMailItem=deletedFolder.Items.Item(1)
     OneMailItem.Move receiveFolder
     '�ړ��͔񓯊������Ȃ̂ňړ����I���܂ő҂�
     Do While  CurrentMailNum = deletedFolder.Items.Count 
           Wscript.Sleep 500
       Loop

  Loop