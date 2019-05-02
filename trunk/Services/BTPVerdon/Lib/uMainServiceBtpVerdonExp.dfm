object SvcSyncBTPVerdonExp: TSvcSyncBTPVerdonExp
  OldCreateOrder = False
  DisplayName = 'Synchronisation BTP VERDON - Export'
  AfterInstall = ServiceAfterInstall
  OnExecute = ServiceExecute
  OnStart = ServiceStart
  OnStop = ServiceStop
  Left = 198
  Top = 117
  Height = 150
  Width = 215
end
