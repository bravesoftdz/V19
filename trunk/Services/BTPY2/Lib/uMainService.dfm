object SvcSyncBTPY2: TSvcSyncBTPY2
  OldCreateOrder = False
  DisplayName = 'Synchronisation BTP Y2'
  AfterInstall = ServiceAfterInstall
  OnExecute = ServiceExecute
  OnStart = ServiceStart
  OnStop = ServiceStop
  Left = 201
  Top = 118
  Height = 150
  Width = 215
end
