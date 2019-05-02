object FExportData: TFExportData
  Left = 423
  Top = 271
  Width = 923
  Height = 572
  Caption = 'Export des partages'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 498
    Width = 907
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 907
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 31
      ClientAreaWidth = 907
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        907
        31)
      object BFerme: TToolbarButton97
        Left = 843
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Cancel = True
        Flat = False
        ModalResult = 2
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 875
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 0
    Width = 907
    Height = 41
    Align = alTop
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object LNomDB: THLabel
      Left = 327
      Top = 12
      Width = 77
      Height = 13
      Caption = 'Base R'#233'f'#233'rence'
    end
    object HLabel1: THLabel
      Left = 13
      Top = 12
      Width = 63
      Height = 13
      Caption = 'Serveur BDD'
    end
    object BConnect: TToolbarButton97
      Left = 279
      Top = 7
      Width = 24
      Height = 24
      Hint = 'Se connecter'
      ParentShowHint = False
      ShowHint = True
      OnClick = BConnectClick
      GlobalIndexImage = 'O0117_S24G1'
    end
    object SERVERNAME: TEdit
      Left = 80
      Top = 8
      Width = 193
      Height = 21
      TabOrder = 0
    end
    object DBNAME: THValComboBox
      Left = 416
      Top = 8
      Width = 185
      Height = 21
      ItemHeight = 13
      TabOrder = 1
      TagDispatch = 0
    end
  end
  object HPanel4: THPanel
    Left = 0
    Top = 450
    Width = 907
    Height = 48
    Align = alBottom
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object LbSaveFile: THLabel
      Left = 16
      Top = 13
      Width = 142
      Height = 13
      Caption = 'Nom du fichier de sauvegarde'
    end
    object BLanceExport: TToolbarButton97
      Left = 424
      Top = 8
      Width = 24
      Height = 24
      Hint = 'Lancer l'#39'export'
      ParentShowHint = False
      ShowHint = True
      OnClick = BLanceExportClick
      GlobalIndexImage = 'O0035_S24G1'
    end
    object SaveFileName: THCritMaskEdit
      Left = 168
      Top = 10
      Width = 233
      Height = 21
      AutoSize = False
      TabOrder = 0
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = SaveFileNameElipsisClick
    end
  end
  object PGCONTROL: THPageControl2
    Left = 0
    Top = 41
    Width = 907
    Height = 409
    ActivePage = TSDB
    Align = alClient
    TabOrder = 3
    object TSDB: TTabSheet
      Caption = 'Bases li'#233'es'
      object LBDB: TListBox
        Left = 0
        Top = 0
        Width = 899
        Height = 381
        Align = alClient
        ItemHeight = 13
        TabOrder = 0
      end
    end
    object TSGRP: TTabSheet
      Caption = 'Groupes'
      ImageIndex = 1
      object LBGRPS: TListBox
        Left = 0
        Top = 0
        Width = 899
        Height = 381
        Align = alClient
        ItemHeight = 13
        TabOrder = 0
      end
    end
    object TSDATAS: TTabSheet
      Caption = 'Donn'#233'es'
      ImageIndex = 2
      object LBDATAS: TListBox
        Left = 0
        Top = 0
        Width = 899
        Height = 381
        Align = alClient
        ItemHeight = 13
        TabOrder = 0
      end
    end
  end
end
