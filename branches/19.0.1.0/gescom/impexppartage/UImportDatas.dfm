object FImportDatas: TFImportDatas
  Left = 468
  Top = 194
  Width = 1030
  Height = 527
  Caption = 'Importation des partages'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 445
    Width = 1014
    Height = 43
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 39
      ClientWidth = 1014
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 39
      ClientAreaWidth = 1014
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        1014
        39)
      object BValider: TToolbarButton97
        Left = 886
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
        ParentFont = False
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 950
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Cancel = True
        Flat = False
        ModalResult = 2
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 982
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
      object bDefaire: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Annuler les modifications'
        Caption = 'Annuler'
        AllowAllUp = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'M0080_S16G1'
        IsControl = True
      end
      object Binsert: TToolbarButton97
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Nouveau'
        AllowAllUp = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BDelete: TToolbarButton97
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Supprimer'
        AllowAllUp = True
        Flat = False
        Visible = False
        GlobalIndexImage = 'Z0005_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 854
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BLANCE: TToolbarButton97
        Left = 919
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Lancer l'#39'int'#233'gration'
        Anchors = [akTop, akRight]
        ParentShowHint = False
        ShowHint = True
        OnClick = BLANCEClick
        GlobalIndexImage = 'O0034_S24G1'
      end
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 0
    Width = 1014
    Height = 41
    Align = alTop
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object HLabel1: THLabel
      Left = 13
      Top = 12
      Width = 63
      Height = 13
      Caption = 'Serveur BDD'
    end
    object BConnect: TToolbarButton97
      Left = 279
      Top = 6
      Width = 24
      Height = 24
      Hint = 'V'#233'rifier la connection'
      ParentShowHint = False
      ShowHint = True
      OnClick = BConnectClick
      GlobalIndexImage = 'O0117_S24G1'
    end
    object LbSaveFile: THLabel
      Left = 328
      Top = 12
      Width = 142
      Height = 13
      Caption = 'Nom du fichier de sauvegarde'
    end
    object BLanceExport: TToolbarButton97
      Left = 744
      Top = 6
      Width = 24
      Height = 24
      OnClick = BLanceExportClick
      GlobalIndexImage = 'O0035_S24G1'
    end
    object SERVERNAME: TEdit
      Left = 80
      Top = 8
      Width = 193
      Height = 21
      TabOrder = 0
    end
    object SaveFileName: THCritMaskEdit
      Left = 496
      Top = 8
      Width = 233
      Height = 21
      AutoSize = False
      TabOrder = 1
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = SaveFileNameElipsisClick
    end
  end
  object PGCONTROL: THPageControl2
    Left = 0
    Top = 41
    Width = 1014
    Height = 404
    ActivePage = TSDB
    Align = alClient
    TabOrder = 2
    object TSDB: TTabSheet
      Caption = 'Bases li'#233'es'
      object LBDB: THListView
        Left = 0
        Top = 0
        Width = 1006
        Height = 376
        Align = alClient
        Columns = <
          item
            AutoSize = True
            Caption = 'Base de donn'#233'e'
          end
          item
            AutoSize = True
            Caption = 'Status'
          end>
        ColumnClick = False
        GridLines = True
        ReadOnly = True
        TabOrder = 0
        ViewStyle = vsReport
      end
    end
    object TSGRP: TTabSheet
      Caption = 'Groupes'
      ImageIndex = 1
      object LBGRPS: THListView
        Left = 0
        Top = 0
        Width = 1006
        Height = 376
        Align = alClient
        Columns = <
          item
            AutoSize = True
            Caption = 'Groupe'
          end
          item
            AutoSize = True
            Caption = 'Status'
          end>
        ColumnClick = False
        GridLines = True
        ReadOnly = True
        TabOrder = 0
        ViewStyle = vsReport
      end
    end
  end
end
