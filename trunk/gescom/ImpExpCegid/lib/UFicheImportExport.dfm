object FexpImpCegid: TFexpImpCegid
  Left = 649
  Top = 205
  Width = 578
  Height = 700
  Caption = 'Import - Export de donn'#233'es CEGID'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 618
    Width = 562
    Height = 43
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 39
      ClientWidth = 562
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 39
      ClientAreaWidth = 562
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        562
        39)
      object BValider: TToolbarButton97
        Left = 466
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
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 498
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
        Left = 530
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
        OnClick = HelpBtnClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 434
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
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
    end
  end
  object PGCTRL: THPageControl2
    Left = 0
    Top = 0
    Width = 562
    Height = 618
    ActivePage = TPSHTCAR
    Align = alClient
    TabOrder = 1
    object TPSHTCAR: TTabSheet
      Caption = 'Traitements'
      object Bevel1: TBevel
        Left = 2
        Top = 122
        Width = 535
        Height = 132
        Shape = bsFrame
      end
      object Bevel2: TBevel
        Left = 1
        Top = 275
        Width = 537
        Height = 246
        Shape = bsFrame
      end
      object Label3: TLabel
        Left = 0
        Top = 0
        Width = 554
        Height = 13
        Align = alTop
      end
      object Label5: TLabel
        Left = 0
        Top = 69
        Width = 554
        Height = 20
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'Assurez-vous qu'#39'il reste suffisamment de place sur les disques'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -16
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label6: TLabel
        Left = 0
        Top = 13
        Width = 554
        Height = 23
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'ATTENTION '
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -16
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label7: TLabel
        Left = 0
        Top = 56
        Width = 554
        Height = 13
        Align = alTop
      end
      object Label4: TLabel
        Left = 0
        Top = 36
        Width = 554
        Height = 20
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'Les temps de traitements peuvent '#234'tre tr'#232's importants'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -16
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label1: TLabel
        Left = 15
        Top = 157
        Width = 174
        Height = 13
        Caption = 'R'#233'pertoire de destination des fichiers'
      end
      object Label2: TLabel
        Left = 15
        Top = 180
        Width = 108
        Height = 13
        Caption = 'Nom du fichier d'#39'export'
      end
      object RDTEXPORT: TRadioButton
        Left = 13
        Top = 116
        Width = 185
        Height = 17
        Caption = 'Exports des donn'#233'es'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        TabStop = True
        OnClick = RDTEXPORTClick
      end
      object RDTIMPORT: TRadioButton
        Left = 13
        Top = 266
        Width = 185
        Height = 17
        Caption = 'Import des donn'#233'es'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 3
        OnClick = RDTIMPORTClick
      end
      object GroupBox1: TGroupBox
        Left = 10
        Top = 292
        Width = 515
        Height = 57
        TabOrder = 4
        object Label8: TLabel
          Left = 15
          Top = 24
          Width = 155
          Height = 13
          Caption = 'Nom du fichier d'#39'export '#224' int'#233'grer'
        end
        object IMPORTFIC: THCritMaskEdit
          Left = 207
          Top = 20
          Width = 298
          Height = 21
          TabOrder = 0
          TagDispatch = 0
          DataType = 'OPENFILE(*.ZIP)'
          ElipsisButton = True
        end
        object CHKZIP: TCheckBox
          Left = 16
          Top = 0
          Width = 105
          Height = 17
          Caption = 'Via un fichier Zip'
          Checked = True
          State = cbChecked
          TabOrder = 1
          OnClick = CHKZIPClick
        end
      end
      object GroupBox2: TGroupBox
        Left = 9
        Top = 368
        Width = 515
        Height = 57
        TabOrder = 5
        object Label9: TLabel
          Left = 14
          Top = 23
          Width = 167
          Height = 13
          Caption = 'R'#233'pertoire de stockage des fichiers'
        end
        object CHKREPERT: TCheckBox
          Left = 10
          Top = 0
          Width = 120
          Height = 17
          Caption = 'Depuis un r'#233'pertoire'
          TabOrder = 0
          OnClick = CHKREPERTClick
        end
        object IMPORTDIR: THCritMaskEdit
          Left = 208
          Top = 19
          Width = 298
          Height = 21
          Enabled = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'DIRECTORY'
          ElipsisButton = True
        end
      end
      object CHKMSGSTRUCT: TCheckBox
        Left = 8
        Top = 442
        Width = 337
        Height = 17
        Alignment = taLeftJustify
        Caption = 
          'Demande de confirmation de traitement sur structures diff'#233'rentes' +
          ' ?'
        Enabled = False
        TabOrder = 6
      end
      object CHKSTOPONERROR: TCheckBox
        Left = 8
        Top = 463
        Width = 337
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Arr'#234't sur anomalie d'#39'int'#233'gration ?'
        Checked = True
        Enabled = False
        State = cbChecked
        TabOrder = 7
      end
      object CBVIDAGEEXP: TCheckBox
        Left = 408
        Top = 232
        Width = 120
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Vidage du r'#233'pertoire'
        Checked = True
        State = cbChecked
        TabOrder = 8
      end
      object CBVIDAGEIMP: TCheckBox
        Left = 8
        Top = 483
        Width = 337
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Vidage du r'#233'pertoire'
        Checked = True
        Enabled = False
        State = cbChecked
        TabOrder = 9
      end
      object CBCOMPTA: TCheckBox
        Left = 16
        Top = 208
        Width = 105
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Comptabilit'#233
        TabOrder = 10
        OnClick = CBCOMPTAClick
      end
      object CBPAYE: TCheckBox
        Left = 16
        Top = 232
        Width = 105
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Paye'
        TabOrder = 11
        OnClick = CBPAYEClick
      end
      object EXPORTFIC: THCritMaskEdit
        Left = 225
        Top = 153
        Width = 298
        Height = 21
        TabOrder = 0
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
      end
      object NOMFIC: THCritMaskEdit
        Left = 225
        Top = 176
        Width = 288
        Height = 21
        TabOrder = 1
        TagDispatch = 0
      end
      object CBDOSSIER: TCheckBox
        Left = 137
        Top = 219
        Width = 105
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Dossier complet'
        Checked = True
        State = cbChecked
        TabOrder = 12
        OnClick = CBDOSSIERClick
      end
    end
    object TBSHTCONTROL: TTabSheet
      Caption = 'Cont'#244'les'
      ImageIndex = 1
      object HPanel1: THPanel
        Left = 0
        Top = 549
        Width = 554
        Height = 41
        Align = alBottom
        BevelOuter = bvNone
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object BLANCETRAIT: TToolbarButton97
          Left = 200
          Top = 8
          Width = 153
          Height = 27
          Caption = 'Lancer le traitement'
          Flat = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          Opaque = False
          ParentFont = False
          Visible = False
          OnClick = BLANCETRAITClick
        end
      end
      object GS: THGrid
        Left = 0
        Top = 0
        Width = 554
        Height = 549
        Align = alClient
        ColCount = 1
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goVertLine, goHorzLine, goRowSelect]
        TabOrder = 1
        OnDblClick = GSDblClick
        SortedCol = -1
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          553)
      end
      object TX: THCritMaskEdit
        Left = 80
        Top = 432
        Width = 65
        Height = 21
        MaxLength = 6
        TabOrder = 2
        Visible = False
        OnExit = TXExit
        TagDispatch = 0
      end
    end
    object RAPPORT: TTabSheet
      Caption = 'RAPPORT'
      ImageIndex = 2
      object TRACE: TListBox
        Left = 0
        Top = 0
        Width = 554
        Height = 590
        Align = alClient
        BiDiMode = bdLeftToRight
        ItemHeight = 13
        ParentBiDiMode = False
        TabOrder = 0
      end
    end
  end
  object hmtrad: THSystemMenu
    Separator = True
    Traduction = False
    Left = 236
    Top = 280
  end
end
