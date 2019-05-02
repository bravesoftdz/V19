object ParamSvcSyncrho: TParamSvcSyncrho
  Left = 85
  Top = 161
  Width = 939
  Height = 677
  Caption = 'Gestion des services'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object gServicesLst: TGroupBox
    Left = 0
    Top = 0
    Width = 458
    Height = 603
    Align = alClient
    Caption = 'Liste des services'
    TabOrder = 0
    object SvcList: THGrid
      Left = 2
      Top = 81
      Width = 454
      Height = 520
      Align = alClient
      ColCount = 3
      DefaultColWidth = 50
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 0
      OnClick = SvcListClick
      SortedCol = 2
      Titres.Strings = (
        ''
        'Nom long'
        'Nom court'
        '')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = True
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      DBIndicator = True
      ColWidths = (
        10
        551
        77)
    end
    object pnl3: TPanel
      Left = 2
      Top = 15
      Width = 454
      Height = 66
      Align = alTop
      TabOrder = 1
      DesignSize = (
        454
        66)
      object lSvcSearch: TLabel
        Left = 6
        Top = 40
        Width = 56
        Height = 13
        Caption = 'Rechercher'
      end
      object SvcLstRefresh: THBitBtn
        Left = 422
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Actualiser la liste'
        Anchors = []
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = SvcLstRefreshClick
        Margin = 4
        GlobalIndexImage = 'M0063_S16G1'
      end
      object ChkLSESrv: TCheckBox
        Left = 6
        Top = 9
        Width = 201
        Height = 17
        Caption = 'Afficher uniquement les services LSE'
        TabOrder = 1
      end
      object SvcSearch: TEdit
        Left = 67
        Top = 36
        Width = 350
        Height = 21
        TabOrder = 2
      end
      object SvcLstSearch: THBitBtn
        Left = 422
        Top = 33
        Width = 28
        Height = 27
        Hint = 'Rechercher'
        Anchors = []
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = SvcLstSearchClick
        Margin = 4
        GlobalIndexImage = 'Z0002_S16G1'
      end
    end
  end
  object gServiceMangement: TGroupBox
    Left = 458
    Top = 0
    Width = 465
    Height = 603
    Align = alRight
    Caption = 'Gestion'
    TabOrder = 1
    object gSvcState: TGroupBox
      Left = 8
      Top = 30
      Width = 449
      Height = 201
      Caption = ' Informations '
      TabOrder = 0
      object lSvcState: TLabel
        Left = 9
        Top = 27
        Width = 19
        Height = 13
        Caption = 'Etat'
      end
      object lSvcStart: TLabel
        Left = 9
        Top = 52
        Width = 92
        Height = 13
        Caption = 'Type de d'#233'marrage'
      end
      object lSvcAccount: TLabel
        Left = 9
        Top = 76
        Width = 75
        Height = 13
        Caption = 'Compte associ'#233
      end
      object lSvcLocation: TLabel
        Left = 9
        Top = 101
        Width = 64
        Height = 13
        Caption = 'Emplacement'
      end
      object ServiceState: TEdit
        Left = 106
        Top = 23
        Width = 327
        Height = 21
        Enabled = False
        TabOrder = 0
      end
      object ServiceTypeStart: TEdit
        Left = 106
        Top = 48
        Width = 327
        Height = 21
        Enabled = False
        TabOrder = 1
      end
      object SvcAccount: TEdit
        Left = 106
        Top = 72
        Width = 327
        Height = 21
        Enabled = False
        TabOrder = 2
      end
      object SvcLocalisation: TEdit
        Left = 106
        Top = 97
        Width = 327
        Height = 21
        Enabled = False
        TabOrder = 3
      end
      object Panel1: TPanel
        Left = 2
        Top = 164
        Width = 445
        Height = 35
        Align = alBottom
        TabOrder = 7
        DesignSize = (
          445
          35)
        object SvcStateRefresh: THBitBtn
          Left = 408
          Top = 5
          Width = 28
          Height = 27
          Hint = 'Actualiser l'#39#233'tat'
          Anchors = []
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnClick = SvcStateRefreshClick
          Margin = 4
          GlobalIndexImage = 'M0063_S16G1'
        end
      end
      object SvcUnInstall: THBitBtn
        Left = 108
        Top = 129
        Width = 93
        Height = 27
        Hint = 'Arr'#234'ter'
        Caption = 'D'#233'sinstaller'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        OnClick = SvcUnInstallClick
        Margin = 4
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object SvcStart: THBitBtn
        Left = 215
        Top = 129
        Width = 93
        Height = 27
        Hint = 'D'#233'marrer'
        Caption = 'D'#233'marrer'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 5
        OnClick = SvcStartClick
        Margin = 4
        GlobalIndexImage = 'Z1449_S16G1'
      end
      object SvcStop: THBitBtn
        Left = 324
        Top = 129
        Width = 93
        Height = 27
        Hint = 'Arr'#234'ter'
        Caption = 'Arr'#234'ter'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 6
        OnClick = SvcStopClick
        Margin = 4
        GlobalIndexImage = 'Z0447_S16G1'
      end
      object SvcSystem: THBitBtn
        Left = 108
        Top = 129
        Width = 309
        Height = 27
        Hint = 'Arr'#234'ter'
        Caption = 'Service syst'#232'me. Cliquez ici pour le g'#233'rer'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 8
        OnClick = SvcSystemClick
        Margin = 4
        GlobalIndexImage = 'Z2298_S24G1'
      end
    end
    object gSvcManagement: TGroupBox
      Left = 8
      Top = 247
      Width = 449
      Height = 342
      Caption = '  Installer un service '
      TabOrder = 1
      object gSvcInstallPath: TGroupBox
        Left = 12
        Top = 24
        Width = 413
        Height = 57
        Caption = '  Emplacement du service  '
        TabOrder = 0
        object SvcInstallPath: THCritMaskEdit
          Left = 10
          Top = 21
          Width = 396
          Height = 21
          TabOrder = 0
          TagDispatch = 0
          DataType = 'OPENFILE(*.EXE)'
          ElipsisButton = True
        end
      end
      object SvcInstallType: TRadioGroup
        Left = 12
        Top = 87
        Width = 413
        Height = 49
        Caption = '  Type de d'#233'marrage  '
        Columns = 3
        Ctl3D = True
        ItemIndex = 0
        Items.Strings = (
          'Manuel'
          'Automatique'
          'D'#233'sactiv'#233)
        ParentCtl3D = False
        TabOrder = 1
      end
      object gSvcInstallConnection: TGroupBox
        Left = 12
        Top = 144
        Width = 414
        Height = 143
        Caption = '  Connexion  '
        TabOrder = 2
        object lInstallAccount: TLabel
          Left = 12
          Top = 66
          Width = 36
          Height = 13
          Caption = 'Compte'
          Visible = False
        end
        object lInstallPwd: TLabel
          Left = 12
          Top = 87
          Width = 64
          Height = 13
          Caption = 'Mot de passe'
          Visible = False
        end
        object lInstallPwdConfirm: TLabel
          Left = 12
          Top = 116
          Width = 124
          Height = 13
          Caption = 'Confirmation mot de passe'
          Visible = False
        end
        object SvcInstallAccount: TEdit
          Left = 138
          Top = 62
          Width = 252
          Height = 21
          TabOrder = 1
          Visible = False
        end
        object SvcInstallPwd: TEdit
          Left = 138
          Top = 87
          Width = 252
          Height = 21
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          PasswordChar = '*'
          TabOrder = 2
          Visible = False
        end
        object SvcInstallPwdConfirm: TEdit
          Left = 138
          Top = 112
          Width = 252
          Height = 21
          PasswordChar = '*'
          TabOrder = 3
          Visible = False
        end
        object SvcAccountTipe: TRadioGroup
          Left = 11
          Top = 20
          Width = 390
          Height = 37
          Caption = '  Type de compte  '
          Color = clBtnFace
          Columns = 2
          Ctl3D = True
          ItemIndex = 0
          Items.Strings = (
            'Syst'#232'me local'
            'Autre compte')
          ParentColor = False
          ParentCtl3D = False
          TabOrder = 0
          OnClick = SvcAccountTipeClick
        end
      end
      object pnl2: TPanel
        Left = 2
        Top = 305
        Width = 445
        Height = 35
        Align = alBottom
        TabOrder = 3
        object SvcInstall: THBitBtn
          Left = 343
          Top = 4
          Width = 93
          Height = 27
          Hint = 'Installer'
          Caption = 'Installer'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnClick = SvcInstallClick
          Margin = 4
          GlobalIndexImage = 'Z0466_S16G1'
        end
      end
    end
  end
  object pButtons: TPanel
    Left = 0
    Top = 603
    Width = 923
    Height = 35
    Align = alBottom
    TabOrder = 2
    DesignSize = (
      923
      35)
    object bFermer: THBitBtn
      Left = 890
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Anchors = [akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = bFermerClick
      Margin = 3
      GlobalIndexImage = 'Z0021_S16G1'
    end
  end
end
