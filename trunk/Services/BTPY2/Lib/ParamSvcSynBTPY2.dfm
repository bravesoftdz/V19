object ParamSvcSyncrho: TParamSvcSyncrho
  Left = 89
  Top = 137
  Width = 665
  Height = 562
  Caption = 'Param'#232'trage du service de synchronisation'
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
  object Tabs: THPageControl2
    Left = 0
    Top = 41
    Width = 649
    Height = 441
    ActivePage = TSParamGeneraux
    Align = alClient
    TabOrder = 0
    object TSParamGeneraux: TTabSheet
      Caption = 'Param'#232'tres g'#233'n'#233'raux'
      DesignSize = (
        641
        413)
      object lSecondTimeout: TLabel
        Left = 3
        Top = 6
        Width = 123
        Height = 13
        Caption = 'D'#233'clenchement toutes les'
      end
      object lSecondTimeoutUnit: TLabel
        Left = 201
        Top = 6
        Width = 46
        Height = 13
        Caption = 'secondes'
      end
      object SecondTimeout: TSpinEdit
        Left = 130
        Top = 4
        Width = 70
        Height = 22
        MaxValue = 1000000
        MinValue = 1
        TabOrder = 0
        Value = 900
        OnClick = SecondTimeoutClick
        OnExit = SecondTimeoutExit
        OnKeyDown = SecondTimeoutKeyDown
        OnKeyUp = SecondTimeoutKeyUp
      end
      object glExecutionPeriodDays: TGroupBox
        Left = 3
        Top = 25
        Width = 638
        Height = 48
        Anchors = [akLeft, akTop, akRight, akBottom]
        Caption = '  Jours d'#39'ex'#233'cution  '
        TabOrder = 1
        object Lundi: TCheckBox
          Left = 17
          Top = 20
          Width = 56
          Height = 17
          Caption = 'Lundi'
          TabOrder = 0
        end
        object Mardi: TCheckBox
          Left = 104
          Top = 20
          Width = 97
          Height = 17
          Caption = 'Mardi'
          TabOrder = 1
        end
        object Mercredi: TCheckBox
          Left = 190
          Top = 20
          Width = 68
          Height = 17
          Caption = 'Mercredi'
          TabOrder = 2
        end
        object Jeudi: TCheckBox
          Left = 277
          Top = 20
          Width = 52
          Height = 17
          Caption = 'Jeudi'
          TabOrder = 3
        end
        object Vendredi: TCheckBox
          Left = 364
          Top = 20
          Width = 63
          Height = 17
          Caption = 'Vendredi'
          TabOrder = 4
        end
        object Samedi: TCheckBox
          Left = 450
          Top = 20
          Width = 57
          Height = 17
          Caption = 'Samedi'
          TabOrder = 5
        end
        object Dimanche: TCheckBox
          Left = 536
          Top = 20
          Width = 73
          Height = 17
          Caption = 'Dimanche'
          TabOrder = 6
        end
      end
      object lExecutionPeriodHours: TGroupBox
        Left = 3
        Top = 74
        Width = 638
        Height = 55
        Anchors = [akLeft, akTop, akRight, akBottom]
        Caption = '  Plage horaire  '
        TabOrder = 2
        object lExecutionPeriodStart: TLabel
          Left = 16
          Top = 23
          Width = 29
          Height = 13
          Caption = 'D'#233'but'
        end
        object lExecutionPeriodEnd: TLabel
          Left = 119
          Top = 22
          Width = 14
          Height = 13
          Caption = 'Fin'
        end
        object ExecutionPeriodStart: TMaskEdit
          Left = 48
          Top = 20
          Width = 57
          Height = 21
          EditMask = '!90:00:00;1;_'
          MaxLength = 8
          TabOrder = 0
          Text = '  :  :  '
          OnExit = ExecutionPeriodStartExit
        end
        object ExecutionPeriodEnd: TMaskEdit
          Left = 150
          Top = 19
          Width = 57
          Height = 21
          EditMask = '!90:00:00;1;_'
          MaxLength = 8
          TabOrder = 1
          Text = '  :  :  '
          OnExit = ExecutionPeriodEndExit
        end
      end
      object LogLevel: TRadioGroup
        Left = 3
        Top = 202
        Width = 638
        Height = 71
        Anchors = [akLeft, akTop, akRight, akBottom]
        Caption = '  Journal  '
        ItemIndex = 0
        Items.Strings = (
          'Aucun'
          'Oui'
          'D'#233'taill'#233)
        TabOrder = 3
        OnClick = LogLevelClick
      end
      object LogType: TRadioGroup
        Left = 3
        Top = 272
        Width = 314
        Height = 71
        Anchors = [akLeft, akTop, akRight, akBottom]
        Caption = '  Type de journal  '
        Items.Strings = (
          'Journalier'
          'D'#233'finir une taille maximum (Mo)')
        TabOrder = 4
        OnClick = LogTypeClick
      end
      object gExchange: TGroupBox
        Left = 3
        Top = 130
        Width = 638
        Height = 71
        Anchors = [akLeft, akTop, akRight, akBottom]
        Caption = '  Echanges  '
        TabOrder = 5
        object lDebugFilesDirectory: TLabel
          Left = 150
          Top = 45
          Width = 168
          Height = 13
          Caption = 'Emplacement xml pi'#232'ces comptable'
          Visible = False
        end
        object Download: TCheckBox
          Left = 17
          Top = 20
          Width = 128
          Height = 17
          Caption = 'R'#233'cup'#233'rer depuis Y2'
          TabOrder = 0
        end
        object Upload: TCheckBox
          Left = 153
          Top = 20
          Width = 128
          Height = 17
          Caption = 'Envoyer vers Y2'
          TabOrder = 1
          OnClick = UploadClick
        end
        object DebugFilesDirectory: THCritMaskEdit
          Left = 320
          Top = 41
          Width = 291
          Height = 21
          TabOrder = 2
          Visible = False
          TagDispatch = 0
          DataType = 'DIRECTORY'
          ElipsisButton = True
        end
      end
      object LogMaxSize: TSpinEdit
        Left = 179
        Top = 316
        Width = 66
        Height = 22
        Anchors = [akTop, akRight, akBottom]
        MaxValue = 1000000
        MinValue = 0
        TabOrder = 6
        Value = 1
        Visible = False
      end
      object LogLevelDebug: TRadioGroup
        Left = 327
        Top = 272
        Width = 314
        Height = 71
        Anchors = [akLeft, akTop, akRight, akBottom]
        Caption = '  Inclure '#233'l'#233'ments de d'#233'bogage   '
        ItemIndex = 0
        Items.Strings = (
          'Aucun'
          'Oui'
          'D'#233'taill'#233)
        TabOrder = 7
        OnClick = LogLevelClick
      end
      object LogSee: TGroupBox
        Left = 3
        Top = 344
        Width = 638
        Height = 65
        Caption = 'Voir les journaux'
        TabOrder = 8
        Visible = False
        DesignSize = (
          638
          65)
        object SeeLogs: THBitBtn
          Left = 9
          Top = 23
          Width = 28
          Height = 27
          Hint = 'Voir les journaux'
          Anchors = [akTop, akRight]
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnClick = SeeLogsClick
          Margin = 3
          GlobalIndexImage = 'Z1327_S16G1'
        end
      end
    end
    object TSUpdateFrequency: TTabSheet
      Caption = 'Fr'#233'quence de mise '#224' jour des tables'
      ImageIndex = 1
      object gFrequency1: THGrid
        Left = 0
        Top = 0
        Width = 641
        Height = 413
        Align = alClient
        Anchors = [akTop, akBottom]
        ColCount = 8
        DefaultColWidth = 50
        DefaultRowHeight = 18
        TabOrder = 0
        OnDblClick = gFrequency1DblClick
        OnMouseUp = gFrequency1MouseUp
        SortedCol = -1
        Titres.Strings = (
          ''
          'Table'
          'A chaque fois'
          'Une fois par jour'
          'Une fois par mois'
          'Une fois par an'
          'Une seule fois'
          'Jamais')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        DBIndicator = True
        ColWidths = (
          10
          50
          50
          50
          50
          50
          50
          50)
      end
    end
    object tFolders: TTabSheet
      Caption = 'Dossiers'
      ImageIndex = 2
      object lFolderChoice: TLabel
        Left = 22
        Top = 16
        Width = 35
        Height = 13
        Caption = 'Dossier'
      end
      object NewFolder: THBitBtn
        Left = 174
        Top = 17
        Width = 31
        Height = 27
        Hint = 'Ajouter'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = NewFolderClick
        Margin = 4
        GlobalIndexImage = 'Z2293_S24G1'
      end
      object DelFolder: THBitBtn
        Left = 243
        Top = 17
        Width = 31
        Height = 27
        Hint = 'Supprimer'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = DelFolderClick
        Margin = 4
        GlobalIndexImage = 'Z0204_S16G1'
      end
      object cFolderChoice: TListBox
        Left = 60
        Top = 16
        Width = 112
        Height = 73
        ItemHeight = 13
        TabOrder = 0
        OnClick = cFolderChoiceClick
        OnDblClick = UpdateFolderClick
      end
      object gFolderParam: TGroupBox
        Left = 22
        Top = 96
        Width = 531
        Height = 193
        Caption = '  Param'#232'tres  '
        TabOrder = 4
        object lLastSynchro: TLabel
          Left = 4
          Top = 20
          Width = 140
          Height = 13
          Caption = 'Date derni'#232're synchronisation'
        end
        object lBTPUser: TLabel
          Left = 4
          Top = 44
          Width = 72
          Height = 13
          Caption = 'Code utilisateur'
        end
        object lBTPServer: TLabel
          Left = 104
          Top = 68
          Width = 37
          Height = 13
          Caption = 'Serveur'
        end
        object lBTP: TLabel
          Left = 4
          Top = 68
          Width = 21
          Height = 13
          Caption = 'BTP'
        end
        object lBTPFolder: TLabel
          Left = 104
          Top = 92
          Width = 35
          Height = 13
          Caption = 'Dossier'
        end
        object lY2Server: TLabel
          Left = 104
          Top = 116
          Width = 37
          Height = 13
          Caption = 'Serveur'
        end
        object lY2: TLabel
          Left = 4
          Top = 116
          Width = 13
          Height = 13
          Caption = 'Y2'
        end
        object lY2Folder: TLabel
          Left = 104
          Top = 140
          Width = 35
          Height = 13
          Caption = 'Dossier'
        end
        object LastSynchro: TMaskEdit
          Left = 148
          Top = 12
          Width = 117
          Height = 21
          Enabled = False
          EditMask = '!90/00/0000 90:00:00;1;_'
          MaxLength = 19
          ParentColor = True
          TabOrder = 0
          Text = '  /  /       :  :  '
        end
        object BTPUser: TMaskEdit
          Left = 148
          Top = 36
          Width = 53
          Height = 21
          Enabled = False
          MaxLength = 3
          TabOrder = 1
        end
        object BTPServer: TMaskEdit
          Left = 148
          Top = 60
          Width = 324
          Height = 21
          Enabled = False
          MaxLength = 70
          TabOrder = 2
        end
        object BTPFolder: TMaskEdit
          Left = 148
          Top = 84
          Width = 324
          Height = 21
          Enabled = False
          MaxLength = 70
          TabOrder = 3
        end
        object Y2Server: TMaskEdit
          Left = 148
          Top = 108
          Width = 324
          Height = 21
          Enabled = False
          MaxLength = 70
          TabOrder = 4
        end
        object Y2Folder: TMaskEdit
          Left = 148
          Top = 132
          Width = 324
          Height = 21
          Enabled = False
          MaxLength = 70
          TabOrder = 5
        end
        object AddNewFolder: THBitBtn
          Left = 478
          Top = 158
          Width = 31
          Height = 27
          Hint = 'Valider la saisie'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 6
          Visible = False
          OnClick = AddNewFolderClick
          Margin = 4
          NumGlyphs = 2
          GlobalIndexImage = 'Z0003_S16G2'
        end
        object CancelNewFolder: THBitBtn
          Left = 444
          Top = 158
          Width = 31
          Height = 27
          Hint = 'Valider la saisie'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 7
          Visible = False
          OnClick = CancelNewFolderClick
          Margin = 4
          GlobalIndexImage = 'M0080_S16G1'
        end
        object ConnectBTP: THBitBtn
          Left = 478
          Top = 70
          Width = 31
          Height = 27
          Hint = 'Tester la connexion BTP'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 8
          Visible = False
          OnClick = TestConnectClick
          Margin = 4
          GlobalIndexImage = 'Z1176_S16G1'
        end
        object ConnectY2: THBitBtn
          Left = 478
          Top = 119
          Width = 31
          Height = 27
          Hint = 'Tester la connexion Y2'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 9
          Visible = False
          OnClick = TestConnectClick
          Margin = 4
          GlobalIndexImage = 'Z1176_S16G1'
        end
      end
      object UpdateFolder: THBitBtn
        Left = 208
        Top = 17
        Width = 31
        Height = 27
        Hint = 'Modifier'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = UpdateFolderClick
        Margin = 4
        GlobalIndexImage = 'Z2104_S24G1'
      end
    end
    object tService: TTabSheet
      Caption = 'Gestion du service'
      ImageIndex = 3
      object gSvcState: TGroupBox
        Left = 8
        Top = 2
        Width = 617
        Height = 65
        Caption = '  Etat du service  '
        TabOrder = 0
        DesignSize = (
          617
          65)
        object ServiceState: TLabel
          Left = 42
          Top = 28
          Width = 162
          Height = 13
          Caption = '_______________________'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object SvcStateRefresh: THBitBtn
          Left = 9
          Top = 21
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
      object gSvcManagement: TGroupBox
        Left = 8
        Top = 69
        Width = 617
        Height = 338
        Caption = '  Gestion  '
        TabOrder = 1
        object SvcInstall: THBitBtn
          Left = 9
          Top = 20
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
        object SvcUnInstall: THBitBtn
          Left = 122
          Top = 20
          Width = 93
          Height = 27
          Hint = 'Arr'#234'ter'
          Caption = 'D'#233'sinstaller'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnClick = SvcUnInstallClick
          Margin = 4
          GlobalIndexImage = 'Z0072_S16G1'
        end
        object SvcStart: THBitBtn
          Left = 235
          Top = 20
          Width = 93
          Height = 27
          Hint = 'D'#233'marrer'
          Caption = 'D'#233'marrer'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
          OnClick = SvcStartClick
          Margin = 4
          GlobalIndexImage = 'Z1449_S16G1'
        end
        object SvcStop: THBitBtn
          Left = 348
          Top = 20
          Width = 93
          Height = 27
          Hint = 'Arr'#234'ter'
          Caption = 'Arr'#234'ter'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 3
          OnClick = SvcStopClick
          Margin = 4
          GlobalIndexImage = 'Z0447_S16G1'
        end
        object gSvcInstalParam: TGroupBox
          Left = 9
          Top = 56
          Width = 599
          Height = 271
          Caption = '  Param'#232'tres  '
          TabOrder = 4
          object SvcInstallType: TRadioGroup
            Left = 12
            Top = 80
            Width = 578
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
            TabOrder = 0
          end
          object gSvcInstallConnection: TGroupBox
            Left = 11
            Top = 130
            Width = 578
            Height = 133
            Caption = '  Connexion  '
            TabOrder = 1
            object lInstallAccount: TLabel
              Left = 155
              Top = 60
              Width = 36
              Height = 13
              Caption = 'Compte'
              Visible = False
            end
            object lInstallPwd: TLabel
              Left = 155
              Top = 83
              Width = 64
              Height = 13
              Caption = 'Mot de passe'
              Visible = False
            end
            object lInstallPwdConfirm: TLabel
              Left = 155
              Top = 109
              Width = 124
              Height = 13
              Caption = 'Confirmation mot de passe'
              Visible = False
            end
            object SvcInstallAccount: TEdit
              Left = 281
              Top = 59
              Width = 252
              Height = 21
              TabOrder = 0
              Visible = False
            end
            object SvcInstallPwd: TEdit
              Left = 281
              Top = 81
              Width = 252
              Height = 21
              Font.Charset = ANSI_CHARSET
              Font.Color = clBlack
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentFont = False
              PasswordChar = '*'
              TabOrder = 1
              Visible = False
            end
            object SvcInstallPwdConfirm: TEdit
              Left = 281
              Top = 106
              Width = 252
              Height = 21
              PasswordChar = '*'
              TabOrder = 2
              Visible = False
            end
            object SvcAccountTipe: TRadioGroup
              Left = 11
              Top = 20
              Width = 534
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
              TabOrder = 3
              OnClick = SvcAccountTipeClick
            end
          end
          object gSvcInstallPath: TGroupBox
            Left = 12
            Top = 22
            Width = 578
            Height = 57
            Caption = '  Emplacement du service  '
            TabOrder = 2
            object SvcInstallPath: THCritMaskEdit
              Left = 19
              Top = 21
              Width = 450
              Height = 21
              TabOrder = 0
              TagDispatch = 0
              DataType = 'DIRECTORY'
              ElipsisButton = True
            end
          end
        end
      end
    end
  end
  object pTitle: TPanel
    Left = 0
    Top = 0
    Width = 649
    Height = 41
    Align = alTop
    TabOrder = 1
    DesignSize = (
      649
      41)
    object LGenerateIn: TLabel
      Left = 16
      Top = 15
      Width = 180
      Height = 13
      Caption = 'Fichier SvcSynBTPY2.ini g'#233'n'#233'r'#233' dans'
    end
    object GenerateIn: TEdit
      Left = 197
      Top = 12
      Width = 412
      Height = 21
      Anchors = [akTop, akRight]
      Enabled = False
      TabOrder = 0
      Text = 'GenerateIn'
    end
    object Reload: THBitBtn
      Left = 615
      Top = 9
      Width = 28
      Height = 27
      Hint = 'Recharger'
      Anchors = [akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Visible = False
      OnClick = ReloadClick
      Margin = 4
      GlobalIndexImage = 'Z0257_S16G1'
    end
  end
  object pButtons: TPanel
    Left = 0
    Top = 482
    Width = 649
    Height = 41
    Align = alBottom
    TabOrder = 2
    DesignSize = (
      649
      41)
    object bValider: THBitBtn
      Left = 617
      Top = 7
      Width = 28
      Height = 27
      Hint = 'G'#233'n'#233'rer'
      Anchors = [akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = bValiderClick
      Margin = 4
      GlobalIndexImage = 'Z2159_S16G1'
    end
    object bFermer: THBitBtn
      Left = 585
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Anchors = [akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = bFermerClick
      Margin = 3
      GlobalIndexImage = 'Z0021_S16G1'
    end
  end
end
