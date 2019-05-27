object Form1: TForm1
  Left = 51
  Top = 0
  Align = alCustom
  Caption = 'Form1'
  ClientHeight = 701
  ClientWidth = 1284
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = [fsBold]
  OldCreateOrder = False
  Position = poDesigned
  WindowState = wsMaximized
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 120
  TextHeight = 16
  object pcShowInfo: TPageControl
    Left = 0
    Top = 0
    Width = 1284
    Height = 701
    ActivePage = TabSheet4
    Align = alClient
    TabOrder = 0
    OnChange = pcShowInfoChange
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object GroupBox1: TGroupBox
        Left = 643
        Top = 153
        Width = 633
        Height = 517
        Align = alClient
        Caption = #20135#21697#39044#35272
        TabOrder = 0
        object wbShowProinfo: TWebBrowser
          AlignWithMargins = True
          Left = 5
          Top = 21
          Width = 623
          Height = 491
          Align = alClient
          TabOrder = 0
          ExplicitWidth = 498
          ExplicitHeight = 405
          ControlData = {
            4C00000064400000BF3200000000000000000000000000000000000000000000
            000000004C000000000000000000000001000000E0D057007335CF11AE690800
            2B2E126208000000000000004C0000000114020000000000C000000000000046
            8000000000000000000000000000000000000000000000000000000000000000
            00000000000000000100000000000000000000000000000000000000}
        end
      end
      object GroupBox2: TGroupBox
        Left = 0
        Top = 0
        Width = 1276
        Height = 153
        Align = alTop
        Caption = #20135#21697#21015#34920#39029#32593#22336
        TabOrder = 1
        object memListUrl: TMemo
          Left = 3
          Top = 24
          Width = 640
          Height = 121
          ImeName = #20013#25991'('#31616#20307') - '#30334#24230#36755#20837#27861
          ScrollBars = ssBoth
          TabOrder = 0
        end
        object BitBtn3: TBitBtn
          Left = 680
          Top = 40
          Width = 153
          Height = 49
          Caption = #20135#21697#21015#34920
          TabOrder = 1
          OnClick = BitBtn3Click
        end
        object BitBtn4: TBitBtn
          Left = 967
          Top = 40
          Width = 162
          Height = 49
          Caption = #19979#19968#27493
          TabOrder = 2
          OnClick = BitBtn4Click
        end
      end
      object GroupBox3: TGroupBox
        Left = 0
        Top = 153
        Width = 643
        Height = 517
        Align = alLeft
        Caption = #21015#34920#35814#24773
        TabOrder = 2
        object sgShowTitle: TStringGrid
          Left = 2
          Top = 18
          Width = 639
          Height = 449
          Align = alClient
          RowCount = 1
          FixedRows = 0
          TabOrder = 0
          OnClick = sgShowTitleClick
          OnDrawCell = sgShowTitleDrawCell
          OnSelectCell = sgShowTitleSelectCell
        end
        object Panel1: TPanel
          Left = 2
          Top = 467
          Width = 639
          Height = 48
          Align = alBottom
          Caption = 'Panel1'
          TabOrder = 1
          object lbSelectTitle: TLabel
            Left = 32
            Top = 8
            Width = 4
            Height = 16
          end
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      object e: TGroupBox
        Left = -4
        Top = 427
        Width = 154
        Height = 60
        Caption = #20135#21697#24120#35268#39033#35774#32622
        TabOrder = 0
        object GroupBox8: TGroupBox
          Left = 19
          Top = 23
          Width = 352
          Height = 212
          Caption = #20135#21697#22823#31867
          TabOrder = 0
          object cbProductsClass1: TComboBox
            Left = 12
            Top = 23
            Width = 149
            Height = 24
            ImeName = #20013#25991'('#31616#20307') - '#30334#24230#36755#20837#27861
            TabOrder = 0
            OnSelect = cbProductsClass1Select
          end
          object rgProductsclassSub: TRadioGroup
            Left = 183
            Top = 13
            Width = 165
            Height = 116
            Caption = #20135#21697#23376#31867
            TabOrder = 1
          end
        end
      end
      object GroupBox4: TGroupBox
        Left = 23
        Top = 16
        Width = 490
        Height = 129
        Caption = #28155#21152#22806#37096#25991#20214
        TabOrder = 1
        object leModeFile: TLabeledEdit
          Left = 88
          Top = 31
          Width = 345
          Height = 24
          EditLabel.Width = 80
          EditLabel.Height = 16
          EditLabel.Caption = #27169#26495#25991#20214#65306
          ImeName = #20013#25991'('#31616#20307') - '#30334#24230#36755#20837#27861
          LabelPosition = lpLeft
          TabOrder = 0
        end
        object lePicPath: TLabeledEdit
          Left = 88
          Top = 95
          Width = 345
          Height = 24
          EditLabel.Width = 80
          EditLabel.Height = 16
          EditLabel.Caption = #25551#36848#22270#29255#65306
          ImeName = #20013#25991'('#31616#20307') - '#30334#24230#36755#20837#27861
          LabelPosition = lpLeft
          TabOrder = 1
        end
        object leProdcutsFile: TLabeledEdit
          Left = 88
          Top = 61
          Width = 345
          Height = 24
          EditLabel.Width = 80
          EditLabel.Height = 16
          EditLabel.Caption = #20379#24212#37197#32622#65306
          ImeName = #20013#25991'('#31616#20307') - '#30334#24230#36755#20837#27861
          LabelPosition = lpLeft
          TabOrder = 2
        end
        object btnModeFile: TBitBtn
          Left = 439
          Top = 32
          Width = 34
          Height = 25
          Caption = '...'
          TabOrder = 3
          OnClick = btnModeFileClick
        end
        object btnPicPath: TBitBtn
          Left = 439
          Top = 95
          Width = 33
          Height = 24
          Caption = '...'
          TabOrder = 4
          OnClick = btnPicPathClick
        end
        object btnProductsFlie: TBitBtn
          Left = 439
          Top = 63
          Width = 34
          Height = 25
          Caption = '...'
          TabOrder = 5
          OnClick = btnProductsFlieClick
        end
      end
      object BitBtn2: TBitBtn
        Left = 391
        Top = 151
        Width = 128
        Height = 50
        Caption = #28120#23453#25968#25454#29983#25104
        TabOrder = 2
        OnClick = BitBtn2Click
      end
      object GroupBox6: TGroupBox
        Left = 3
        Top = 240
        Width = 774
        Height = 204
        Caption = #20379#24212#21830#20449#24687
        TabOrder = 3
        object edtSupplyName: TEdit
          Left = 16
          Top = 52
          Width = 129
          Height = 24
          ImeName = #20013#25991'('#31616#20307') - '#30334#24230#36755#20837#27861
          TabOrder = 0
          OnChange = edtSupplyNameChange
        end
        object bntAddSupple: TBitBtn
          Left = 288
          Top = 23
          Width = 131
          Height = 28
          Caption = #28155#21152#20379#36135#21830#20449#24687
          TabOrder = 1
          OnClick = bntAddSuppleClick
        end
        object cbSupplerName: TComboBox
          Left = 16
          Top = 23
          Width = 185
          Height = 24
          ImeName = #20013#25991'('#31616#20307') - '#30334#24230#36755#20837#27861
          TabOrder = 2
        end
        object vleSupplerInfo: TValueListEditor
          Left = 16
          Top = 83
          Width = 307
          Height = 178
          Strings.Strings = (
            '')
          TabOrder = 3
          ColWidths = (
            136
            165)
        end
        object Memo1: TMemo
          Left = 329
          Top = 80
          Width = 432
          Height = 151
          Margins.Left = 4
          Margins.Top = 4
          Margins.Right = 4
          Margins.Bottom = 4
          Lines.Strings = (
            'Memo1')
          ScrollBars = ssBoth
          TabOrder = 4
        end
      end
      object vleBaseProduceInfo: TValueListEditor
        Left = 23
        Top = 176
        Width = 303
        Height = 65
        FixedCols = 1
        Strings.Strings = (
          #36134#25143#21517#31216'=hmingyou'
          #23453#36125#20215#26684'=')
        TabOrder = 4
        OnSetEditText = vleBaseProduceInfo1SetEditText
        ColWidths = (
          150
          147)
        RowHeights = (
          18
          18
          18)
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'TabSheet3'
      ImageIndex = 2
      object Label1: TLabel
        Left = 43
        Top = 40
        Width = 64
        Height = 16
        Caption = #28120#23453#38142#25509
      end
      object edtTaoBaoUrl: TEdit
        Left = 105
        Top = 37
        Width = 446
        Height = 24
        TabOrder = 0
      end
      object Button2: TButton
        Left = 421
        Top = 67
        Width = 130
        Height = 46
        Caption = #36716#31227#21040#28120#23453
        TabOrder = 1
        OnClick = Button2Click
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'TabSheet4'
      ImageIndex = 3
      object Button1: TButton
        Left = 24
        Top = 128
        Width = 121
        Height = 33
        Caption = 'Button1'
        TabOrder = 0
        OnClick = Button1Click
      end
      object Button3: TButton
        Left = 272
        Top = 112
        Width = 145
        Height = 41
        Caption = 'Button3'
        TabOrder = 1
        OnClick = Button3Click
      end
      object Memo2: TMemo
        Left = 496
        Top = 88
        Width = 425
        Height = 209
        Lines.Strings = (
          'Memo2')
        TabOrder = 2
      end
      object DBGrid1: TDBGrid
        Left = 88
        Top = 288
        Width = 401
        Height = 169
        TabOrder = 3
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -13
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = [fsBold]
      end
    end
  end
  object IdHttpListPage: TIdHTTP
    IOHandler = IdSSLIOHandlerSocketOpenSSL1
    AllowCookies = True
    ProxyParams.BasicAuthentication = False
    ProxyParams.ProxyPort = 0
    Request.ContentLength = -1
    Request.ContentRangeEnd = -1
    Request.ContentRangeStart = -1
    Request.ContentRangeInstanceLength = -1
    Request.Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    Request.BasicAuthentication = False
    Request.UserAgent = 'Mozilla/3.0 (compatible; Indy Library)'
    Request.Ranges.Units = 'bytes'
    Request.Ranges = <>
    HTTPOptions = [hoForceEncodeParams]
    Left = 848
    Top = 24
  end
  object odFileBox: TOpenDialog
    Left = 696
    Top = 32
  end
  object IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL
    MaxLineAction = maException
    Port = 0
    DefaultPort = 0
    SSLOptions.Mode = sslmUnassigned
    SSLOptions.VerifyMode = []
    SSLOptions.VerifyDepth = 0
    Left = 972
    Top = 67
  end
end
