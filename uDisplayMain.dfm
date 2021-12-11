object fmDisplayMain: TfmDisplayMain
  Left = 0
  Top = 0
  Caption = 'Main '#54868#47732
  ClientHeight = 742
  ClientWidth = 1095
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #47569#51008' '#44256#46357
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 17
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 1095
    Height = 742
    ActivePage = TabSheet2
    Align = alClient
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = '  '#48120#49324' '#52280#49437' '#51077#47141' (QR Code)  '
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object sgData: TStringGrid
        Left = 0
        Top = 145
        Width = 1087
        Height = 565
        Align = alClient
        ColCount = 7
        RowCount = 1
        FixedRows = 0
        TabOrder = 0
        RowHeights = (
          29)
      end
      object pnDataInsert: TPanel
        Left = 0
        Top = 0
        Width = 1087
        Height = 145
        Align = alTop
        BevelInner = bvLowered
        ParentBackground = False
        TabOrder = 1
        object Label1: TLabel
          Left = 23
          Top = 14
          Width = 26
          Height = 17
          Caption = #49457#47749
        end
        object Label2: TLabel
          Left = 174
          Top = 14
          Width = 39
          Height = 17
          Caption = #49464#47168#47749
        end
        object Label3: TLabel
          Left = 336
          Top = 14
          Width = 39
          Height = 17
          Caption = #50672#46973#52376
        end
        object Label4: TLabel
          Left = 23
          Top = 54
          Width = 26
          Height = 17
          Caption = #51452#49548
        end
        object Label5: TLabel
          Left = 514
          Top = 14
          Width = 26
          Height = 17
          Caption = #49457#48324
        end
        object Label18: TLabel
          Left = 23
          Top = 90
          Width = 692
          Height = 17
          Caption = 
            #8251' "'#48120#49324#52280#47168' '#46321#47197' '#48520#47049'"'#51060' '#48156#49373#54624' '#44221#50864' '#54532#47196#44536#47016' '#51333#47308#54980' '#51116#49892#54665' '#46608#45716' '#44288#47532#51088#50640#44172' '#51204#54868' '#47928#51032' '#48148#46989#45768#45796' (010-478' +
            '7-3013)'
        end
        object lbTtsStatus: TLabel
          Left = 1048
          Top = 120
          Width = 21
          Height = 17
          Alignment = taRightJustify
          Caption = 'TTS'
        end
        object lbDataStatus: TLabel
          Left = 23
          Top = 122
          Width = 26
          Height = 17
          Caption = #49345#53468
        end
        object edName: TEdit
          Left = 60
          Top = 10
          Width = 97
          Height = 25
          Enabled = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = #47569#51008' '#44256#46357
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
        end
        object edbaptismalName: TEdit
          Left = 220
          Top = 10
          Width = 97
          Height = 25
          Enabled = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = #47569#51008' '#44256#46357
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
        end
        object edTelNo: TEdit
          Left = 380
          Top = 10
          Width = 113
          Height = 25
          Enabled = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = #47569#51008' '#44256#46357
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
        end
        object edAddress: TEdit
          Left = 60
          Top = 50
          Width = 657
          Height = 25
          Enabled = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = #47569#51008' '#44256#46357
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 3
        end
        object edSex: TEdit
          Left = 546
          Top = 10
          Width = 65
          Height = 25
          Enabled = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = #47569#51008' '#44256#46357
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 4
        end
        object Panel2: TPanel
          Left = 842
          Top = 2
          Width = 243
          Height = 141
          Align = alRight
          BevelOuter = bvNone
          TabOrder = 5
          object Panel3: TPanel
            Left = 0
            Top = 84
            Width = 243
            Height = 28
            Align = alTop
            Caption = #48120#49324' '#52280#47168' '#46321#47197' '#48520#47049'('#51216#44160' '#54596#50836')'
            Color = clRed
            ParentBackground = False
            TabOrder = 0
          end
          object Panel4: TPanel
            Left = 0
            Top = 28
            Width = 243
            Height = 28
            Align = alTop
            Caption = #48120#49324#52280#47168' '#51221#49345' '#46321#47197
            Color = clLime
            ParentBackground = False
            TabOrder = 1
          end
          object Panel5: TPanel
            Left = 0
            Top = 56
            Width = 243
            Height = 28
            Align = alTop
            Caption = '1'#49884#44036#51060#45236' '#51060#48120' '#46321#47197#46120'('#48120#46321#47197' '#52376#47532')'
            Color = clTeal
            ParentBackground = False
            TabOrder = 2
          end
          object Panel6: TPanel
            Left = 0
            Top = 0
            Width = 243
            Height = 28
            Align = alTop
            Caption = #48120#49324' '#52280#47168' '#49888#51088' '#46321#47197' '#45824#44592
            ParentBackground = False
            TabOrder = 3
          end
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = '  '#48120#49324' '#52280#49437' '#49688#51089#50629' '#51077#47141'  '
      ImageIndex = 1
      object pnQueryInsert: TPanel
        Left = 0
        Top = 0
        Width = 1087
        Height = 137
        Align = alTop
        BevelInner = bvLowered
        ParentBackground = False
        TabOrder = 0
        object Label6: TLabel
          Left = 23
          Top = 14
          Width = 26
          Height = 17
          Caption = #49457#47749
        end
        object Label7: TLabel
          Left = 174
          Top = 14
          Width = 39
          Height = 17
          Caption = #49464#47168#47749
        end
        object Label8: TLabel
          Left = 336
          Top = 14
          Width = 39
          Height = 17
          Caption = #50672#46973#52376
        end
        object Label9: TLabel
          Left = 23
          Top = 54
          Width = 26
          Height = 17
          Caption = #51452#49548
        end
        object Label10: TLabel
          Left = 514
          Top = 14
          Width = 26
          Height = 17
          Caption = #49457#48324
        end
        object Label12: TLabel
          Left = 23
          Top = 112
          Width = 381
          Height = 17
          Caption = #9632' '#44033' '#54637#47785#51012' '#49688#46041#51004#47196' '#51077#47141#54980' "'#49688#46041' '#46321#47197'" '#48260#53948#51012' Click '#54616#49901#49884#50724
        end
        object Label11: TLabel
          Left = 21
          Top = 85
          Width = 67
          Height = 17
          Caption = #44368#44396' ID No'
        end
        object edQueryName: TEdit
          Left = 60
          Top = 10
          Width = 97
          Height = 25
          TabOrder = 0
        end
        object edQueryBaptismalName: TEdit
          Left = 220
          Top = 10
          Width = 97
          Height = 25
          TabOrder = 1
        end
        object edQueryTelNo: TEdit
          Left = 380
          Top = 10
          Width = 113
          Height = 25
          TabOrder = 2
        end
        object edQueryAddress: TEdit
          Left = 60
          Top = 50
          Width = 614
          Height = 25
          TabOrder = 3
        end
        object btQuery: TButton
          Left = 960
          Top = 130
          Width = 97
          Height = 25
          Caption = #51312' '#54924
          TabOrder = 4
          Visible = False
          OnClick = btQueryClick
        end
        object cbQuerySex: TComboBox
          Left = 546
          Top = 10
          Width = 126
          Height = 25
          TabOrder = 5
          Items.Strings = (
            #54805#51228
            #51088#47588)
        end
        object btManualReg: TButton
          Left = 704
          Top = 14
          Width = 105
          Height = 61
          Caption = #49688#46041' '#46321#47197
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlue
          Font.Height = -16
          Font.Name = #47569#51008' '#44256#46357
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 6
          OnClick = btManualRegClick
        end
        object Panel7: TPanel
          Left = 842
          Top = 2
          Width = 243
          Height = 133
          Align = alRight
          BevelOuter = bvNone
          TabOrder = 7
          object Panel8: TPanel
            Left = 0
            Top = 84
            Width = 243
            Height = 28
            Align = alTop
            Caption = #48120#49324' '#52280#47168' '#46321#47197' '#48520#47049'('#51216#44160' '#54596#50836')'
            Color = clRed
            ParentBackground = False
            TabOrder = 0
          end
          object Panel9: TPanel
            Left = 0
            Top = 28
            Width = 243
            Height = 28
            Align = alTop
            Caption = #48120#49324#52280#47168' '#51221#49345' '#46321#47197
            Color = clLime
            ParentBackground = False
            TabOrder = 1
          end
          object Panel10: TPanel
            Left = 0
            Top = 56
            Width = 243
            Height = 28
            Align = alTop
            Caption = '1'#49884#44036#51060#45236' '#51060#48120' '#46321#47197#46120'('#48120#46321#47197' '#52376#47532')'
            Color = clTeal
            ParentBackground = False
            TabOrder = 2
          end
          object Panel11: TPanel
            Left = 0
            Top = 0
            Width = 243
            Height = 28
            Align = alTop
            Caption = #48120#49324' '#52280#47168' '#49888#51088' '#46321#47197' '#45824#44592
            ParentBackground = False
            TabOrder = 3
          end
        end
        object edExtData: TEdit
          Left = 110
          Top = 81
          Width = 564
          Height = 25
          TabOrder = 8
        end
      end
      object ckQueryList: TCheckListBox
        Left = 0
        Top = 137
        Width = 1087
        Height = 392
        Align = alClient
        ItemHeight = 17
        TabOrder = 1
        OnClick = ckQueryListClick
      end
      object moQueryStatus: TMemo
        Left = 0
        Top = 529
        Width = 1087
        Height = 181
        Align = alBottom
        ScrollBars = ssBoth
        TabOrder = 2
        WordWrap = False
      end
    end
  end
  object ComPort1: TComPort
    BaudRate = br9600
    Port = 'COM1'
    Parity.Bits = prNone
    StopBits = sbOneStopBit
    DataBits = dbEight
    Events = [evRxChar, evTxEmpty, evRxFlag, evRing, evBreak, evCTS, evDSR, evError, evRLSD, evRx80Full]
    FlowControl.OutCTSFlow = False
    FlowControl.OutDSRFlow = False
    FlowControl.ControlDTR = dtrDisable
    FlowControl.ControlRTS = rtsDisable
    FlowControl.XonXoffOut = False
    FlowControl.XonXoffIn = False
    StoredProps = [spBasic]
    TriggersOnRxChar = True
    OnRxChar = ComPort1RxChar
    Left = 748
  end
  object Timer1: TTimer
    Enabled = False
    Interval = 100
    OnTimer = Timer1Timer
    Left = 708
    Top = 65532
  end
  object tmCheckDb: TTimer
    Enabled = False
    Interval = 3600000
    OnTimer = tmCheckDbTimer
    Left = 820
    Top = 4
  end
end
