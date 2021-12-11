object fmReport: TfmReport
  Left = 0
  Top = 0
  Caption = 'fmReport'
  ClientHeight = 554
  ClientWidth = 1079
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #47569#51008' '#44256#46357
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 17
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 1079
    Height = 554
    ActivePage = TabSheet2
    Align = alClient
    TabOrder = 0
    ExplicitHeight = 625
    object TabSheet2: TTabSheet
      Caption = '  '#51068#51068' / '#44592#44036' '#52636#49437' '#51312#54924
      ImageIndex = 1
      ExplicitHeight = 593
      DesignSize = (
        1071
        522)
      object Label14: TLabel
        Left = 8
        Top = 20
        Width = 57
        Height = 17
        Caption = #49884#51089' '#45216#51088
      end
      object Label15: TLabel
        Left = 219
        Top = 20
        Width = 57
        Height = 17
        Caption = #51333#47308' '#45216#51088
      end
      object lbReportStatus: TLabel
        Left = 15
        Top = 110
        Width = 87
        Height = 17
        Anchors = [akTop]
        Caption = 'lbReportStatus'
      end
      object dtReportStartTime: TDateTimePicker
        Left = 79
        Top = 16
        Width = 121
        Height = 25
        Date = 43949.909724907400000000
        Time = 43949.909724907400000000
        TabOrder = 0
      end
      object dtReportEndTime: TDateTimePicker
        Left = 288
        Top = 16
        Width = 121
        Height = 25
        Date = 43949.909724907400000000
        Time = 43949.909724907400000000
        TabOrder = 1
      end
      object moReport: TMemo
        Left = 0
        Top = 133
        Width = 1071
        Height = 389
        Align = alBottom
        Anchors = [akBottom]
        ScrollBars = ssBoth
        TabOrder = 2
        WordWrap = False
      end
      object Button4: TButton
        Left = 232
        Top = 59
        Width = 177
        Height = 35
        Caption = #51068#51068' '#47112#54252#53944' '#51200#51109
        TabOrder = 3
        Visible = False
        OnClick = Button4Click
      end
      object btReportSave: TButton
        Left = 23
        Top = 59
        Width = 177
        Height = 35
        Caption = #44592#44036' '#47112#54252#53944' '#51200#51109
        TabOrder = 4
        OnClick = btReportSaveClick
      end
      object Button2: TButton
        Left = 888
        Top = 104
        Width = 75
        Height = 25
        Caption = 'Data'#51204#49569
        TabOrder = 5
        Visible = False
        OnClick = Button2Click
      end
    end
    object TabSheet4: TTabSheet
      Caption = #44368#44396' '#44368#51064' '#51221#48372' '#46321#47197
      ImageIndex = 3
      ExplicitHeight = 593
      object Panel1: TPanel
        Left = 0
        Top = 73
        Width = 1071
        Height = 449
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 0
        ExplicitHeight = 520
        object Splitter1: TSplitter
          Left = 0
          Top = 241
          Width = 1071
          Height = 8
          Cursor = crVSplit
          Align = alBottom
          ExplicitTop = 249
          ExplicitWidth = 969
        end
        object pnRegErrorList: TPanel
          Left = 0
          Top = 249
          Width = 1071
          Height = 200
          Align = alBottom
          BevelInner = bvLowered
          ParentBackground = False
          TabOrder = 0
          ExplicitTop = 320
          object Label19: TLabel
            Left = 8
            Top = 9
            Width = 88
            Height = 17
            Caption = #46321#47197' '#49892#54224' '#51221#48372
          end
          object moError: TMemo
            Left = 2
            Top = 32
            Width = 1067
            Height = 166
            Align = alBottom
            ScrollBars = ssBoth
            TabOrder = 0
            WordWrap = False
          end
        end
        object pnRegistration: TPanel
          Left = 0
          Top = 0
          Width = 1071
          Height = 241
          Align = alClient
          BevelInner = bvLowered
          ParentBackground = False
          TabOrder = 1
          ExplicitHeight = 312
          object Label20: TLabel
            Left = 8
            Top = 15
            Width = 88
            Height = 17
            Caption = #46321#47197' '#44368#51064' '#51221#48372
          end
          object lbRegStatus: TLabel
            Left = 160
            Top = 17
            Width = 36
            Height = 17
            Caption = 'Status'
          end
          object moRegData: TMemo
            Left = 2
            Top = -64
            Width = 1067
            Height = 303
            Align = alBottom
            ScrollBars = ssBoth
            TabOrder = 0
            WordWrap = False
            ExplicitTop = 7
          end
        end
      end
      object Panel2: TPanel
        Left = 0
        Top = 0
        Width = 1071
        Height = 73
        Align = alTop
        BevelInner = bvLowered
        TabOrder = 1
        object Label4: TLabel
          Left = 256
          Top = 32
          Width = 469
          Height = 17
          Caption = #45236#50857' : "'#44368#44396' ID", "'#49457#47749'", "'#49464#47168#47749'", "'#51204#54868#48264#54840'", "'#49373#45380#50900#51068'", "'#49457#48324'", "Level", "'#51452#49548'"'
        end
        object btReg: TButton
          Left = 19
          Top = 13
          Width = 201
          Height = 49
          Caption = #51204#51077' '#44368#51064' '#51221#48372' '#52628#44032' / '#49688#51221
          TabOrder = 0
          OnClick = btRegClick
        end
        object Button1: TButton
          Left = 976
          Top = 25
          Width = 75
          Height = 25
          Caption = 'Delete'
          TabOrder = 1
          OnClick = Button1Click
        end
      end
    end
    object TabSheet1: TTabSheet
      Caption = #44368#44396' '#44368#51064' '#51221#48372' '#49688#51221
      ImageIndex = 2
      ExplicitHeight = 598
      object pnModify: TPanel
        Left = 0
        Top = 41
        Width = 1071
        Height = 248
        Align = alTop
        BevelInner = bvLowered
        ParentBackground = False
        TabOrder = 0
        object Label1: TLabel
          Left = 16
          Top = 24
          Width = 67
          Height = 17
          Caption = #44368#44396' ID No'
        end
        object Label6: TLabel
          Left = 23
          Top = 102
          Width = 26
          Height = 17
          Caption = #49457#47749
        end
        object Label7: TLabel
          Left = 174
          Top = 102
          Width = 39
          Height = 17
          Caption = #49464#47168#47749
        end
        object Label8: TLabel
          Left = 384
          Top = 102
          Width = 52
          Height = 17
          Caption = #51204#54868#48264#54840
        end
        object Label10: TLabel
          Left = 576
          Top = 102
          Width = 26
          Height = 17
          Caption = #49457#48324
        end
        object Label2: TLabel
          Left = 23
          Top = 142
          Width = 29
          Height = 17
          Caption = 'Level'
        end
        object Label3: TLabel
          Left = 174
          Top = 142
          Width = 52
          Height = 17
          Caption = #49373#45380#50900#51068
        end
        object Label9: TLabel
          Left = 23
          Top = 190
          Width = 26
          Height = 17
          Caption = #51452#49548
        end
        object edIdNo: TEdit
          Left = 89
          Top = 21
          Width = 265
          Height = 25
          TabOrder = 0
          OnKeyPress = edIdNoKeyPress
        end
        object btView: TButton
          Left = 384
          Top = 16
          Width = 75
          Height = 36
          Caption = #51312#54924
          TabOrder = 1
          OnClick = btViewClick
        end
        object btMember: TButton
          Left = 606
          Top = 16
          Width = 130
          Height = 36
          Caption = #49888#44508' '#46321#47197' / '#49688#51221
          TabOrder = 2
          OnClick = btMemberClick
        end
        object edQueryName: TEdit
          Left = 60
          Top = 98
          Width = 95
          Height = 25
          TabOrder = 3
        end
        object edQueryBaptismalName: TEdit
          Left = 231
          Top = 98
          Width = 120
          Height = 25
          TabOrder = 4
        end
        object edQueryTelNo: TEdit
          Left = 442
          Top = 98
          Width = 113
          Height = 25
          TabOrder = 5
        end
        object cbQuerySex: TComboBox
          Left = 616
          Top = 98
          Width = 120
          Height = 25
          TabOrder = 6
          Items.Strings = (
            #45224
            #50668)
        end
        object edQueryLevel: TEdit
          Left = 60
          Top = 138
          Width = 95
          Height = 25
          TabOrder = 7
        end
        object edQueryBirthDay: TEdit
          Left = 231
          Top = 138
          Width = 120
          Height = 25
          TabOrder = 8
        end
        object edQueryAddress: TEdit
          Left = 60
          Top = 186
          Width = 676
          Height = 25
          TabOrder = 9
        end
      end
      object Panel3: TPanel
        Left = 0
        Top = 0
        Width = 1071
        Height = 41
        Align = alTop
        BevelInner = bvLowered
        TabOrder = 1
        object lbModify: TLabel
          Left = 24
          Top = 15
          Width = 9
          Height = 17
          Caption = '...'
        end
      end
    end
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = '.xlsx'
    Filter = #50641#49472#54028#51068'|*.xlsx'
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 780
    Top = 12
  end
  object OpenDialog2: TOpenDialog
    DefaultExt = '.log'
    Filter = #47196#44536#54028#51068'|*.log'
    Left = 828
    Top = 12
  end
end
