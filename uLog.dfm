object fmLog: TfmLog
  Left = 0
  Top = 0
  BorderIcons = []
  Caption = 'Log'
  ClientHeight = 535
  ClientWidth = 941
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 494
    Width = 941
    Height = 41
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    DesignSize = (
      941
      41)
    object Button2: TButton
      Left = 840
      Top = 6
      Width = 75
      Height = 25
      Anchors = [akTop, akRight]
      Caption = 'Log Clear'
      TabOrder = 0
      OnClick = Button2Click
    end
  end
  object Memo1: TMemo
    Left = 0
    Top = 456
    Width = 941
    Height = 38
    Align = alBottom
    ScrollBars = ssVertical
    TabOrder = 1
    WordWrap = False
  end
  object sgLogInfo: TStringGrid
    Left = 0
    Top = 0
    Width = 941
    Height = 456
    Align = alClient
    ColCount = 7
    DefaultRowHeight = 20
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
    TabOrder = 2
    OnDrawCell = sgLogInfoDrawCell
  end
  object JvThLog: TJvThread
    Exclusive = True
    MaxCount = 0
    RunOnCreate = True
    FreeOnTerminate = True
    OnExecute = JvThLogExecute
    Left = 728
    Top = 32
  end
end
