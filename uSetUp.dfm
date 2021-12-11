object fmSetUp: TfmSetUp
  Left = 0
  Top = 0
  Caption = '`'
  ClientHeight = 555
  ClientWidth = 977
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #47569#51008' '#44256#46357
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 17
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 977
    Height = 555
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = 'System SetUp'
      object Panel1: TPanel
        Left = 0
        Top = 482
        Width = 969
        Height = 41
        Align = alBottom
        BevelInner = bvLowered
        TabOrder = 0
        DesignSize = (
          969
          41)
        object Button1: TButton
          Left = 872
          Top = 8
          Width = 75
          Height = 25
          Anchors = [akTop, akRight]
          Caption = 'Save'
          TabOrder = 0
          OnClick = Button1Click
        end
      end
      object zObjSystem: TzObjectInspector
        Left = 0
        Top = 0
        Width = 969
        Height = 482
        Align = alClient
        Color = clWhite
        BorderStyle = bsSingle
        Component = zObjSystem
        TabOrder = 1
        AllowSearch = True
        AutoCompleteText = True
        DefaultCategoryName = 'Miscellaneous'
        ShowGutter = True
        GutterColor = clCream
        GutterEdgeColor = clGray
        NameColor = clBtnText
        ValueColor = clNavy
        NonDefaultValueColor = clNavy
        BoldNonDefaultValue = True
        HighlightColor = 14737632
        ReferencesColor = clMaroon
        SubPropertiesColor = clGreen
        ShowHeader = False
        ShowGridLines = False
        GridColor = clBlack
        SplitterColor = clGray
        ReadOnlyColor = clGrayText
        FixedSplitter = False
        ReadOnly = False
        TrackChange = False
        GutterWidth = 12
        ShowItemHint = True
        SortByCategory = False
        SplitterPos = 100
        HeaderPropText = 'Property'
        HeaderValueText = 'Value'
        FloatPreference.MaxDigits = 2
        FloatPreference.ExpPrecision = 6
      end
    end
  end
end
