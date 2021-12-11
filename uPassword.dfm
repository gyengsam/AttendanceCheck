object fmPassword: TfmPassword
  Left = 0
  Top = 0
  Caption = #44288#47532#51088' '#50516#54840' '#54869#51064
  ClientHeight = 225
  ClientWidth = 584
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #47569#51008' '#44256#46357
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 17
  object Label1: TLabel
    Left = 48
    Top = 59
    Width = 96
    Height = 17
    Caption = #44288#47532#51088' '#48708#48128#48264#54840
  end
  object Label18: TLabel
    Left = -57
    Top = 90
    Width = 15
    Height = 17
    Caption = #8251' '
  end
  object Label2: TLabel
    Left = 48
    Top = 112
    Width = 356
    Height = 17
    Caption = #8251' Password '#51077#47141#54980' "ENTER" '#46608#45716' "'#54869#51064'" '#48260#53948#51012' '#45572#47476#49901#49884#50724
  end
  object edPassword: TEdit
    Left = 168
    Top = 56
    Width = 201
    Height = 25
    PasswordChar = '*'
    TabOrder = 0
    OnKeyPress = edPasswordKeyPress
  end
  object Button1: TButton
    Left = 392
    Top = 56
    Width = 75
    Height = 25
    Caption = #54869#51064
    TabOrder = 1
    OnClick = Button1Click
  end
end
