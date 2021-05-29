object Form2: TForm2
  Left = 0
  Top = 0
  Caption = #1071#1074#1082#1072' '#1085#1072' '#1088#1072#1073#1086#1090#1091
  ClientHeight = 339
  ClientWidth = 1112
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnActivate = FormActivate
  DesignSize = (
    1112
    339)
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 247
    Top = 14
    Width = 75
    Height = 25
    Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100
    TabOrder = 0
    OnClick = Button1Click
  end
  object ComboBox1: TComboBox
    Left = 16
    Top = 16
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemIndex = 0
    TabOrder = 1
    Text = #1071#1085#1074#1072#1088#1100
    Items.Strings = (
      #1071#1085#1074#1072#1088#1100
      #1060#1077#1074#1088#1072#1083#1100
      #1052#1072#1088#1090
      #1040#1087#1088#1077#1083#1100
      #1052#1072#1081
      #1048#1102#1085#1100
      #1040#1074#1075#1091#1089#1090
      #1057#1077#1085#1090#1103#1073#1088#1100
      #1054#1082#1090#1103#1073#1088#1100
      #1053#1086#1103#1073#1088#1100
      #1044#1077#1082#1072#1073#1088#1100)
  end
  object SpinEdit1: TSpinEdit
    Left = 176
    Top = 16
    Width = 65
    Height = 22
    MaxValue = 9999
    MinValue = 2020
    TabOrder = 2
    Value = 2021
  end
  object StringGrid1: TStringGrid
    Left = 8
    Top = 120
    Width = 1096
    Height = 177
    Anchors = [akLeft, akTop, akRight]
    ColCount = 32
    DefaultColWidth = 40
    DefaultColAlignment = taRightJustify
    DrawingStyle = gdsGradient
    FixedColor = clSilver
    RowCount = 6
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goFixedRowDefAlign]
    TabOrder = 3
    OnDrawCell = StringGrid1DrawCell
  end
  object ComboBox2: TComboBox
    Left = 16
    Top = 56
    Width = 169
    Height = 21
    Style = csDropDownList
    ItemIndex = 0
    TabOrder = 4
    Text = #1057#1040#1056#1040#1053#1057#1050#1050#1040#1041#1045#1051#1068'-'#1054#1055#1058#1048#1050#1040
    OnChange = ComboBox2Change
    Items.Strings = (
      #1057#1040#1056#1040#1053#1057#1050#1050#1040#1041#1045#1051#1068'-'#1054#1055#1058#1048#1050#1040
      #1069#1052'-'#1050#1040#1041#1045#1051#1068
      #1069#1052'-'#1050#1040#1058
      #1069#1052'-'#1055#1051#1040#1057#1058)
  end
  object ComboBox3: TComboBox
    Left = 200
    Top = 56
    Width = 177
    Height = 21
    Style = csDropDownList
    TabOrder = 5
  end
  object Button2: TButton
    Left = 432
    Top = 303
    Width = 75
    Height = 25
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
    TabOrder = 6
    OnClick = Button2Click
  end
  object IdHTTP1: TIdHTTP
    HandleRedirects = True
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
    Left = 832
    Top = 16
  end
end
