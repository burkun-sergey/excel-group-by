object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = #1055#1088#1077#1086#1073#1088#1072#1079#1086#1074#1072#1085#1080#1077' XLS'
  ClientHeight = 280
  ClientWidth = 636
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 24
    Top = 8
    Width = 80
    Height = 13
    Caption = #1048#1089#1093#1086#1076#1085#1099#1081' '#1092#1072#1081#1083
  end
  object Label2: TLabel
    Left = 24
    Top = 48
    Width = 118
    Height = 13
    Caption = #1056#1077#1079#1091#1083#1100#1090#1080#1088#1091#1102#1097#1080#1081' '#1092#1072#1081#1083
  end
  object edSrc: TEdit
    Left = 24
    Top = 24
    Width = 273
    Height = 21
    TabOrder = 0
  end
  object bSrcFilenameBrowse: TButton
    Left = 303
    Top = 22
    Width = 33
    Height = 25
    Caption = '...'
    TabOrder = 1
    OnClick = bSrcFilenameBrowseClick
  end
  object edNew: TEdit
    Left = 24
    Top = 64
    Width = 273
    Height = 21
    TabOrder = 2
  end
  object bNewFilenameBrowse: TButton
    Left = 303
    Top = 62
    Width = 33
    Height = 25
    Caption = '...'
    TabOrder = 3
    OnClick = bNewFilenameBrowseClick
  end
  object GroupBox1: TGroupBox
    Left = 24
    Top = 136
    Width = 313
    Height = 137
    Caption = ' '#1048#1085#1092'. '#1086#1073' '#1086#1087#1077#1088#1072#1094#1080#1080' '
    TabOrder = 4
    object Label7: TLabel
      Left = 16
      Top = 24
      Width = 98
      Height = 13
      Caption = #1054#1073#1088#1072#1073#1086#1090#1082#1072' '#1089#1090#1088#1086#1082#1080':'
    end
    object Label8: TLabel
      Left = 16
      Top = 43
      Width = 52
      Height = 13
      Caption = #1057#1082#1086#1088#1086#1089#1090#1100':'
    end
    object Label9: TLabel
      Left = 16
      Top = 62
      Width = 43
      Height = 13
      Caption = #1055#1088#1086#1096#1083#1086':'
    end
    object Label10: TLabel
      Left = 16
      Top = 81
      Width = 52
      Height = 13
      Caption = #1054#1089#1090#1072#1083#1086#1089#1100':'
    end
    object lInfRowNum: TLabel
      Left = 128
      Top = 24
      Width = 58
      Height = 13
      Caption = 'lInfRowNum'
    end
    object lInfOPS: TLabel
      Left = 128
      Top = 43
      Width = 36
      Height = 13
      Caption = 'lInfOPS'
    end
    object lInfPassed: TLabel
      Left = 128
      Top = 62
      Width = 50
      Height = 13
      Caption = 'lInfPassed'
    end
    object lInfRemain: TLabel
      Left = 128
      Top = 81
      Width = 51
      Height = 13
      Caption = 'lInfRemain'
    end
  end
  object GroupBox2: TGroupBox
    Left = 352
    Top = 8
    Width = 273
    Height = 265
    Caption = ' '#1053#1072#1089#1090#1088#1086#1081#1082#1080' '#1080#1089#1093#1086#1076#1085#1099#1093' '#1076#1072#1085#1085#1099#1093' '
    TabOrder = 5
    object Label3: TLabel
      Left = 16
      Top = 24
      Width = 80
      Height = 13
      Caption = #1051#1080#1089#1090' '#1089' '#1076#1072#1085#1085#1099#1084#1080
    end
    object Label15: TLabel
      Left = 16
      Top = 56
      Width = 85
      Height = 13
      Caption = #1054#1073#1083#1072#1089#1090#1100' '#1076#1072#1085#1085#1099#1093
    end
    object Label16: TLabel
      Left = 16
      Top = 102
      Width = 75
      Height = 13
      Caption = #1040#1076#1088#1077#1089' '#1086#1073#1083#1072#1089#1090#1080
    end
    object seOptionSrcSheetIndex: TSpinEdit
      Left = 104
      Top = 21
      Width = 60
      Height = 22
      MaxValue = 0
      MinValue = 0
      TabOrder = 0
      Value = 1
    end
    object cbOptionSrcRange: TComboBox
      Left = 16
      Top = 72
      Width = 241
      Height = 21
      Style = csDropDownList
      ItemIndex = 0
      TabOrder = 1
      Text = #1059#1082#1072#1079#1072#1085#1085#1072#1103
      OnSelect = cbOptionSrcRangeSelect
      Items.Strings = (
        #1059#1082#1072#1079#1072#1085#1085#1072#1103
        #1040#1074#1090#1086#1084#1072#1090#1080#1095#1077#1089#1082#1080)
    end
    object edOptionSrcRangeAddress: TEdit
      Left = 104
      Top = 99
      Width = 153
      Height = 21
      Hint = #1044#1083#1103' '#1088#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1103' '#1074#1099#1073#1077#1088#1080#1090#1077' '#1086#1073#1083#1072#1089#1090#1100' '#1076#1072#1085#1085#1099#1093' - "'#1091#1082#1072#1079#1072#1085#1085#1072#1103'"'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      Text = 'A1:B100'
    end
    object GroupBox3: TGroupBox
      Left = 16
      Top = 128
      Width = 241
      Height = 121
      Caption = ' '#1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1086#1073#1083#1072#1089#1090#1080' '
      TabOrder = 3
      object Label4: TLabel
        Left = 16
        Top = 28
        Width = 89
        Height = 13
        Caption = #1053#1072#1095#1072#1090#1100' '#1089#1086' '#1089#1090#1088#1086#1082#1080
      end
      object Label5: TLabel
        Left = 16
        Top = 56
        Width = 133
        Height = 13
        Caption = #1050#1086#1083#1086#1085#1082#1072' '#1080#1076#1077#1085#1090#1080#1092#1080#1082#1072#1090#1086#1088#1072
      end
      object Label6: TLabel
        Left = 16
        Top = 84
        Width = 103
        Height = 13
        Caption = #1050#1086#1083#1086#1085#1082#1072' '#1096#1090#1088#1080#1093#1082#1086#1076#1072
      end
      object seOptionSrcStartRowNum: TSpinEdit
        Left = 160
        Top = 25
        Width = 65
        Height = 22
        MaxValue = 0
        MinValue = 0
        TabOrder = 0
        Value = 1
      end
      object seOptionSrcColNumId: TSpinEdit
        Left = 160
        Top = 53
        Width = 65
        Height = 22
        MaxValue = 0
        MinValue = 0
        TabOrder = 1
        Value = 1
      end
      object seOptionSrcColNumBarcode: TSpinEdit
        Left = 160
        Top = 81
        Width = 65
        Height = 22
        MaxValue = 0
        MinValue = 0
        TabOrder = 2
        Value = 2
      end
    end
  end
  object bStartStop: TButton
    Left = 24
    Top = 96
    Width = 185
    Height = 25
    Caption = #1053#1072#1095#1072#1090#1100' '#1087#1088#1077#1086#1073#1088#1072#1079#1086#1074#1072#1085#1080#1077
    TabOrder = 6
    OnClick = bStartStopClick
  end
  object pBar: TProgressBar
    Left = 216
    Top = 100
    Width = 121
    Height = 17
    TabOrder = 7
  end
  object OpenDialogXLS: TOpenDialog
    Filter = #1060#1072#1081#1083#1099' Excel|*.xls;*.xlsx'
    Left = 248
    Top = 8
  end
  object SaveDialogXLS: TSaveDialog
    Filter = #1060#1072#1081#1083#1099' Excel|*.xls;*.xlsx'
    Left = 248
    Top = 56
  end
  object TimerViewInf: TTimer
    Enabled = False
    Interval = 500
    OnTimer = TimerViewInfTimer
    Left = 280
    Top = 152
  end
end
