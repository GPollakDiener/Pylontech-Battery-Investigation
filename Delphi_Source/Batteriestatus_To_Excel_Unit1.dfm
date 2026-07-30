object Form1: TForm1
  Left = 408
  Top = 303
  Width = 490
  Height = 150
  Caption = 'Log Pylontech Batteries'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 48
    Width = 45
    Height = 13
    Caption = 'No Name'
  end
  object Label2: TLabel
    Left = 208
    Top = 96
    Width = 54
    Height = 13
    Caption = 'Zeitintervall'
  end
  object Label3: TLabel
    Left = 320
    Top = 96
    Width = 77
    Height = 13
    Caption = 'Anzahl Batterien'
  end
  object Label4: TLabel
    Left = 320
    Top = 40
    Width = 46
    Height = 13
    Caption = 'COM-Port'
  end
  object Start_Button: TButton
    Left = 16
    Top = 72
    Width = 73
    Height = 25
    Caption = 'Start'
    TabOrder = 0
    OnClick = Start_ButtonClick
  end
  object Ende: TButton
    Left = 208
    Top = 16
    Width = 75
    Height = 25
    Caption = 'Ende'
    TabOrder = 1
    OnClick = EndeClick
  end
  object Stop_Button1: TButton
    Left = 112
    Top = 72
    Width = 73
    Height = 25
    Caption = 'Stopp'
    TabOrder = 2
    OnClick = Stop_Button1Click
  end
  object Datei_Waehlen_Button: TButton
    Left = 16
    Top = 16
    Width = 73
    Height = 25
    Caption = 'Dateiname'
    TabOrder = 3
    OnClick = Datei_Waehlen_ButtonClick
  end
  object Speichern_Button: TButton
    Left = 112
    Top = 16
    Width = 75
    Height = 25
    Caption = 'Speichern'
    TabOrder = 4
    OnClick = Speichern_ButtonClick
  end
  object IntervallEdit: TEdit
    Left = 208
    Top = 72
    Width = 73
    Height = 21
    TabOrder = 5
    Text = '20'
  end
  object Anzahl_Bat_Edit: TEdit
    Left = 320
    Top = 72
    Width = 73
    Height = 21
    TabOrder = 6
    Text = '16'
  end
  object COM_Port_Edit: TEdit
    Left = 320
    Top = 16
    Width = 73
    Height = 21
    TabOrder = 7
    Text = '1'
  end
  object Timer1: TTimer
    Interval = 20000
    OnTimer = Timer1Timer
    Left = 416
    Top = 72
  end
  object SaveDialog1: TSaveDialog
    Left = 448
    Top = 72
  end
end
