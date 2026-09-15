object Form1: TForm1
  Left = 192
  Top = 114
  Width = 273
  Height = 259
  Caption = 'BASS plugin test'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 176
    Width = 249
    Height = 17
    Alignment = taCenter
    AutoSize = False
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 249
    Height = 129
    Caption = 'Loaded plugins'
    TabOrder = 0
    object ListBox1: TListBox
      Left = 8
      Top = 16
      Width = 233
      Height = 105
      ItemHeight = 13
      TabOrder = 0
    end
  end
  object Button1: TButton
    Left = 8
    Top = 144
    Width = 249
    Height = 25
    Caption = 'Click here to open a file...'
    TabOrder = 1
    OnClick = Button1Click
  end
  object TrackBar1: TTrackBar
    Left = 8
    Top = 200
    Width = 249
    Height = 17
    TabOrder = 2
    ThumbLength = 10
    TickMarks = tmBoth
    TickStyle = tsNone
    OnChange = TrackBar1Change
  end
  object Timer1: TTimer
    Interval = 500
    OnTimer = Timer1Timer
  end
  object OpenDialog1: TOpenDialog
    Filter = 'All files|*.*'
    Left = 32
  end
end
