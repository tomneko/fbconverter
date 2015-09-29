object frmLogViewer: TfrmLogViewer
  Left = 0
  Top = 0
  Caption = 'Metadata Viewer'
  ClientHeight = 442
  ClientWidth = 624
  Color = clBtnFace
  Constraints.MinHeight = 100
  DefaultMonitor = dmMainForm
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDefault
  Scaled = False
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 290
    Width = 624
    Height = 3
    Cursor = crVSplit
    Align = alBottom
    ExplicitLeft = 28
    ExplicitTop = 267
  end
  object mmMessage: TMemo
    Left = 0
    Top = 0
    Width = 624
    Height = 290
    Align = alClient
    Constraints.MinHeight = 64
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #65325#65331' '#12468#12471#12483#12463
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssBoth
    TabOrder = 0
    WantTabs = True
    WordWrap = False
    OnKeyPress = mm_KeyPress
  end
  object mmError: TMemo
    Left = 0
    Top = 293
    Width = 624
    Height = 149
    Align = alBottom
    Constraints.MinHeight = 64
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #65325#65331' '#12468#12471#12483#12463
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 1
    WantTabs = True
    OnKeyPress = mm_KeyPress
  end
end
