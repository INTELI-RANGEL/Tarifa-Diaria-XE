object frmProcesaGenerador: TfrmProcesaGenerador
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'Proceso de generador'
  ClientHeight = 404
  ClientWidth = 389
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Position = poScreenCenter
  Visible = True
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object cxGroupBox1: TcxGroupBox
    Left = 8
    Top = 8
    Caption = 'Procesar Generadores'
    Style.LookAndFeel.NativeStyle = False
    Style.LookAndFeel.SkinName = 'Lilian'
    StyleDisabled.LookAndFeel.NativeStyle = False
    StyleDisabled.LookAndFeel.SkinName = 'Lilian'
    StyleFocused.LookAndFeel.NativeStyle = False
    StyleFocused.LookAndFeel.SkinName = 'Lilian'
    StyleHot.LookAndFeel.NativeStyle = False
    StyleHot.LookAndFeel.SkinName = 'Lilian'
    TabOrder = 0
    Height = 385
    Width = 369
    object Label1: TLabel
      Left = 15
      Top = 21
      Width = 22
      Height = 13
      Caption = 'Folio'
    end
    object CmbFolios: TCurvyCombo
      AlignWithMargins = True
      Left = 53
      Top = 18
      Width = 310
      Height = 24
      Margins.Left = 50
      Align = alTop
      TabOrder = 0
      Version = '1.1.0.0'
      Controls = <>
      Sorted = False
      Style = csDropDownList
      ExplicitLeft = 56
      ExplicitTop = 21
      ExplicitWidth = 304
    end
    object cbxEquipo: TcxCheckBox
      AlignWithMargins = True
      Left = 6
      Top = 73
      Align = alTop
      Caption = 'Equipo'
      State = cbsChecked
      Style.LookAndFeel.NativeStyle = False
      Style.LookAndFeel.SkinName = 'Lilian'
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.SkinName = 'Lilian'
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.SkinName = 'Lilian'
      StyleHot.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.SkinName = 'Lilian'
      TabOrder = 1
      Transparent = True
      ExplicitLeft = 64
      ExplicitTop = 80
      ExplicitWidth = 121
      Width = 357
    end
    object cbxPersonal: TcxCheckBox
      AlignWithMargins = True
      Left = 6
      Top = 48
      Align = alTop
      Caption = 'Personal'
      State = cbsChecked
      Style.LookAndFeel.NativeStyle = False
      Style.LookAndFeel.SkinName = 'Lilian'
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.SkinName = 'Lilian'
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.SkinName = 'Lilian'
      StyleHot.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.SkinName = 'Lilian'
      TabOrder = 2
      Transparent = True
      Width = 357
    end
    object mFolios: TcxMemo
      AlignWithMargins = True
      Left = 6
      Top = 159
      Align = alBottom
      TabOrder = 3
      ExplicitLeft = 15
      ExplicitTop = -7
      Height = 116
      Width = 357
    end
    object Panel2: TPanel
      AlignWithMargins = True
      Left = 6
      Top = 281
      Width = 357
      Height = 60
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 4
      ExplicitLeft = 9
      ExplicitTop = 226
      object cxButton1: TcxButton
        AlignWithMargins = True
        Left = 279
        Top = 15
        Width = 75
        Height = 30
        Margins.Top = 15
        Margins.Bottom = 15
        Align = alRight
        Caption = 'Folios en 0'
        TabOrder = 0
        OnClick = Button2Click
        LookAndFeel.NativeStyle = False
        LookAndFeel.SkinName = 'DevExpressStyle'
        ExplicitLeft = 144
        ExplicitTop = 8
        ExplicitHeight = 25
      end
      object cxButton2: TcxButton
        AlignWithMargins = True
        Left = 198
        Top = 15
        Width = 75
        Height = 30
        Margins.Top = 15
        Margins.Bottom = 15
        Align = alRight
        Caption = 'Procesar'
        TabOrder = 1
        OnClick = Button1Click
        LookAndFeel.NativeStyle = False
        LookAndFeel.SkinName = 'DevExpressStyle'
        ExplicitLeft = 144
        ExplicitTop = 8
        ExplicitHeight = 25
      end
    end
    object PanelProgress: TPanel
      AlignWithMargins = True
      Left = 6
      Top = 347
      Width = 357
      Height = 25
      Align = alBottom
      BevelOuter = bvNone
      ParentBackground = False
      TabOrder = 5
      ExplicitTop = 328
      object Label15: TLabel
        Left = 0
        Top = 0
        Width = 357
        Height = 16
        Align = alTop
        Caption = 'Procesando espere...'
        Color = 7237230
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        Visible = False
        ExplicitWidth = 134
      end
      object BarraEstado: TProgressBar
        AlignWithMargins = True
        Left = 3
        Top = 5
        Width = 351
        Height = 17
        Align = alBottom
        TabOrder = 0
        ExplicitLeft = 0
        ExplicitTop = -3
      end
    end
  end
  object ZFolios: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select distinct(snumeroorden) from actividadesxorden where sCont' +
        'rato = :Contrato order by snumeroorden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
    Left = 384
    Top = 8
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
  end
  object ZCeros: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select distinct(tb.folio) from'
      
        '(select distinct(snumeroorden) as folio from bitacoradepersonal ' +
        'where dCantHHGenerador = 0 and scontrato = :contrato'
      'union '
      
        'select distinct(snumeroorden) as folio from bitacoradeequipos wh' +
        'ere dCantHHGenerador = 0 and scontrato = :contrato) tb')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
    Left = 416
    Top = 8
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
  end
end
