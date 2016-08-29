object FrmImportaProject: TFrmImportaProject
  Left = 0
  Top = 0
  Caption = 'I M P O R T A C I O N  D E  F O L I O S  E N  M S P R O J E C T'
  ClientHeight = 562
  ClientWidth = 937
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -10
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Scaled = False
  Visible = True
  WindowState = wsMaximized
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object Panel2: TPanel
    Left = 0
    Top = 63
    Width = 937
    Height = 499
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    object Panel3: TPanel
      Left = 0
      Top = 20
      Width = 216
      Height = 479
      Align = alLeft
      BevelOuter = bvNone
      TabOrder = 1
      ExplicitTop = 21
      ExplicitHeight = 478
      object NxInsColumnas: TNextInspector
        Left = 0
        Top = 17
        Width = 216
        Height = 462
        Align = alClient
        TabOrder = 1
        ExplicitHeight = 461
      end
      object Panel5: TPanel
        Left = 0
        Top = 0
        Width = 216
        Height = 17
        Align = alTop
        BevelOuter = bvLowered
        Caption = 'Definir Parametros de Importacion'
        Color = 11053056
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -10
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 0
      end
    end
    object cxSplitter1: TcxSplitter
      Left = 216
      Top = 20
      Width = 6
      Height = 479
      Control = Panel3
      ExplicitTop = 21
      ExplicitHeight = 478
    end
    object CxTlsPrograma: TcxTreeList
      Left = 222
      Top = 20
      Width = 715
      Height = 479
      Align = alClient
      Bands = <>
      Navigator.Buttons.CustomButtons = <>
      Styles.OnGetContentStyle = CxTlsProgramaStylesGetContentStyle
      TabOrder = 3
      ExplicitTop = 21
      ExplicitHeight = 478
    end
    object CxPbAvance: TcxProgressBar
      Left = 0
      Top = 0
      Align = alTop
      Properties.ShowOverload = True
      Properties.ShowPeak = True
      Style.Shadow = True
      TabOrder = 0
      Transparent = True
      Visible = False
      Width = 937
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 937
    Height = 63
    Align = alTop
    BevelInner = bvLowered
    BevelKind = bkSoft
    BevelOuter = bvSpace
    TabOrder = 0
    object Panel4: TPanel
      Left = 2
      Top = 2
      Width = 929
      Height = 55
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 0
      object Label2: TLabel
        Left = 20
        Top = 9
        Width = 53
        Height = 13
        Caption = 'Contrato:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label3: TLabel
        Left = 20
        Top = 32
        Width = 29
        Height = 13
        Caption = 'Folio:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object JFedtArchivo: TJvFilenameEdit
        Left = 312
        Top = 21
        Width = 360
        Height = 24
        OnAfterDialog = JFedtArchivoAfterDialog
        AddQuotes = False
        Filter = 'Microsoft Project |*.mpp'
        DirectInput = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = []
        ButtonWidth = 16
        ParentFont = False
        TabOrder = 2
      end
      object cxLabel1: TcxLabel
        Left = 315
        Top = 3
        Caption = 'Archivo:'
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Arial'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object btnImportar: TButton
        Left = 718
        Top = 22
        Width = 75
        Height = 25
        Caption = '&Importar'
        Enabled = False
        TabOrder = 3
        OnClick = btnImportarClick
      end
      object jDblCmbFolio: TJvDBLookupCombo
        Left = 77
        Top = 30
        Width = 187
        Height = 23
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = []
        LookupField = 'sNumeroOrden'
        LookupDisplay = 'sNumeroOrden'
        LookupSource = dsFolios
        ParentFont = False
        TabOrder = 4
      end
      object jDblCmbContrato: TJvDBLookupCombo
        Left = 77
        Top = 3
        Width = 187
        Height = 23
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = []
        LookupField = 'sContrato'
        LookupDisplay = 'sContrato'
        LookupSource = dsContratos
        ParentFont = False
        TabOrder = 0
        OnChange = jDblCmbContratoChange
      end
    end
  end
  object ArchivoMsP: TFileOpenDialog
    FavoriteLinks = <>
    FileTypes = <
      item
        DisplayName = 'Microsoft Project '
        FileMask = '*.mpp'
      end>
    Options = []
    Left = 280
    Top = 152
  end
  object cxStyleRepository1: TcxStyleRepository
    PixelsPerInch = 96
    object cxStyle1: TcxStyle
      AssignedValues = [svFont]
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
    end
  end
  object QrContratos: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select sContrato from contratos')
    Params = <>
    Left = 16
    Top = 8
  end
  object dsContratos: TDataSource
    AutoEdit = False
    DataSet = QrContratos
    Left = 16
    Top = 40
  end
  object dsFolios: TDataSource
    AutoEdit = False
    DataSet = QrFolios
    Left = 16
    Top = 80
  end
  object QrFolios: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select sContrato, sNumeroOrden from ordenesdetrabajo')
    Params = <>
    MasterFields = 'sContrato'
    MasterSource = dsContratos
    LinkedFields = 'sContrato'
    Left = 16
    Top = 112
  end
  object StyleRepository: TcxStyleRepository
    Left = 160
    Top = 8
    PixelsPerInch = 96
    object cxStyle2: TcxStyle
      AssignedValues = [svColor]
      Color = 15451300
    end
    object cxStyle3: TcxStyle
      AssignedValues = [svColor]
      Color = 15451300
    end
    object cxStyle4: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 12937777
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      TextColor = clWhite
    end
    object cxStyle5: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 15252642
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      TextColor = 11032875
    end
    object cxStyle6: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 16247513
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      TextColor = clBlack
    end
    object cxStyle7: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 15784893
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      TextColor = clBlack
    end
    object cxStyle8: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 16247513
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      TextColor = clBlack
    end
    object cxStyle9: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 14811135
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      TextColor = clNavy
    end
    object cxStyle10: TcxStyle
      AssignedValues = [svColor]
      Color = 15451300
    end
    object cxStyle11: TcxStyle
      AssignedValues = [svColor, svTextColor]
      Color = 4707838
      TextColor = clBlack
    end
    object cxStyle12: TcxStyle
      AssignedValues = [svColor, svTextColor]
      Color = 15451300
      TextColor = clBlack
    end
    object cxStyle13: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 14811135
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      TextColor = clNavy
    end
    object cxStyle14: TcxStyle
      AssignedValues = [svColor, svTextColor]
      Color = 16048336
      TextColor = clBlack
    end
    object stlGroupNode: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 15253902
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      TextColor = clWhite
    end
    object stlFixedBand: TcxStyle
      AssignedValues = [svColor]
      Color = 15322014
    end
    object TreeListStyleSheetDevExpress: TcxTreeListStyleSheet
      Caption = 'DevExpress'
      Styles.Background = cxStyle2
      Styles.Content = cxStyle6
      Styles.Inactive = cxStyle10
      Styles.Selection = cxStyle14
      Styles.BandBackground = cxStyle3
      Styles.BandHeader = cxStyle4
      Styles.ColumnHeader = cxStyle5
      Styles.ContentEven = cxStyle8
      Styles.ContentOdd = cxStyle7
      Styles.Footer = cxStyle9
      Styles.IncSearch = cxStyle11
      Styles.Indicator = cxStyle12
      Styles.Preview = cxStyle13
      BuiltIn = True
    end
  end
end
