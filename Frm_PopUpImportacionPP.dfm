object FrmPopUpImportacionPP: TFrmPopUpImportacionPP
  Left = 0
  Top = 0
  BorderStyle = bsToolWindow
  Caption = ' Importacion Puntos de Programa (Ms Project)'
  ClientHeight = 139
  ClientWidth = 337
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -10
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Scaled = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object NxPCImportacion: TNxPageControl
    Left = 0
    Top = 0
    Width = 337
    Height = 139
    ActivePage = NxTabSheet1
    ActivePageIndex = 0
    Align = alClient
    TabOrder = 0
    BackgroundColor = clSilver
    BackgroundKind = bkSolid
    Margin = 0
    Spacing = 0
    TabHeight = 17
    object NxTabSheet1: TNxTabSheet
      Caption = 'NxTabSheet1'
      PageIndex = 0
      ParentTabFont = False
      TabFont.Charset = DEFAULT_CHARSET
      TabFont.Color = clWindowText
      TabFont.Height = -13
      TabFont.Name = 'Tahoma'
      TabFont.Style = []
      object Label1: TLabel
        Left = 0
        Top = 0
        Width = 337
        Height = 16
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'Configuraci'#243'n de Orden de Trabajo y Folio'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Panel1: TPanel
        Left = 0
        Top = 16
        Width = 337
        Height = 102
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 0
        object Label2: TLabel
          Left = 55
          Top = 14
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
          Left = 55
          Top = 37
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
        object jDblCmbFolio: TJvDBLookupCombo
          Left = 112
          Top = 36
          Width = 187
          Height = 19
          LookupField = 'sNumeroOrden'
          LookupDisplay = 'sNumeroOrden'
          LookupSource = dsFolios
          TabOrder = 1
        end
        object btnCancelar: TNxButton
          Left = 242
          Top = 78
          Width = 57
          Height = 17
          Caption = '&Cancelar'
          ModalResult = 2
          TabOrder = 3
        end
        object jDblCmbContrato: TJvDBLookupCombo
          Left = 112
          Top = 14
          Width = 187
          Height = 19
          LookupField = 'sContrato'
          LookupDisplay = 'sContrato'
          LookupSource = dsContratos
          TabOrder = 0
          OnChange = jDblCmbContratoChange
        end
        object btnAceptar: TNxButton
          Left = 112
          Top = 78
          Width = 56
          Height = 17
          Caption = '&Aceptar'
          TabOrder = 2
          OnClick = btnAceptarClick
        end
      end
    end
    object NxTabSheet2: TNxTabSheet
      Caption = 'NxTabSheet2'
      PageIndex = 1
      ParentTabFont = False
      TabFont.Charset = DEFAULT_CHARSET
      TabFont.Color = clWindowText
      TabFont.Height = -13
      TabFont.Name = 'Tahoma'
      TabFont.Style = []
      object Label4: TLabel
        Left = 0
        Top = 0
        Width = 337
        Height = 16
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'Configurar Columnas Predefinidas'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object AdvGroupBox1: TAdvGroupBox
        Left = 0
        Top = 16
        Width = 337
        Height = 102
        BorderStyle = bsDouble
        RoundEdges = True
        Align = alClient
        TabOrder = 0
        object Label5: TLabel
          Left = 53
          Top = 14
          Width = 64
          Height = 13
          Caption = 'No. Partida  :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label6: TLabel
          Left = 53
          Top = 36
          Width = 65
          Height = 13
          Caption = 'Ponderado   :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object btnAceptar2: TNxButton
          Left = 112
          Top = 78
          Width = 56
          Height = 17
          Caption = '&Aceptar'
          ModalResult = 1
          TabOrder = 2
        end
        object btnCancelar2: TNxButton
          Left = 242
          Top = 78
          Width = 57
          Height = 17
          Caption = '&Cancelar'
          TabOrder = 3
          OnClick = btnCancelar2Click
        end
        object lCmbPartida: TLUCombo
          Left = 144
          Top = 14
          Width = 163
          Height = 20
          Color = clWindow
          Version = '2.3.1.0'
          Visible = True
          ButtonWidth = 23
          Style = csDropDownList
          DropWidth = 0
          Enabled = True
          ItemIndex = -1
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -13
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          TabOrder = 0
          AutoHistory = False
          AutoHistoryLimit = 0
          AutoHistoryDirection = ahdFirst
          AutoSynchronize = False
          ReturnIsTab = False
          FileLookup = False
          Persist.Enable = False
          Persist.Storage = stInifile
          Persist.Count = 0
          Persist.MaxCount = False
          ModifiedColor = clHighlight
          ShowModified = False
        end
        object lCmbPonderado: TLUCombo
          Left = 144
          Top = 36
          Width = 163
          Height = 20
          Color = clWindow
          Version = '2.3.1.0'
          Visible = True
          ButtonWidth = 23
          Style = csDropDownList
          DropWidth = 0
          Enabled = True
          ItemIndex = -1
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -13
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          TabOrder = 1
          AutoHistory = False
          AutoHistoryLimit = 0
          AutoHistoryDirection = ahdFirst
          AutoSynchronize = False
          ReturnIsTab = False
          FileLookup = False
          Persist.Enable = False
          Persist.Storage = stInifile
          Persist.Count = 0
          Persist.MaxCount = False
          ModifiedColor = clHighlight
          ShowModified = False
        end
      end
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
end
