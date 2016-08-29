object frmListado_Personal: TfrmListado_Personal
  Left = 0
  Top = 0
  Caption = 'Listado Personal'
  ClientHeight = 394
  ClientWidth = 636
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Scaled = False
  Visible = True
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  inline frmBarra1: TfrmBarra
    Left = 0
    Top = 33
    Width = 76
    Height = 222
    VertScrollBar.Style = ssHotTrack
    Align = alLeft
    TabOrder = 1
    ExplicitTop = 33
    ExplicitWidth = 76
    ExplicitHeight = 222
    inherited AdvPanel1: TAdvPanel
      Width = 76
      Height = 222
      Align = alRight
      ParentShowHint = False
      ShowHint = True
      ExplicitWidth = 76
      ExplicitHeight = 222
      FullHeight = 0
      inherited btnRefresh: TAdvGlowButton
        Left = 5
        Top = 161
        Width = 71
        Height = 26
        OnClick = frmBarra1btnRefreshClick
        ExplicitLeft = 5
        ExplicitTop = 161
        ExplicitWidth = 71
        ExplicitHeight = 26
      end
      inherited btnEdit: TAdvGlowButton
        Left = 5
        Width = 71
        Height = 26
        OnClick = frmBarra1btnEditClick
        ExplicitLeft = 5
        ExplicitWidth = 71
        ExplicitHeight = 26
      end
      inherited btnPost: TAdvGlowButton
        Left = 5
        Top = 54
        Width = 71
        Height = 28
        OnClick = frmBarra1btnPostClick
        ExplicitLeft = 5
        ExplicitTop = 54
        ExplicitWidth = 71
        ExplicitHeight = 28
      end
      inherited btnCancel: TAdvGlowButton
        Left = 5
        Top = 82
        Width = 71
        Height = 26
        OnClick = frmBarra1btnCancelClick
        ExplicitLeft = 5
        ExplicitTop = 82
        ExplicitWidth = 71
        ExplicitHeight = 26
      end
      inherited btnDelete: TAdvGlowButton
        Left = 5
        Top = 108
        Width = 71
        Height = 27
        OnClick = frmBarra1btnDeleteClick
        ExplicitLeft = 5
        ExplicitTop = 108
        ExplicitWidth = 71
        ExplicitHeight = 27
      end
      inherited btnPrinter: TAdvGlowButton
        Left = 5
        Top = 135
        Width = 71
        Height = 27
        ExplicitLeft = 5
        ExplicitTop = 135
        ExplicitWidth = 71
        ExplicitHeight = 27
      end
      inherited btnExit: TAdvGlowButton
        Left = 5
        Top = 188
        Width = 71
        Height = 28
        OnClick = frmBarra1btnExitClick
        ExplicitLeft = 5
        ExplicitTop = 188
        ExplicitWidth = 71
        ExplicitHeight = 28
      end
      inherited btnAdd: TAdvGlowButton
        Left = 5
        Top = 1
        Width = 71
        Height = 27
        OnClick = frmBarra1btnAddClick
        ExplicitLeft = 5
        ExplicitTop = 1
        ExplicitWidth = 71
        ExplicitHeight = 27
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 255
    Width = 636
    Height = 139
    Align = alBottom
    TabOrder = 3
    object Label1: TLabel
      Left = 5
      Top = 9
      Width = 25
      Height = 13
      Caption = 'Ficha'
    end
    object Label2: TLabel
      Left = 5
      Top = 36
      Width = 37
      Height = 13
      Caption = 'Nombre'
    end
    object Label3: TLabel
      Left = 5
      Top = 63
      Width = 78
      Height = 13
      Caption = 'Apellido Paterno'
    end
    object Label4: TLabel
      Left = 5
      Top = 90
      Width = 80
      Height = 13
      Caption = 'Apellido Materno'
    end
    object Label5: TLabel
      Left = 253
      Top = 9
      Width = 47
      Height = 13
      Caption = 'Categoria'
    end
    object Label6: TLabel
      Left = 5
      Top = 117
      Width = 16
      Height = 13
      Caption = 'Rfc'
    end
    object Label7: TLabel
      Left = 253
      Top = 63
      Width = 47
      Height = 13
      Caption = 'Compa'#241'ia'
    end
    object Label8: TLabel
      Left = 253
      Top = 36
      Width = 58
      Height = 13
      Caption = 'Especialidad'
    end
    object Label9: TLabel
      Left = 253
      Top = 90
      Width = 69
      Height = 13
      Caption = 'Libreta de Mar'
    end
    object Label10: TLabel
      Left = 253
      Top = 114
      Width = 101
      Height = 13
      Caption = 'Vigencia de la Libreta'
    end
    object edtFicha: TDBEdit
      Left = 88
      Top = 6
      Width = 140
      Height = 21
      Color = 15138559
      DataField = 'sIdTripulacion'
      DataSource = ds_listadoper
      TabOrder = 0
      OnEnter = edtFichaEnter
      OnExit = edtFichaExit
      OnKeyPress = edtFichaKeyPress
    end
    object edtNombre: TDBEdit
      Left = 88
      Top = 33
      Width = 140
      Height = 21
      Color = 15138559
      DataField = 'sNombre'
      DataSource = ds_listadoper
      TabOrder = 2
      OnEnter = edtNombreEnter
      OnExit = edtNombreExit
      OnKeyPress = edtNombreKeyPress
    end
    object edtApellidoP: TDBEdit
      Left = 89
      Top = 60
      Width = 140
      Height = 21
      Color = 15138559
      DataField = 'sApellidoP'
      DataSource = ds_listadoper
      TabOrder = 4
      OnEnter = edtApellidoPEnter
      OnExit = edtApellidoPExit
      OnKeyPress = edtApellidoPKeyPress
    end
    object edtApellidoM: TDBEdit
      Left = 88
      Top = 87
      Width = 140
      Height = 21
      Color = 15138559
      DataField = 'sApellidoM'
      DataSource = ds_listadoper
      TabOrder = 6
      OnEnter = edtApellidoMEnter
      OnExit = edtApellidoMExit
      OnKeyPress = edtApellidoMKeyPress
    end
    object edtRfc: TDBEdit
      Left = 88
      Top = 113
      Width = 140
      Height = 21
      Color = 15138559
      DataField = 'sRfc'
      DataSource = ds_listadoper
      TabOrder = 8
      OnEnter = edtRfcEnter
      OnExit = edtRfcExit
      OnKeyPress = edtRfcKeyPress
    end
    object lkcbCompania: TDBLookupComboBox
      Left = 360
      Top = 60
      Width = 268
      Height = 21
      Color = 15138559
      DataField = 'sIdCompania'
      DataSource = ds_listadoper
      KeyField = 'sIdCompania'
      ListField = 'sDescripcion'
      ListSource = ds_compania
      TabOrder = 5
      OnEnter = lkcbCompaniaEnter
      OnExit = lkcbCompaniaExit
      OnKeyPress = lkcbCompaniaKeyPress
    end
    object lkcbEsp: TDBLookupComboBox
      Left = 360
      Top = 33
      Width = 268
      Height = 21
      Color = 15138559
      DataField = 'sIdPersonal'
      DataSource = ds_listadoper
      KeyField = 'sIdPersonal'
      ListField = 'esp '
      ListSource = ds_Esp
      TabOrder = 3
      OnEnter = lkcbEspEnter
      OnExit = lkcbEspExit
      OnKeyPress = lkcbEspKeyPress
    end
    object lkcbCategoria: TDBLookupComboBox
      Left = 360
      Top = 6
      Width = 268
      Height = 21
      Color = 15138559
      KeyField = 'sIdTipoPersonal'
      ListField = 'sDescripcion'
      ListSource = ds_categoria
      TabOrder = 1
      OnEnter = lkcbCategoriaEnter
      OnExit = lkcbCategoriaExit
      OnKeyPress = lkcbCategoriaKeyPress
    end
    object edtLibreta: TDBEdit
      Left = 360
      Top = 87
      Width = 268
      Height = 21
      Color = 15138559
      DataField = 'sLibretadeMar'
      DataSource = ds_listadoper
      TabOrder = 7
      OnEnter = edtLibretaEnter
      OnExit = edtLibretaExit
      OnKeyPress = edtLibretaKeyPress
    end
    object dedtVigencia: TDBDateEdit
      Left = 360
      Top = 113
      Width = 268
      Height = 21
      Margins.Left = 4
      Margins.Top = 1
      DataField = 'dVigencia'
      DataSource = ds_listadoper
      Color = 15138559
      NumGlyphs = 2
      TabOrder = 9
      OnEnter = dedtVigenciaEnter
      OnExit = dedtVigenciaExit
    end
  end
  object GridListPer: TDBGrid
    Left = 76
    Top = 33
    Width = 560
    Height = 222
    Align = alClient
    Color = 15138559
    DataSource = ds_listadoper
    ReadOnly = True
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'sIdTripulacion'
        Title.Caption = 'Ficha'
        Width = 55
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'sNombre'
        Title.Caption = 'Nombre'
        Width = 109
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'sApellidoP'
        Title.Caption = 'Apellido Paterno'
        Width = 126
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'sApellidoM'
        Title.Caption = 'Apellido Materno'
        Width = 105
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'sRfc'
        Title.Caption = 'Rfc'
        Width = 134
        Visible = True
      end>
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 636
    Height = 33
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object edtBuscar: TEdit
      Left = 183
      Top = 7
      Width = 279
      Height = 21
      Color = 15138559
      TabOrder = 1
      OnKeyPress = edtBuscarKeyPress
    end
    object cbbBuscar: TAdvComboBox
      Left = 89
      Top = 7
      Width = 88
      Height = 21
      Color = 15138559
      Version = '1.3.2.2'
      Visible = True
      Style = csDropDownList
      DropWidth = 0
      Enabled = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        'Ficha'
        'Nombre')
      LabelCaption = 'Buscar Por: '
      LabelPosition = lpLeftCenter
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clWindowText
      LabelFont.Height = -9
      LabelFont.Name = 'Tahoma'
      LabelFont.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      Text = 'Ficha'
      OnKeyPress = cbbBuscarKeyPress
    end
    object btnBuscar: TButton
      Left = 468
      Top = 8
      Width = 49
      Height = 20
      Caption = '&Buscar'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = btnBuscarClick
    end
  end
  object zq_listadoper: TZQuery
    Connection = connection.zConnection
    AfterScroll = zq_listadoperAfterScroll
    SQL.Strings = (
      'select * from tripulacion_listado')
    Params = <>
    Left = 232
    Top = 80
  end
  object zq_compania: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'Select * from compersonal')
    Params = <>
    Left = 232
    Top = 112
  end
  object zq_categoria: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select sIdTipoPersonal, sDescripcion from tiposdepersonal'
      'order by sDescripcion')
    Params = <>
    Left = 232
    Top = 144
  end
  object ds_listadoper: TDataSource
    DataSet = zq_listadoper
    Left = 272
    Top = 80
  end
  object ds_compania: TDataSource
    DataSet = zq_compania
    Left = 272
    Top = 112
  end
  object ds_categoria: TDataSource
    DataSet = zq_categoria
    OnDataChange = ds_categoriaDataChange
    Left = 272
    Top = 144
  end
  object zq_Esp: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select sIdPersonal, sDescripcion, sIdTipoPersonal, CONCAT(sIdPer' +
        'sonal, '#39' '#39', sDescripcion) as esp '
      'from personal'
      'where sContrato=:Contrato and '
      
        '(:TipoPer = -1 or (:TipoPer <> -1 and sIdTipoPersonal = :TipoPer' +
        ')) and '
      '(:Per = -1 or (:Per <> -1 and sIdPersonal = :Per))')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'TipoPer'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'per'
        ParamType = ptUnknown
      end>
    Left = 232
    Top = 176
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'TipoPer'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'per'
        ParamType = ptUnknown
      end>
  end
  object ds_Esp: TDataSource
    DataSet = zq_Esp
    Left = 272
    Top = 176
  end
end
