object frmFormatos: TfrmFormatos
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  Caption = 'Formatos de reportes'
  ClientHeight = 463
  ClientWidth = 815
  Color = 16316664
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 81
    Top = 192
    Width = 39
    Height = 13
    Caption = 'Reporte'
  end
  object Label2: TLabel
    Left = 81
    Top = 219
    Width = 34
    Height = 13
    Caption = 'Modulo'
  end
  object Label3: TLabel
    Left = 81
    Top = 246
    Width = 54
    Height = 13
    Caption = 'Descripci'#243'n'
  end
  object pnlCarga: TAdvSmoothPanel
    AlignWithMargins = True
    Left = 10
    Top = 344
    Width = 405
    Height = 109
    Cursor = crDefault
    Margins.Left = 10
    Margins.Top = 10
    Margins.Right = 400
    Margins.Bottom = 10
    Caption.Text = 'Carga'
    Caption.HTMLFont.Charset = DEFAULT_CHARSET
    Caption.HTMLFont.Color = clWindowText
    Caption.HTMLFont.Height = -11
    Caption.HTMLFont.Name = 'Tahoma'
    Caption.HTMLFont.Style = []
    Caption.Font.Charset = DEFAULT_CHARSET
    Caption.Font.Color = clWindowText
    Caption.Font.Height = -16
    Caption.Font.Name = 'Calibri'
    Caption.Font.Style = []
    Caption.ColorStart = 5978398
    Caption.ColorEnd = 5978398
    Caption.LineColor = 5978398
    Caption.TextRendering = tAntiAlias
    Fill.Color = clWhite
    Fill.ColorTo = 15590880
    Fill.ColorMirror = clNone
    Fill.ColorMirrorTo = clNone
    Fill.GradientMirrorType = gtVertical
    Fill.BorderColor = 13815240
    Fill.Rounding = 10
    Fill.ShadowOffset = 10
    Fill.Glow = gmRadial
    Version = '1.0.9.6'
    Align = alBottom
    TabOrder = 0
    object txtArchivo: TFilenameEdit
      Left = 14
      Top = 39
      Width = 296
      Height = 19
      OnAfterDialog = txtArchivoAfterDialog
      Filter = 'Excel (*.xlsx)|*.xlsx|Excel (*.xlsm)|*.xlsm'
      Ctl3D = False
      NumGlyphs = 1
      ParentCtl3D = False
      TabOrder = 0
    end
    object btnSubir: TcxButton
      Left = 14
      Top = 65
      Width = 75
      Height = 25
      Caption = 'Subir a la BD'
      Enabled = False
      TabOrder = 1
      OnClick = btnSubirClick
    end
    object cxButton5: TcxButton
      Left = 95
      Top = 65
      Width = 75
      Height = 25
      Caption = 'Exportar'
      TabOrder = 2
      OnClick = cxButton5Click
    end
    object chkAbrirFinalizar: TAdvOfficeCheckBox
      Left = 180
      Top = 67
      Width = 108
      Height = 20
      Checked = True
      TabOrder = 3
      Alignment = taLeftJustify
      Caption = 'Abrir al finalizar'
      ReturnIsTab = False
      State = cbChecked
      Version = '1.3.2.0'
    end
  end
  inline frmBarra1: TfrmBarra
    Left = 3
    Top = 2
    Width = 70
    Height = 175
    VertScrollBar.Style = ssHotTrack
    Align = alCustom
    Color = 7847370
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    ExplicitLeft = 3
    ExplicitTop = 2
    ExplicitWidth = 70
    ExplicitHeight = 175
    inherited AdvPanel1: TAdvPanel
      Width = 81
      Height = 224
      Align = alCustom
      Color = 16645114
      Font.Color = 7485192
      ParentShowHint = False
      ShowHint = True
      BorderColor = 16765615
      BorderShadow = False
      Caption.Color = 16575452
      Caption.ColorTo = 16571329
      Caption.GradientDirection = gdVertical
      Caption.Indent = 2
      Caption.ShadeType = stNormal
      CollapsColor = clNone
      ColorTo = 16643051
      StatusBar.BorderColor = 13542013
      StatusBar.Font.Color = 7485192
      StatusBar.Color = 16575452
      StatusBar.ColorTo = 16571329
      ExplicitWidth = 81
      ExplicitHeight = 224
      FullHeight = 0
      inherited btnRefresh: TAdvGlowButton
        Left = 0
        Top = 122
        OnClick = frmBarra1btnRefreshClick
        ExplicitLeft = 0
        ExplicitTop = 122
      end
      inherited btnEdit: TAdvGlowButton
        Left = 0
        Top = 25
        OnClick = frmBarra1btnEditClick
        ExplicitLeft = 0
        ExplicitTop = 25
      end
      inherited btnPost: TAdvGlowButton
        Left = 0
        Top = 50
        Caption = #39
        OnClick = frmBarra1btnPostClick
        ExplicitLeft = 0
        ExplicitTop = 50
      end
      inherited btnCancel: TAdvGlowButton
        Left = 0
        Top = 74
        OnClick = frmBarra1btnCancelClick
        ExplicitLeft = 0
        ExplicitTop = 74
      end
      inherited btnDelete: TAdvGlowButton
        Left = 0
        Top = 98
        OnClick = frmBarra1btnDeleteClick
        ExplicitLeft = 0
        ExplicitTop = 98
      end
      inherited btnPrinter: TAdvGlowButton
        Left = 0
        Top = 204
        Visible = False
        Appearance.BorderColor = 13542013
        Appearance.BorderColorHot = 16504504
        Appearance.BorderColorDown = 13542013
        Appearance.BorderColorChecked = clHighlight
        Appearance.Color = 16575452
        Appearance.ColorTo = 16571329
        Appearance.ColorChecked = 16645114
        Appearance.ColorCheckedTo = 16643051
        Appearance.ColorDown = 16575452
        Appearance.ColorDownTo = 16571329
        Appearance.ColorHot = 16645114
        Appearance.ColorHotTo = 16643051
        Appearance.ColorMirror = 16571329
        Appearance.ColorMirrorTo = 16571329
        Appearance.ColorMirrorHot = 16643051
        Appearance.ColorMirrorHotTo = 16645114
        Appearance.ColorMirrorDown = 16571329
        Appearance.ColorMirrorDownTo = 16575452
        Appearance.ColorMirrorChecked = 16575452
        Appearance.ColorMirrorCheckedTo = 16575452
        Appearance.GradientHot = ggVertical
        Appearance.GradientMirrorHot = ggVertical
        Appearance.GradientDown = ggVertical
        Appearance.GradientMirrorDown = ggVertical
        Appearance.GradientChecked = ggVertical
        ExplicitLeft = 0
        ExplicitTop = 204
      end
      inherited btnExit: TAdvGlowButton
        Left = 0
        Top = 147
        OnClick = frmBarra1btnExitClick
        ExplicitLeft = 0
        ExplicitTop = 147
      end
      inherited btnAdd: TAdvGlowButton
        Left = 0
        Top = 2
        OnClick = frmBarra1btnAddClick
        ExplicitLeft = 0
        ExplicitTop = 2
      end
    end
    inherited StylePanel: TAdvPanelStyler
      Settings.BorderColor = 16765615
      Settings.BorderShadow = False
      Settings.Caption.Color = 16575452
      Settings.Caption.ColorTo = 16571329
      Settings.Caption.GradientDirection = gdVertical
      Settings.Caption.Indent = 2
      Settings.Caption.ShadeType = stNormal
      Settings.Caption.Visible = False
      Settings.CollapsColor = clNone
      Settings.Color = 16645114
      Settings.ColorTo = 16643051
      Settings.Font.Color = 7485192
      Settings.StatusBar.BorderColor = 13542013
      Settings.StatusBar.Font.Color = 7485192
      Settings.StatusBar.Color = 16575452
      Settings.StatusBar.ColorTo = 16571329
      Style = psWindows7
    end
  end
  object DBGrid1: TDBGrid
    Left = 76
    Top = 5
    Width = 733
    Height = 169
    Ctl3D = False
    DataSource = dsFormatos
    ParentCtl3D = False
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
        FieldName = 'sTitulo'
        Title.Alignment = taCenter
        Title.Caption = 'Reporte'
        Width = 113
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'sModulo'
        Title.Alignment = taCenter
        Title.Caption = 'Modulo'
        Width = 122
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'mDescripcion'
        Title.Alignment = taCenter
        Title.Caption = 'Descripcion'
        Width = 328
        Visible = True
      end
      item
        Color = 13035452
        Expanded = False
        FieldName = 'sTipo'
        Title.Alignment = taCenter
        Title.Caption = 'Tipo archivo'
        Visible = True
      end>
  end
  object dbModulo: TDBEdit
    Left = 141
    Top = 218
    Width = 200
    Height = 19
    Ctl3D = False
    DataField = 'sModulo'
    DataSource = dsFormatos
    ParentCtl3D = False
    TabOrder = 3
    OnEnter = dbTituloEnter
    OnExit = dbTituloExit
    OnKeyPress = dbModuloKeyPress
  end
  object dbDescripcion: TDBEdit
    Left = 141
    Top = 246
    Width = 436
    Height = 19
    Ctl3D = False
    DataField = 'mDescripcion'
    DataSource = dsFormatos
    ParentCtl3D = False
    TabOrder = 4
    OnEnter = dbTituloEnter
    OnExit = dbTituloExit
    OnKeyPress = dbDescripcionKeyPress
  end
  object dbTitulo: TDBEdit
    Left = 141
    Top = 190
    Width = 200
    Height = 19
    Ctl3D = False
    DataField = 'sTitulo'
    DataSource = dsFormatos
    ParentCtl3D = False
    TabOrder = 5
    OnEnter = dbTituloEnter
    OnExit = dbTituloExit
    OnKeyPress = dbTituloKeyPress
  end
  object qrFormatos: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select * from tarifa_formatos where sContrato = :contrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end>
    Left = 8
    Top = 208
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end>
  end
  object dsFormatos: TDataSource
    DataSet = qrFormatos
    Left = 8
    Top = 240
  end
end
