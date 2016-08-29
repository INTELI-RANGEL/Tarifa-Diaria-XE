object frmgrupoperimetros: Tfrmgrupoperimetros
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  Caption = 'ALTA GRUPO PERIMETROS'
  ClientHeight = 278
  ClientWidth = 519
  Color = 14145495
  DefaultMonitor = dmPrimary
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Position = poDesktopCenter
  Visible = True
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 3
    Top = 214
    Width = 41
    Height = 14
    Caption = 'Id Grupo'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 3
    Top = 231
    Width = 57
    Height = 14
    Caption = 'Descripci'#243'n'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 3
    Top = 255
    Width = 37
    Height = 14
    Caption = 'Simbolo'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
  end
  object gridgrupoperimetros: TDBGrid
    Left = 73
    Top = -1
    Width = 438
    Height = 200
    Hint = 'Aqu'#237' se refleja el resultado de la consulta.'
    Color = 15138559
    Ctl3D = False
    DataSource = dsgrupoperimetros
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    Options = [dgTitles, dgIndicator, dgColLines, dgRowLines, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    ParentCtl3D = False
    ParentFont = False
    ReadOnly = True
    TabOrder = 3
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Arial'
    TitleFont.Style = [fsBold]
    OnCellClick = gridgrupoperimetrosCellClick
    OnTitleClick = gridgrupoperimetrosTitleClick
    Columns = <
      item
        Expanded = False
        FieldName = 'iIdGrupo'
        Title.Caption = 'ID'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'sDescripcion'
        Title.Caption = 'DESCRIPCION'
        Width = 263
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'sSimbolo'
        Title.Caption = 'SIMBOLO'
        Width = 88
        Visible = True
      end>
  end
  inline frmBarra1: TfrmBarra
    Left = 0
    Top = 0
    Width = 72
    Height = 207
    VertScrollBar.Style = ssHotTrack
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
    TabOrder = 4
    ExplicitWidth = 72
    ExplicitHeight = 207
    inherited AdvPanel1: TAdvPanel
      Width = 72
      Height = 207
      ParentShowHint = False
      ExplicitWidth = 72
      ExplicitHeight = 207
      FullHeight = 0
      inherited btnRefresh: TAdvGlowButton
        OnClick = frmBarra1btnRefreshClick
      end
      inherited btnEdit: TAdvGlowButton
        OnClick = frmBarra1btnEditClick
      end
      inherited btnPost: TAdvGlowButton
        OnClick = frmBarra1btnPostClick
      end
      inherited btnCancel: TAdvGlowButton
        Hint = 'Cancelar cambios (F11)'
        OnClick = frmBarra1btnCancelClick
      end
      inherited btnDelete: TAdvGlowButton
        OnClick = frmBarra1btnDeleteClick
      end
      inherited btnPrinter: TAdvGlowButton
        Enabled = False
      end
      inherited btnExit: TAdvGlowButton
        OnClick = frmBarra1btnExitClick
      end
      inherited btnAdd: TAdvGlowButton
        Left = 3
        OnClick = frmBarra1btnAddClick
        ExplicitLeft = 3
      end
    end
  end
  object tsIdGrupo: TDBEdit
    Left = 77
    Top = 205
    Width = 90
    Height = 22
    Hint = 'Id Grupo'
    CharCase = ecUpperCase
    Color = 15138559
    DataField = 'iIdGrupo'
    DataSource = dsgrupoperimetros
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    OnEnter = tsIdGrupoEnter
    OnExit = tsIdGrupoExit
    OnKeyPress = tsIdGrupoKeyPress
  end
  object tsDescripcion: TDBEdit
    Left = 77
    Top = 229
    Width = 432
    Height = 22
    Hint = 'Descripci'#243'n del Grupo'
    CharCase = ecUpperCase
    Color = 15138559
    DataField = 'sDescripcion'
    DataSource = dsgrupoperimetros
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnEnter = tsDescripcionEnter
    OnExit = tsDescripcionExit
    OnKeyPress = tsDescripcionKeyPress
  end
  object tsSimbolo: TDBEdit
    Left = 77
    Top = 253
    Width = 121
    Height = 21
    Hint = 'Simbolo'
    Color = 15138559
    DataField = 'sSimbolo'
    DataSource = dsgrupoperimetros
    TabOrder = 2
    OnEnter = tsSimboloEnter
    OnExit = tsSimboloExit
    OnKeyPress = tsSimboloKeyPress
  end
  object dsgrupoperimetros: TDataSource
    AutoEdit = False
    DataSet = zgrupoperimetros
    Left = 76
    Top = 41
  end
  object zgrupoperimetros: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select * from gruposperimetros Order By iIdGrupo')
    Params = <>
    UpdateMode = umUpdateAll
    WhereMode = wmWhereAll
    Left = 104
    Top = 41
    object zgrupoperimetrosiIdGrupo: TIntegerField
      FieldName = 'iIdGrupo'
    end
    object zgrupoperimetrossDescripcion: TStringField
      FieldName = 'sDescripcion'
      Required = True
      Size = 100
    end
    object zgrupoperimetrossSimbolo: TStringField
      FieldName = 'sSimbolo'
      Required = True
    end
  end
end
