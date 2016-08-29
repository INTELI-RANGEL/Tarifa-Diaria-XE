object FrmBuscaPersonal: TFrmBuscaPersonal
  Left = 0
  Top = 0
  BorderIcons = []
  BorderStyle = bsToolWindow
  Caption = 'Personal'
  ClientHeight = 427
  ClientWidth = 1135
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  Scaled = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object NxPCPersonal: TNxPageControl
    Left = 0
    Top = 0
    Width = 1135
    Height = 427
    ActivePage = NxTabSheet1
    ActivePageIndex = 0
    Align = alClient
    TabOrder = 0
    OnChange = NxPCPersonalChange
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
      TabFont.Height = -11
      TabFont.Name = 'Tahoma'
      TabFont.Style = []
      ExplicitTop = 0
      object Panel1: TPanel
        Left = 0
        Top = 365
        Width = 1135
        Height = 41
        Align = alBottom
        BevelKind = bkSoft
        TabOrder = 1
        object Panel2: TPanel
          Left = 888
          Top = 1
          Width = 242
          Height = 35
          Align = alRight
          BevelOuter = bvNone
          TabOrder = 1
          object btnCerrar: TButton
            Left = 143
            Top = 5
            Width = 75
            Height = 25
            Caption = '&Cerrar'
            TabOrder = 1
            OnClick = btnCerrarClick
          end
          object btnAgregar: TButton
            Left = 39
            Top = 5
            Width = 75
            Height = 25
            Caption = '&Agregar'
            TabOrder = 0
            OnClick = btnAgregarClick
          end
        end
        object Panel3: TPanel
          Left = 1
          Top = 1
          Width = 88
          Height = 35
          Align = alLeft
          BevelOuter = bvNone
          TabOrder = 0
          object btnNuevo: TButton
            Left = 16
            Top = 5
            Width = 57
            Height = 25
            Caption = '&Nuevo'
            TabOrder = 0
            OnClick = btnNuevoClick
          end
        end
      end
      object cxgrdListado: TcxGrid
        Left = 0
        Top = 0
        Width = 1135
        Height = 365
        Align = alClient
        TabOrder = 0
        LookAndFeel.Kind = lfStandard
        LookAndFeel.NativeStyle = False
        object cxDbGridListadoDBTable: TcxGridDBTableView
          OnDblClick = cxDbGridListadoDBTableDblClick
          Navigator.Buttons.CustomButtons = <>
          OnCellDblClick = cxDbGridListadoDBTableCellDblClick
          OnEditDblClick = cxDbGridListadoDBTableEditDblClick
          DataController.DataSource = dsListado
          DataController.Summary.DefaultGroupSummaryItems = <>
          DataController.Summary.FooterSummaryItems = <>
          DataController.Summary.SummaryGroups = <>
          FilterRow.Visible = True
          OptionsView.GroupByBox = False
          object cxDbGridListadoDBTableColumn1: TcxGridDBColumn
            Caption = 'Ficha'
            DataBinding.FieldName = 'sIdTripulacion'
            PropertiesClassName = 'TcxTextEditProperties'
            Properties.Alignment.Horz = taCenter
            HeaderAlignmentHorz = taCenter
            Width = 70
          end
          object cxDbGridListadoDBTableColumn2: TcxGridDBColumn
            Caption = 'Nombre'
            DataBinding.FieldName = 'NombreCompleto'
            HeaderAlignmentHorz = taCenter
            Width = 300
          end
          object cxDbGridListadoDBTableColumn3: TcxGridDBColumn
            Caption = 'Id Categoria'
            DataBinding.FieldName = 'sIdPersonal'
            HeaderAlignmentHorz = taCenter
            Width = 80
          end
          object cxDbGridListadoDBTableColumn4: TcxGridDBColumn
            Caption = 'Categoria'
            DataBinding.FieldName = 'categoria'
            HeaderAlignmentHorz = taCenter
            Options.Editing = False
            Width = 150
          end
          object cxDbGridListadoDBTableColumn5: TcxGridDBColumn
            Caption = 'RFC'
            DataBinding.FieldName = 'sRfc'
            HeaderAlignmentHorz = taCenter
            Width = 90
          end
          object cxDbGridListadoDBTableColumn6: TcxGridDBColumn
            Caption = 'Compa'#241'ia'
            DataBinding.FieldName = 'compania'
            HeaderAlignmentHorz = taCenter
            Width = 120
          end
          object cxDbGridListadoDBTableColumn7: TcxGridDBColumn
            Caption = 'Tipo de Personal'
            DataBinding.FieldName = 'TipoPersonal'
            FooterAlignmentHorz = taCenter
            HeaderAlignmentHorz = taCenter
            Width = 300
          end
        end
        object cxgrdListadoLevel1: TcxGridLevel
          GridView = cxDbGridListadoDBTable
        end
      end
    end
    object NxTabSheet2: TNxTabSheet
      Caption = 'NxTabSheet2'
      PageIndex = 1
      ParentTabFont = False
      TabFont.Charset = DEFAULT_CHARSET
      TabFont.Color = clWindowText
      TabFont.Height = -11
      TabFont.Name = 'Tahoma'
      TabFont.Style = []
      object pnlContenedor: TPanel
        Left = 0
        Top = 0
        Width = 1135
        Height = 406
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 0
        object Panel4: TPanel
          Left = 0
          Top = 261
          Width = 1135
          Height = 145
          Align = alBottom
          TabOrder = 1
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
            Top = 117
            Width = 101
            Height = 13
            Caption = 'Vigencia de la Libreta'
          end
          object edtFicha: TDBEdit
            Left = 89
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
            Left = 89
            Top = 114
            Width = 139
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
            Left = 359
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
            Left = 359
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
            Left = 359
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
            Left = 359
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
            Left = 359
            Top = 114
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
        object Panel5: TPanel
          Left = 0
          Top = 0
          Width = 1135
          Height = 261
          Align = alClient
          Caption = 'Panel5'
          TabOrder = 0
          inline frmBarra1: TfrmBarra
            Left = 1
            Top = 34
            Width = 73
            Height = 226
            VertScrollBar.Style = ssHotTrack
            Align = alLeft
            TabOrder = 1
            ExplicitLeft = 1
            ExplicitTop = 34
            ExplicitWidth = 73
            ExplicitHeight = 226
            inherited AdvPanel1: TAdvPanel
              Width = 73
              Height = 226
              ExplicitWidth = 73
              ExplicitHeight = 226
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
                OnClick = frmBarra1btnCancelClick
              end
              inherited btnDelete: TAdvGlowButton
                OnClick = frmBarra1btnDeleteClick
              end
              inherited btnExit: TAdvGlowButton
                OnClick = frmBarra1btnExitClick
              end
              inherited btnAdd: TAdvGlowButton
                OnClick = frmBarra1btnAddClick
              end
            end
          end
          object GridListPer: TDBGrid
            Left = 74
            Top = 34
            Width = 1060
            Height = 226
            TabStop = False
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
          object Panel6: TPanel
            Left = 1
            Top = 1
            Width = 1133
            Height = 33
            Align = alTop
            BevelOuter = bvNone
            TabOrder = 0
            object edtBuscar: TEdit
              Left = 183
              Top = 6
              Width = 279
              Height = 21
              Color = 15138559
              TabOrder = 0
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
              TabOrder = 1
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
        end
      end
    end
  end
  object QListado: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select tl.*,concat(tl.sNombre,'#39' '#39',tl.sApellidoP,'#39' '#39',tl.sApellido' +
        'M) as NombreCompleto,p.sDescripcion as categoria,'
      
        'tp.sIdTipoPersonal,tp.sDescripcion as TipoPersonal,cp.sdescripci' +
        'on as compania'
      'from tripulacion_listado tl'
      'inner join personal p'
      'on(p.sContrato=tl.sContrato and p.sIdPersonal=tl.sIdPersonal)'
      'inner join tiposdepersonal tp'
      'on(tp.sIdTipoPersonal=p.sIdTipoPersonal)'
      'inner join compersonal cp'
      'on(cp.sIdcompania=tl.sIdcompania)'
      'left join tripulaciondiaria_listado t'
      
        'on(tl.sContrato=t.sContrato and tl.sIdTripulacion=t.sIdTripulaci' +
        'on and t.didfecha=:Fecha)'
      'where tl.sContrato=:Contrato and t.sIdTripulacion is null')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
    Left = 304
    Top = 208
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
  end
  object dsListado: TDataSource
    AutoEdit = False
    DataSet = QListado
    Left = 360
    Top = 208
  end
  object QExterno: TZQuery
    Params = <>
    Left = 512
    Top = 208
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
    Left = 688
    Top = 160
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
    Left = 728
    Top = 160
  end
  object ds_categoria: TDataSource
    DataSet = zq_categoria
    OnDataChange = ds_categoriaDataChange
    Left = 728
    Top = 128
  end
  object zq_categoria: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select sIdTipoPersonal, sDescripcion from tiposdepersonal'
      'order by sDescripcion')
    Params = <>
    Left = 688
    Top = 128
  end
  object zq_compania: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'Select * from compersonal')
    Params = <>
    Left = 688
    Top = 96
  end
  object ds_compania: TDataSource
    DataSet = zq_compania
    Left = 728
    Top = 96
  end
  object ds_listadoper: TDataSource
    AutoEdit = False
    DataSet = zq_listadoper
    Left = 728
    Top = 64
  end
  object zq_listadoper: TZQuery
    Connection = connection.zConnection
    AfterScroll = zq_listadoperAfterScroll
    SQL.Strings = (
      'select * from tripulacion_listado')
    Params = <>
    Left = 688
    Top = 64
  end
end
