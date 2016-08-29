object frmModuloReporteGerencial: TfrmModuloReporteGerencial
  Left = 0
  Top = 0
  BorderStyle = bsSingle
  Caption = 'Modulo Reporte Gerencial'
  ClientHeight = 385
  ClientWidth = 775
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Arial'
  Font.Style = [fsBold]
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Visible = True
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object pBarraAvance: TProgressBar
    Left = 0
    Top = 368
    Width = 775
    Height = 17
    Align = alBottom
    TabOrder = 0
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 775
    Height = 65
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object Label1: TLabel
      Left = 8
      Top = 8
      Width = 43
      Height = 16
      Caption = 'Fecha:'
    end
    object dIdFecha: TDateTimePicker
      Left = 80
      Top = 4
      Width = 186
      Height = 24
      Date = 41509.931811944440000000
      Time = 41509.931811944440000000
      TabOrder = 0
      OnChange = dIdFechaChange
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 65
    Width = 775
    Height = 239
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object NxPageControl1: TNxPageControl
      Left = 0
      Top = 0
      Width = 775
      Height = 239
      ActivePage = NxTabSheet5
      ActivePageIndex = 4
      Align = alClient
      TabOrder = 0
      BackgroundColor = clSilver
      BackgroundKind = bkSolid
      Margin = 0
      Spacing = 0
      TabHeight = 17
      object NxTabSheet1: TNxTabSheet
        Caption = 'Personal a bordo'
        PageIndex = 0
        ParentTabFont = False
        TabFont.Charset = DEFAULT_CHARSET
        TabFont.Color = clWindowText
        TabFont.Height = -13
        TabFont.Name = 'Arial'
        TabFont.Style = [fsBold]
        object Grid_Reportes: TDBGrid
          Left = 0
          Top = 0
          Width = 775
          Height = 218
          Hint = 'Aqu'#237' se reflejan los resultados de las consultas.'
          Align = alClient
          Color = 15138559
          Ctl3D = False
          DataSource = ds_personalabordo
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Arial'
          TitleFont.Style = [fsBold]
          Columns = <
            item
              Expanded = False
              FieldName = 'sPartida'
              Title.Caption = 'Partida'
              Width = 80
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'sDescripcion'
              Title.Caption = 'Descripci'#243'n'
              Width = 255
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'dCantidad'
              Title.Caption = 'Cantidad'
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'dCantidadaBordo'
              Title.Caption = 'A Bordo'
              Visible = True
            end>
        end
      end
      object NxTabSheet2: TNxTabSheet
        Caption = 'Personal faltante'
        PageIndex = 1
        ParentTabFont = False
        TabFont.Charset = DEFAULT_CHARSET
        TabFont.Color = clWindowText
        TabFont.Height = -13
        TabFont.Name = 'Arial'
        TabFont.Style = [fsBold]
        object DBGridPersonalFaltante: TDBGrid
          Left = 0
          Top = 0
          Width = 775
          Height = 218
          Hint = 'Aqu'#237' se reflejan los resultados de las consultas.'
          Align = alClient
          Color = 15138559
          Ctl3D = False
          DataSource = ds_PersonalFaltante
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Arial'
          TitleFont.Style = [fsBold]
          Columns = <
            item
              Expanded = False
              FieldName = 'sIdRecurso'
              Title.Caption = 'Partida'
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'sPersonal'
              Title.Caption = 'Descripci'#243'n'
              Width = 255
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'dCantidad'
              Title.Caption = 'Cantidad'
              Visible = True
            end>
        end
      end
      object NxTabSheet3: TNxTabSheet
        Caption = 'Personal Pendiente'
        PageIndex = 2
        ParentTabFont = False
        TabFont.Charset = DEFAULT_CHARSET
        TabFont.Color = clWindowText
        TabFont.Height = -13
        TabFont.Name = 'Arial'
        TabFont.Style = [fsBold]
        object DBGridPersonalPendiente: TDBGrid
          Left = 0
          Top = 0
          Width = 775
          Height = 218
          Hint = 'Aqu'#237' se reflejan los resultados de las consultas.'
          Align = alClient
          Color = 15138559
          Ctl3D = False
          DataSource = ds_PersonalPendiente
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Arial'
          TitleFont.Style = [fsBold]
          Columns = <
            item
              Expanded = False
              FieldName = 'sIdRecurso'
              Title.Caption = 'Partida'
              Width = 100
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'sPersonal'
              Title.Caption = 'Descripci'#243'n'
              Width = 255
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'dCantidad'
              Title.Caption = 'Cantidad'
              Visible = True
            end>
        end
      end
      object NxTabSheet4: TNxTabSheet
        Caption = 'Equipo a descuento'
        PageIndex = 3
        ParentTabFont = False
        TabFont.Charset = DEFAULT_CHARSET
        TabFont.Color = clWindowText
        TabFont.Height = -13
        TabFont.Name = 'Arial'
        TabFont.Style = [fsBold]
        object DBGridEquipoDescuento: TDBGrid
          Left = 0
          Top = 0
          Width = 775
          Height = 218
          Hint = 'Aqu'#237' se reflejan los resultados de las consultas.'
          Align = alClient
          Color = 15138559
          Ctl3D = False
          DataSource = ds_EquipoDescuento
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Arial'
          TitleFont.Style = [fsBold]
          Columns = <
            item
              Expanded = False
              FieldName = 'sIdRecurso'
              Title.Caption = 'Partida'
              Width = 100
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'sEquipo'
              Title.Caption = 'Descripci'#243'n'
              Width = 255
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'dCantidad'
              Title.Caption = 'Cantidad'
              Visible = True
            end>
        end
      end
      object NxTabSheet5: TNxTabSheet
        Caption = 'Equipo Fuera de Operaci'#243'n'
        PageIndex = 4
        ParentTabFont = False
        TabFont.Charset = DEFAULT_CHARSET
        TabFont.Color = clWindowText
        TabFont.Height = -13
        TabFont.Name = 'Arial'
        TabFont.Style = [fsBold]
        object DBGridFO: TDBGrid
          Left = 0
          Top = 0
          Width = 775
          Height = 218
          Hint = 'Aqu'#237' se reflejan los resultados de las consultas.'
          Align = alClient
          Color = 15138559
          Ctl3D = False
          DataSource = ds_EquiposFueradeOperacion
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Arial'
          TitleFont.Style = [fsBold]
          Columns = <
            item
              Expanded = False
              FieldName = 'sIdRecurso'
              Title.Caption = 'Partida'
              Width = 100
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'sEquipo'
              Title.Caption = 'Descripci'#243'n'
              Width = 255
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'dCantidad'
              Title.Caption = 'Cantidad'
              Visible = True
            end>
        end
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 304
    Width = 775
    Height = 64
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 3
  end
  object Panel: tNewGroupBox
    Left = 182
    Top = 86
    Width = 227
    Height = 133
    Align = alCustom
    Caption = 'Coincidencias Encontradas'
    Color = clSilver
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindow
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
    TabOrder = 4
    TabStop = True
    Visible = False
    HighLightColor = clWindowText
    ShadowColor = clWindowText
    Arc = 15
    Bevel = bnRaisedLine
    Title.Offset = 0
    Title.Width = 700
    Title.HighLightColor = cl3DDkShadow
    Title.ShadowColor = clWindowText
    Title.Arc = 15
    Title.Shape = tsRect
    Title.Separation = 2
    Title.Bevel = bnRaisedLine
    Title.Border = True
    Title.TextAlign = ttLeft
    Title.Align = taTop
    Title.Height = 20
    Title.BkColor = clGray
    TransparentMode = False
    Border = True
    Shape = tsRect
    object ListaObjeto: TRxDBGrid
      Left = 3
      Top = 27
      Width = 220
      Height = 103
      Hint = 'Doble Click para Seleccionar'
      Align = alCustom
      Anchors = [akLeft, akTop, akRight, akBottom]
      Color = 15138559
      Ctl3D = False
      DataSource = ds_buscaobjeto
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      Options = [dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindow
      TitleFont.Height = -12
      TitleFont.Name = 'Arial'
      TitleFont.Style = [fsBold]
      OnExit = ListaObjetoExit
      OnKeyPress = ListaObjetoKeyPress
      TitleButtons = True
      Columns = <
        item
          Expanded = False
          FieldName = 'sDescripcion'
          Width = 680
          Visible = True
        end>
    end
  end
  object QryFolios: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'SELECT '
      '  ot.* '
      'FROM ordenesdetrabajo AS ot '
      '  INNER JOIN contratos AS c '
      '       ON (ot.sContrato=c.sContrato) '
      '  INNER JOIN bitacoradeactividades AS ba '
      
        '       ON (ba.sContrato=c.sContrato AND ba.sNumeroOrden=ot.sNume' +
        'roOrden) '
      '  INNER JOIN tiposdemovimiento AS tm '
      
        '       ON (tm.sContrato= :ContratoBarco AND tm.sIdTipoMovimiento' +
        '=ba.sIdTipoMovimiento AND tm.sClasificacion="Tarifa Diaria") '
      'WHERE '
      '  (c.sContrato= :Contrato) '
      '  AND ba.dIdFecha= :Fecha '
      'GROUP BY '
      '  ot.sContrato, '
      '  ot.sNumeroorden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'ContratoBarco'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    Left = 624
    Top = 136
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'ContratoBarco'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    object QryFoliossContrato: TStringField
      FieldName = 'sContrato'
      Required = True
      Size = 15
    end
    object QryFoliossIdFolio: TStringField
      FieldName = 'sIdFolio'
      Required = True
      Size = 35
    end
    object QryFoliossNumeroOrden: TStringField
      FieldName = 'sNumeroOrden'
      Required = True
      Size = 35
    end
    object QryFoliossDescripcionCorta: TStringField
      FieldName = 'sDescripcionCorta'
      Required = True
      Size = 50
    end
    object QryFoliossOficioAutorizacion: TStringField
      FieldName = 'sOficioAutorizacion'
      Required = True
      Size = 50
    end
    object QryFoliosmDescripcion: TMemoField
      FieldName = 'mDescripcion'
      Required = True
      BlobType = ftMemo
    end
    object QryFoliossIdTipoOrden: TStringField
      FieldName = 'sIdTipoOrden'
      Required = True
      Size = 4
    end
    object QryFoliossApoyo: TStringField
      FieldName = 'sApoyo'
      Required = True
      Size = 10
    end
    object QryFoliossIdPlataforma: TStringField
      FieldName = 'sIdPlataforma'
      Size = 50
    end
    object QryFoliossIdPernocta: TStringField
      FieldName = 'sIdPernocta'
      Required = True
      Size = 10
    end
    object QryFoliosdFiProgramado: TDateField
      FieldName = 'dFiProgramado'
      Required = True
    end
    object QryFoliosdFfProgramado: TDateField
      FieldName = 'dFfProgramado'
      Required = True
    end
    object QryFolioscIdStatus: TStringField
      FieldName = 'cIdStatus'
      Required = True
      Size = 1
    end
    object QryFoliosmComentarios: TMemoField
      FieldName = 'mComentarios'
      BlobType = ftMemo
    end
    object QryFoliossFormato: TStringField
      FieldName = 'sFormato'
      Required = True
      Size = 30
    end
    object QryFoliosiConsecutivo: TIntegerField
      FieldName = 'iConsecutivo'
      Required = True
    end
    object QryFoliosiConsecutivoTierra: TIntegerField
      FieldName = 'iConsecutivoTierra'
      Required = True
    end
    object QryFoliosiJornada: TIntegerField
      FieldName = 'iJornada'
      Required = True
    end
    object QryFolioslGeneraAnexo: TStringField
      FieldName = 'lGeneraAnexo'
      Required = True
      Size = 2
    end
    object QryFolioslGeneraConsumibles: TStringField
      FieldName = 'lGeneraConsumibles'
      Required = True
      Size = 2
    end
    object QryFolioslGeneraPersonal: TStringField
      FieldName = 'lGeneraPersonal'
      Required = True
      Size = 2
    end
    object QryFolioslGeneraEquipo: TStringField
      FieldName = 'lGeneraEquipo'
      Required = True
      Size = 2
    end
    object QryFoliossDepsolicitante: TStringField
      FieldName = 'sDepsolicitante'
      Size = 45
    end
    object QryFoliosdFechaInicioT: TDateField
      FieldName = 'dFechaInicioT'
    end
    object QryFoliosdFechaSitioM: TDateField
      FieldName = 'dFechaSitioM'
    end
    object QryFoliossEquipo: TStringField
      FieldName = 'sEquipo'
    end
    object QryFoliossPozo: TStringField
      FieldName = 'sPozo'
    end
    object QryFoliosdFechaElaboracion: TDateField
      FieldName = 'dFechaElaboracion'
    end
    object QryFoliossPuestoPep: TStringField
      FieldName = 'sPuestoPep'
      Size = 60
    end
    object QryFoliossFirmantePep: TStringField
      FieldName = 'sFirmantePep'
      Size = 60
    end
    object QryFoliossPuestocia: TStringField
      FieldName = 'sPuestocia'
      Size = 60
    end
    object QryFoliossFirmantecia: TStringField
      FieldName = 'sFirmantecia'
      Size = 60
    end
    object QryFolioslMostrarAvanceProgramado: TStringField
      FieldName = 'lMostrarAvanceProgramado'
      Size = 2
    end
    object QryFoliossTipoOrden: TStringField
      FieldName = 'sTipoOrden'
    end
    object QryFoliosbAvanceFrente: TStringField
      FieldName = 'bAvanceFrente'
      Required = True
      Size = 2
    end
    object QryFoliosbAvanceContrato: TStringField
      FieldName = 'bAvanceContrato'
      Required = True
      Size = 2
    end
    object QryFoliosbComentarios: TStringField
      FieldName = 'bComentarios'
      Required = True
      Size = 2
    end
    object QryFoliosbPermisos: TStringField
      FieldName = 'bPermisos'
      Required = True
      Size = 2
    end
    object QryFoliosbTipoAdmon: TStringField
      FieldName = 'bTipoAdmon'
      Required = True
      Size = 2
    end
    object QryFoliosbCostaFuera: TStringField
      FieldName = 'bCostaFuera'
      Size = 2
    end
    object QryFoliossTipoPrograma: TStringField
      FieldName = 'sTipoPrograma'
      Required = True
      Size = 21
    end
    object QryFoliossTipoImpresionActividad: TStringField
      FieldName = 'sTipoImpresionActividad'
      Required = True
      Size = 2
    end
    object QryFoliossTipoAvanceAdmon: TStringField
      FieldName = 'sTipoAvanceAdmon'
      Required = True
      Size = 2
    end
    object QryFoliosiDecimales: TIntegerField
      FieldName = 'iDecimales'
      Required = True
    end
    object QryFoliosiNiveles: TIntegerField
      FieldName = 'iNiveles'
      Required = True
    end
    object QryFolioslImprimeProgramado: TStringField
      FieldName = 'lImprimeProgramado'
      Required = True
      Size = 2
    end
    object QryFolioslImprimeFisico: TStringField
      FieldName = 'lImprimeFisico'
      Required = True
      Size = 2
    end
    object QryFolioslImprimePlaticas: TStringField
      FieldName = 'lImprimePlaticas'
      Required = True
      Size = 2
    end
    object QryFolioslImprimePersonalTM: TStringField
      FieldName = 'lImprimePersonalTM'
      Required = True
      Size = 2
    end
    object QryFolioslPersonalxPartida: TStringField
      FieldName = 'lPersonalxPartida'
      Required = True
      Size = 2
    end
    object QryFolioslImprimeFases: TStringField
      FieldName = 'lImprimeFases'
      Required = True
      Size = 2
    end
    object QryFolioslMostrarPartidasReportes: TStringField
      FieldName = 'lMostrarPartidasReportes'
      Required = True
      Size = 10
    end
    object QryFolioslMostrarPartidasGeneradores: TStringField
      FieldName = 'lMostrarPartidasGeneradores'
      Required = True
      Size = 10
    end
    object QryFoliosdFechaIniPReportes: TDateField
      FieldName = 'dFechaIniPReportes'
    end
    object QryFoliosdFechaFinPReportes: TDateField
      FieldName = 'dFechaFinPReportes'
    end
    object QryFoliosdFechaIniPGeneradores: TDateField
      FieldName = 'dFechaIniPGeneradores'
    end
    object QryFoliosdFechaFinPGeneradores: TDateField
      FieldName = 'dFechaFinPGeneradores'
    end
    object QryFolioslEstado: TStringField
      FieldName = 'lEstado'
      Size = 7
    end
  end
  object dsFolios: TDataSource
    DataSet = QryFolios
    Left = 656
    Top = 136
  end
  object zq_personalabordo: TZQuery
    Connection = connection.zConnection
    AfterInsert = zq_personalabordoAfterInsert
    SQL.Strings = (
      
        'SELECT * FROM gerencial_abordo WHERE sContrato = :Contrato AND d' +
        'IdFecha = :Fecha')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    Left = 16
    Top = 304
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    object zq_personalabordoiId: TIntegerField
      FieldName = 'iId'
      Required = True
    end
    object zq_personalabordodIdFecha: TDateTimeField
      FieldName = 'dIdFecha'
    end
    object zq_personalabordosContrato: TStringField
      FieldName = 'sContrato'
      Size = 50
    end
    object zq_personalabordosPartida: TStringField
      FieldName = 'sPartida'
    end
    object zq_personalabordosDescripcion: TStringField
      FieldName = 'sDescripcion'
      Size = 255
    end
    object zq_personalabordodCantidad: TFloatField
      FieldName = 'dCantidad'
    end
    object zq_personalabordodCantidadaBordo: TFloatField
      FieldName = 'dCantidadaBordo'
    end
  end
  object ds_personalabordo: TDataSource
    DataSet = zq_personalabordo
    Left = 48
    Top = 304
  end
  object zq_PersonalFaltante: TZQuery
    Connection = connection.zConnection
    AfterInsert = zq_PersonalFaltanteAfterInsert
    SQL.Strings = (
      
        'SELECT * FROM gerencial_recursosfaltantes WHERE sContrato = :Con' +
        'trato AND dIdFecha = :Fecha AND sTipoRecurso = "Personal" AND sT' +
        'ipoFaltante = "Faltante";')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    Left = 152
    Top = 304
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    object zq_PersonalFaltanteiId: TIntegerField
      FieldName = 'iId'
      Required = True
    end
    object zq_PersonalFaltantedIdFecha: TDateField
      FieldName = 'dIdFecha'
    end
    object zq_PersonalFaltantesContrato: TStringField
      FieldName = 'sContrato'
      Size = 30
    end
    object zq_PersonalFaltantesIdRecurso: TStringField
      FieldName = 'sIdRecurso'
      OnChange = zq_PersonalFaltantesIdRecursoChange
    end
    object zq_PersonalFaltantedCantidad: TFloatField
      FieldName = 'dCantidad'
    end
    object zq_PersonalFaltantesTipoRecurso: TStringField
      FieldName = 'sTipoRecurso'
    end
    object zq_PersonalFaltantesTipoFaltante: TStringField
      FieldName = 'sTipoFaltante'
    end
    object zq_PersonalFaltantesPersonal: TStringField
      FieldKind = fkLookup
      FieldName = 'sPersonal'
      LookupDataSet = zq_Personal
      LookupKeyFields = 'sIdPersonal'
      LookupResultField = 'sDescripcion'
      KeyFields = 'sIdRecurso'
      Size = 200
      Lookup = True
    end
  end
  object ds_PersonalFaltante: TDataSource
    DataSet = zq_PersonalFaltante
    Left = 184
    Top = 304
  end
  object zq_Personal: TZQuery
    Connection = connection.zConnection
    AfterInsert = zq_personalabordoAfterInsert
    SQL.Strings = (
      'SELECT * FROM personal WHERE sContrato = :ContratoBarco')
    Params = <
      item
        DataType = ftUnknown
        Name = 'ContratoBarco'
        ParamType = ptUnknown
      end>
    Left = 152
    Top = 344
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'ContratoBarco'
        ParamType = ptUnknown
      end>
    object zq_PersonalsContrato: TStringField
      FieldName = 'sContrato'
      Required = True
      Size = 15
    end
    object zq_PersonalsIdPersonal: TStringField
      FieldName = 'sIdPersonal'
      Required = True
      Size = 25
    end
    object zq_PersonaliItemOrden: TIntegerField
      FieldName = 'iItemOrden'
      Required = True
    end
    object zq_PersonalsDescripcion: TStringField
      FieldName = 'sDescripcion'
      Required = True
      Size = 250
    end
    object zq_PersonalsIdTipoPersonal: TStringField
      FieldName = 'sIdTipoPersonal'
      Required = True
      Size = 4
    end
    object zq_PersonalsMedida: TStringField
      FieldName = 'sMedida'
      Required = True
      Size = 8
    end
    object zq_PersonaldCantidad: TFloatField
      FieldName = 'dCantidad'
      Required = True
    end
    object zq_PersonaldCostoMN: TFloatField
      FieldName = 'dCostoMN'
      Required = True
    end
    object zq_PersonaldCostoDLL: TFloatField
      FieldName = 'dCostoDLL'
      Required = True
    end
    object zq_PersonaldVentaMN: TFloatField
      FieldName = 'dVentaMN'
      Required = True
    end
    object zq_PersonaldVentaDLL: TFloatField
      FieldName = 'dVentaDLL'
      Required = True
    end
    object zq_PersonaldFechaInicio: TDateField
      FieldName = 'dFechaInicio'
      Required = True
    end
    object zq_PersonaldFechaFinal: TDateField
      FieldName = 'dFechaFinal'
      Required = True
    end
    object zq_PersonallProrrateo: TStringField
      FieldName = 'lProrrateo'
      Required = True
      Size = 2
    end
    object zq_PersonallCobro: TStringField
      FieldName = 'lCobro'
      Required = True
      Size = 2
    end
    object zq_PersonallImprime: TStringField
      FieldName = 'lImprime'
      Required = True
      Size = 2
    end
    object zq_PersonallAplicaTM: TStringField
      FieldName = 'lAplicaTM'
      Required = True
      Size = 2
    end
    object zq_PersonaliJornada: TIntegerField
      FieldName = 'iJornada'
      Required = True
    end
    object zq_PersonallDistribuye: TStringField
      FieldName = 'lDistribuye'
      Required = True
      Size = 2
    end
    object zq_PersonallPernocta: TStringField
      FieldName = 'lPernocta'
      Required = True
      Size = 2
    end
    object zq_PersonalsAgrupaPersonal: TStringField
      FieldName = 'sAgrupaPersonal'
      Size = 25
    end
    object zq_PersonallTotalizarPernocta: TStringField
      FieldName = 'lTotalizarPernocta'
      Required = True
      Size = 2
    end
    object zq_PersonallAplicaGerencial: TStringField
      FieldName = 'lAplicaGerencial'
      Required = True
      Size = 2
    end
    object zq_PersonaliId_AgrupadorPersonal: TIntegerField
      FieldName = 'iId_AgrupadorPersonal'
    end
    object zq_PersonallSumaSolicitado: TStringField
      FieldName = 'lSumaSolicitado'
      Size = 2
    end
  end
  object ds_Personal: TDataSource
    DataSet = zq_Personal
    Left = 184
    Top = 344
  end
  object BuscaObjeto: TZReadOnlyQuery
    Connection = connection.zConnection
    Params = <>
    Left = 249
    Top = 187
  end
  object ds_buscaobjeto: TDataSource
    AutoEdit = False
    DataSet = BuscaObjeto
    Left = 280
    Top = 187
  end
  object zq_PersonalPendiente: TZQuery
    Connection = connection.zConnection
    AfterInsert = zq_PersonalPendienteAfterInsert
    SQL.Strings = (
      
        'SELECT * FROM gerencial_recursosfaltantes WHERE sContrato = :Con' +
        'trato AND dIdFecha = :Fecha AND sTipoRecurso = "Personal" AND sT' +
        'ipoFaltante = "Pendiente";')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    Left = 272
    Top = 304
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    object zq_PersonalPendienteiId: TIntegerField
      FieldName = 'iId'
      Required = True
    end
    object zq_PersonalPendientedIdFecha: TDateField
      FieldName = 'dIdFecha'
    end
    object zq_PersonalPendientesContrato: TStringField
      FieldName = 'sContrato'
      Size = 30
    end
    object zq_PersonalPendientesIdRecurso: TStringField
      FieldName = 'sIdRecurso'
      OnChange = zq_PersonalPendientesIdRecursoChange
    end
    object zq_PersonalPendientedCantidad: TFloatField
      FieldName = 'dCantidad'
    end
    object zq_PersonalPendientesTipoRecurso: TStringField
      FieldName = 'sTipoRecurso'
    end
    object zq_PersonalPendientesTipoFaltante: TStringField
      FieldName = 'sTipoFaltante'
    end
    object zq_PersonalPendientesPersonal: TStringField
      FieldKind = fkLookup
      FieldName = 'sPersonal'
      LookupDataSet = zq_Personal
      LookupKeyFields = 'sIdPersonal'
      LookupResultField = 'sDescripcion'
      KeyFields = 'sIdRecurso'
      Size = 100
      Lookup = True
    end
  end
  object ds_PersonalPendiente: TDataSource
    DataSet = zq_PersonalPendiente
    Left = 304
    Top = 304
  end
  object zq_EquipoDescuento: TZQuery
    Connection = connection.zConnection
    AfterInsert = zq_EquipoDescuentoAfterInsert
    SQL.Strings = (
      
        'SELECT * FROM gerencial_recursosfaltantes WHERE sContrato = :Con' +
        'trato AND dIdFecha = :Fecha AND sTipoRecurso = "Equipo" AND sTip' +
        'oFaltante = "Descuento";')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    Left = 408
    Top = 304
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    object zq_EquipoDescuentoiId: TIntegerField
      FieldName = 'iId'
      Required = True
    end
    object zq_EquipoDescuentodIdFecha: TDateField
      FieldName = 'dIdFecha'
    end
    object zq_EquipoDescuentosContrato: TStringField
      FieldName = 'sContrato'
      Size = 30
    end
    object zq_EquipoDescuentosIdRecurso: TStringField
      FieldName = 'sIdRecurso'
      OnChange = zq_EquipoDescuentosIdRecursoChange
    end
    object zq_EquipoDescuentodCantidad: TFloatField
      FieldName = 'dCantidad'
    end
    object zq_EquipoDescuentosTipoRecurso: TStringField
      FieldName = 'sTipoRecurso'
    end
    object zq_EquipoDescuentosTipoFaltante: TStringField
      FieldName = 'sTipoFaltante'
    end
    object zq_EquipoDescuentosEquipo: TStringField
      FieldKind = fkLookup
      FieldName = 'sEquipo'
      LookupDataSet = zq_Equipos
      LookupKeyFields = 'sIdEquipo'
      LookupResultField = 'sDescripcion'
      KeyFields = 'sIdRecurso'
      Size = 200
      Lookup = True
    end
  end
  object ds_EquipoDescuento: TDataSource
    DataSet = zq_EquipoDescuento
    Left = 440
    Top = 304
  end
  object zq_Equipos: TZQuery
    Connection = connection.zConnection
    AfterInsert = zq_personalabordoAfterInsert
    SQL.Strings = (
      'SELECT * FROM equipos WHERE sContrato = :ContratoBarco')
    Params = <
      item
        DataType = ftUnknown
        Name = 'ContratoBarco'
        ParamType = ptUnknown
      end>
    Left = 408
    Top = 344
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'ContratoBarco'
        ParamType = ptUnknown
      end>
    object zq_EquipossContrato: TStringField
      FieldName = 'sContrato'
      Required = True
      Size = 15
    end
    object zq_EquipossIdEquipo: TStringField
      FieldName = 'sIdEquipo'
      Required = True
      Size = 25
    end
    object zq_EquiposiItemOrden: TIntegerField
      FieldName = 'iItemOrden'
      Required = True
    end
    object zq_EquipossDescripcion: TStringField
      FieldName = 'sDescripcion'
      Required = True
      Size = 750
    end
    object zq_EquipossIdTipoEquipo: TStringField
      FieldName = 'sIdTipoEquipo'
      Required = True
      Size = 4
    end
    object zq_EquipossMedida: TStringField
      FieldName = 'sMedida'
      Required = True
      Size = 8
    end
    object zq_EquiposdCantidad: TFloatField
      FieldName = 'dCantidad'
      Required = True
    end
    object zq_EquiposdCostoMN: TFloatField
      FieldName = 'dCostoMN'
      Required = True
    end
    object zq_EquiposdCostoDLL: TFloatField
      FieldName = 'dCostoDLL'
      Required = True
    end
    object zq_EquiposdVentaMN: TFloatField
      FieldName = 'dVentaMN'
      Required = True
    end
    object zq_EquiposdVentaDLL: TFloatField
      FieldName = 'dVentaDLL'
      Required = True
    end
    object zq_EquiposdFechaInicio: TDateField
      FieldName = 'dFechaInicio'
      Required = True
    end
    object zq_EquiposdFechaFinal: TDateField
      FieldName = 'dFechaFinal'
      Required = True
    end
    object zq_EquiposlProrrateo: TStringField
      FieldName = 'lProrrateo'
      Required = True
      Size = 2
    end
    object zq_EquiposlCobro: TStringField
      FieldName = 'lCobro'
      Required = True
      Size = 2
    end
    object zq_EquiposlImprime: TStringField
      FieldName = 'lImprime'
      Required = True
      Size = 2
    end
    object zq_EquiposiJornada: TIntegerField
      FieldName = 'iJornada'
      Required = True
    end
    object zq_EquiposlDistribuye: TStringField
      FieldName = 'lDistribuye'
      Required = True
      Size = 2
    end
    object zq_EquiposlCuadraEquipo: TStringField
      FieldName = 'lCuadraEquipo'
      Required = True
      Size = 2
    end
    object zq_EquiposlAplicaDiesel: TStringField
      FieldName = 'lAplicaDiesel'
      Size = 2
    end
    object zq_EquipossDescripcionDiesel: TStringField
      FieldName = 'sDescripcionDiesel'
      Size = 100
    end
    object zq_EquiposlSumaSolicitado: TStringField
      FieldName = 'lSumaSolicitado'
      Size = 2
    end
  end
  object ds_Equipos: TDataSource
    DataSet = zq_Equipos
    Left = 440
    Top = 344
  end
  object zq_EquiposFueradeOperacion: TZQuery
    Connection = connection.zConnection
    AfterInsert = zq_EquiposFueradeOperacionAfterInsert
    SQL.Strings = (
      
        'SELECT * FROM gerencial_recursosfaltantes WHERE sContrato = :Con' +
        'trato AND dIdFecha = :Fecha AND sTipoRecurso = "Equipo" AND sTip' +
        'oFaltante = "FO";')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    Left = 544
    Top = 304
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    object zq_EquiposFueradeOperacioniId: TIntegerField
      FieldName = 'iId'
      Required = True
    end
    object zq_EquiposFueradeOperaciondIdFecha: TDateField
      FieldName = 'dIdFecha'
    end
    object zq_EquiposFueradeOperacionsContrato: TStringField
      FieldName = 'sContrato'
      Size = 30
    end
    object zq_EquiposFueradeOperacionsIdRecurso: TStringField
      FieldName = 'sIdRecurso'
      OnChange = zq_EquiposFueradeOperacionsIdRecursoChange
    end
    object zq_EquiposFueradeOperaciondCantidad: TFloatField
      FieldName = 'dCantidad'
    end
    object zq_EquiposFueradeOperacionsTipoRecurso: TStringField
      FieldName = 'sTipoRecurso'
    end
    object zq_EquiposFueradeOperacionsTipoFaltante: TStringField
      FieldName = 'sTipoFaltante'
    end
    object zq_EquiposFueradeOperacionsEquipo: TStringField
      FieldKind = fkLookup
      FieldName = 'sEquipo'
      LookupDataSet = zq_Equipos
      LookupKeyFields = 'sIdEquipo'
      LookupResultField = 'sDescripcion'
      KeyFields = 'sIdRecurso'
      Size = 200
      Lookup = True
    end
  end
  object ds_EquiposFueradeOperacion: TDataSource
    DataSet = zq_EquiposFueradeOperacion
    Left = 576
    Top = 304
  end
end
