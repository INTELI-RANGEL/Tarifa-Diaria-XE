object frmModuloAdmonPersonal: TfrmModuloAdmonPersonal
  Left = 0
  Top = 0
  BorderStyle = bsSingle
  Caption = 'M'#243'dulo de Resumen de personal'
  ClientHeight = 385
  ClientWidth = 761
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
    Width = 761
    Height = 17
    Align = alBottom
    TabOrder = 0
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 761
    Height = 65
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    DesignSize = (
      761
      65)
    object Label1: TLabel
      Left = 8
      Top = 8
      Width = 43
      Height = 16
      Caption = 'Fecha:'
    end
    object Label2: TLabel
      Left = 279
      Top = 7
      Width = 36
      Height = 16
      Caption = 'Folio:'
    end
    object Label3: TLabel
      Left = 8
      Top = 33
      Width = 66
      Height = 16
      Caption = 'Categor'#237'a:'
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
    object ComboBox: TComboBox
      Left = 321
      Top = 4
      Width = 224
      Height = 24
      Style = csDropDownList
      ItemHeight = 0
      TabOrder = 1
      OnChange = ComboBoxChange
    end
    object cmbCategoria: TComboBox
      Left = 80
      Top = 30
      Width = 465
      Height = 24
      Style = csDropDownList
      ItemHeight = 0
      TabOrder = 2
      OnChange = cmbCategoriaChange
    end
    object Button2: TButton
      Left = 601
      Top = 13
      Width = 153
      Height = 41
      Anchors = [akRight, akBottom]
      Caption = 'Traer D'#237'a anterior'
      TabOrder = 3
      OnClick = Button2Click
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 65
    Width = 761
    Height = 239
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object NextGrid: TNextGrid
      Left = 0
      Top = 0
      Width = 761
      Height = 239
      Align = alClient
      AppearanceOptions = [aoAlphaBlendedSelection, aoHideSelection, aoHighlightSlideCells]
      HighlightedTextColor = clBlack
      Options = [goFooter, goHeader]
      TabOrder = 0
      TabStop = True
      OnAfterEdit = NextGridAfterEdit
      object NxCheckBoxColumn1: TNxCheckBoxColumn
        DefaultWidth = 30
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing]
        ParentFont = False
        Position = 0
        SortType = stBoolean
        Width = 30
      end
      object NxTextColumn1: TNxTextColumn
        DefaultWidth = 90
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Arial'
        Font.Style = []
        Header.Caption = 'Categor'#237'a'
        ParentFont = False
        Position = 1
        SortType = stAlphabetic
        Width = 90
      end
      object NxTextColumn2: TNxTextColumn
        DefaultWidth = 300
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Arial'
        Font.Style = []
        Header.Caption = 'Descripci'#243'n'
        ParentFont = False
        Position = 2
        SortType = stAlphabetic
        Width = 300
      end
      object NxNumberColumn1: TNxNumberColumn
        Color = 16774120
        DefaultValue = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Arial'
        Font.Style = []
        Footer.FormulaKind = fkSum
        Footer.FormatMask = '0.00'
        Footer.FormatMaskKind = mkFloat
        Header.Caption = 'Cantidad'
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing]
        ParentFont = False
        Position = 3
        SortType = stNumeric
        Increment = 1.000000000000000000
        Precision = 0
      end
      object NxComboBoxColumn1: TNxComboBoxColumn
        DefaultWidth = 200
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Arial'
        Font.Style = []
        Header.Caption = 'Pernocta'
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing, coShowTextFitHint]
        ParentFont = False
        Position = 4
        SortType = stAlphabetic
        Width = 200
        OnSelect = NxComboBoxColumn1Select
        Style = cbsDropDownList
      end
      object NxTextColumn3: TNxTextColumn
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        Position = 5
        SortType = stAlphabetic
        Visible = False
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 304
    Width = 761
    Height = 64
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 3
    DesignSize = (
      761
      64)
    object Label4: TLabel
      Left = 8
      Top = 37
      Width = 115
      Height = 16
      Caption = 'Ajustar Pernoctas:'
    end
    object Label5: TLabel
      Left = 8
      Top = 12
      Width = 71
      Height = 16
      Caption = 'Pernoctas: '
    end
    object LabelPernoctas: TLabel
      Left = 129
      Top = 12
      Width = 25
      Height = 16
      Caption = '0.00'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
    end
    object Button1: TButton
      Left = 601
      Top = 22
      Width = 153
      Height = 36
      Anchors = [akRight, akBottom]
      Caption = 'Guardar Cambios'
      TabOrder = 0
      OnClick = Button1Click
    end
    object NxNumberEdit1: TNxNumberEdit
      Left = 129
      Top = 34
      Width = 121
      Height = 24
      TabOrder = 1
      Text = '0.00000000'
      Precision = 8
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
    Left = 248
    Top = 8
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
    Left = 280
    Top = 8
  end
end
