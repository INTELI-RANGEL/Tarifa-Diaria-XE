object frmRepExcelNew: TfrmRepExcelNew
  Left = 0
  Top = 0
  Caption = 'Consulta de Rendimiento'
  ClientHeight = 442
  ClientWidth = 740
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Scaled = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 401
    Width = 740
    Height = 41
    Align = alBottom
    Padding.Left = 6
    Padding.Top = 4
    Padding.Right = 6
    Padding.Bottom = 4
    TabOrder = 0
    object btnReporte: TButton
      Left = 129
      Top = 5
      Width = 75
      Height = 31
      Align = alLeft
      Caption = '&Reporte'
      Enabled = False
      TabOrder = 0
      OnClick = btnReporteClick
    end
    object btnCerrar: TButton
      Left = 658
      Top = 5
      Width = 75
      Height = 31
      Align = alRight
      Caption = '&Cerrar'
      TabOrder = 1
      OnClick = btnCerrarClick
    end
    object btnProcesar: TButton
      Left = 7
      Top = 5
      Width = 75
      Height = 31
      Align = alLeft
      Caption = '&Procesar'
      Enabled = False
      TabOrder = 2
      OnClick = btnProcesarClick
    end
    object Panel3: TPanel
      Left = 82
      Top = 5
      Width = 47
      Height = 31
      Align = alLeft
      BevelOuter = bvNone
      TabOrder = 3
    end
  end
  object pnlGenerar: TPanel
    Left = 0
    Top = 333
    Width = 740
    Height = 68
    Align = alBottom
    BevelOuter = bvLowered
    TabOrder = 1
    Visible = False
    object Label5: TLabel
      Left = 16
      Top = 13
      Width = 48
      Height = 13
      Caption = 'Actividad:'
    end
    object Label6: TLabel
      Left = 25
      Top = 37
      Width = 38
      Height = 13
      Caption = 'Partida:'
    end
    object cbActividad: TComboBox
      Left = 69
      Top = 10
      Width = 84
      Height = 21
      Style = csDropDownList
      Sorted = True
      TabOrder = 0
      OnChange = cbActividadChange
    end
    object cbPartida: TComboBox
      Left = 69
      Top = 34
      Width = 84
      Height = 21
      Style = csDropDownList
      Enabled = False
      Sorted = True
      TabOrder = 1
    end
  end
  object Panel5: TPanel
    Left = 0
    Top = 0
    Width = 740
    Height = 333
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object Panel2: TPanel
      Left = 0
      Top = 0
      Width = 740
      Height = 41
      Align = alTop
      TabOrder = 0
      object Label3: TLabel
        Left = 16
        Top = 14
        Width = 61
        Height = 13
        Caption = 'Fecha Inicio:'
      end
      object Label4: TLabel
        Left = 240
        Top = 14
        Width = 74
        Height = 13
        Caption = 'Fecha Termino:'
      end
      object FechaInicio: TDateTimePicker
        Left = 83
        Top = 11
        Width = 95
        Height = 21
        Date = 42550.733234328700000000
        Time = 42550.733234328700000000
        TabOrder = 0
      end
      object FechaFinal: TDateTimePicker
        Left = 320
        Top = 11
        Width = 95
        Height = 21
        Date = 42550.733234328700000000
        Time = 42550.733234328700000000
        TabOrder = 1
      end
      object btnBuscar: TButton
        Left = 464
        Top = 10
        Width = 75
        Height = 25
        Caption = '&Buscar'
        TabOrder = 2
        OnClick = btnBuscarClick
      end
    end
    object Panel4: TPanel
      Left = 0
      Top = 41
      Width = 740
      Height = 292
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object Panel6: TPanel
        Left = 0
        Top = 0
        Width = 367
        Height = 292
        Align = alLeft
        BevelOuter = bvNone
        Padding.Left = 4
        Padding.Right = 4
        TabOrder = 0
        object DBGrid1: TDBGrid
          Left = 4
          Top = 25
          Width = 359
          Height = 267
          Align = alClient
          DataSource = dsContratos
          ReadOnly = True
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Tahoma'
          TitleFont.Style = []
          Columns = <
            item
              Expanded = False
              FieldName = 'sContrato'
              Title.Caption = 'Contrato'
              Width = 74
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'mDescripcion'
              Title.Caption = 'Descripci'#243'n'
              Width = 265
              Visible = True
            end>
        end
        object Panel8: TPanel
          Left = 4
          Top = 0
          Width = 359
          Height = 25
          Align = alTop
          BevelOuter = bvNone
          TabOrder = 1
          object Label1: TLabel
            Left = 8
            Top = 7
            Width = 47
            Height = 13
            Caption = 'Contrato:'
          end
        end
      end
      object Panel7: TPanel
        Left = 367
        Top = 0
        Width = 373
        Height = 292
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 1
        object DBGrid2: TDBGrid
          Left = 0
          Top = 25
          Width = 373
          Height = 267
          Align = alClient
          DataSource = dsOrdenes
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Tahoma'
          TitleFont.Style = []
          Columns = <
            item
              Expanded = False
              FieldName = 'sNumeroOrden'
              Title.Caption = 'N'#250'mero Orden'
              Width = 79
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'sIdFolio'
              Title.Caption = 'Folio'
              Width = 134
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'mDescripcion'
              Title.Caption = 'Descripci'#243'n'
              Width = 163
              Visible = True
            end>
        end
        object Panel9: TPanel
          Left = 0
          Top = 0
          Width = 373
          Height = 25
          Align = alTop
          BevelOuter = bvNone
          TabOrder = 1
          object Label2: TLabel
            Left = 22
            Top = 7
            Width = 34
            Height = 13
            Caption = 'Orden:'
          end
        end
      end
    end
  end
  object roqReporte: TZReadOnlyQuery
    Connection = connection.zConnection
    AfterOpen = roqReporteAfterOpen
    AfterClose = roqReporteAfterClose
    SQL.Strings = (
      '/* MANO DE OBRA */'
      'SELECT'
      '  0 AS iIdOrdenTipoAnexo,'
      '  ot.sContrato,'
      '  ot.sNumeroOrden,'
      '  ot.sIdFolio,'
      '  ot.sDescripcionCorta AS mDescripcionCortaOrden,'
      '  ot.mDescripcion AS mDescripcionOrden,'
      '  ba.mDescripcion AS mDescripcionBitacora,'
      '  "MO" AS sTipoAnexo,'
      '  "" AS Frente,'
      '  ba.sIdClasificacion,'
      '  ba.dIdFecha,'
      '  mo.iItemOrden,'
      '  ba.sHoraInicio,'
      '  ba.sHoraFinal,'
      '  ba.sWbs,'
      '  ba.sNumeroActividad,'
      '  SUM(bp.dCantidad) AS dCantidad,'
      '  SUM(bp.dCantHH) AS dJornada,'
      '  ba.mDescripcion AS mDescripcionBitacoraActividades,'
      '  mo.dVentaMN,'
      '  mo.dVentaDLL,'
      '  mo.sIdPersonal AS sIdPartidaAnexo,'
      '  CAST(mo.sDescripcion AS CHAR) AS sTituloPartidaAnexo,'
      '  ROUND(SUM(bp.dCantHH) * mo.dVentaMN, 2) AS dCostoMN,'
      '  ROUND(SUM(bp.dCantHH) * mo.dVentaDLL, 2) AS dCostoDLL'
      ''
      'FROM'
      '  ordenesdetrabajo ot'
      ''
      'INNER JOIN'
      '  bitacoradeactividades ba'
      '    ON (ba.sContrato = ot.sContrato AND'
      '        ba.sNumeroOrden = ot.sNumeroOrden AND'
      '        ba.dCantidad > 0 AND'
      '        ba.dIdFecha BETWEEN :FechaInicio AND :FechaFinal)'
      ''
      'INNER JOIN'
      '  bitacoradepersonal bp'
      '    ON (bp.sContrato = ba.sContrato AND'
      '        bp.sNumeroOrden = ba.sNumeroOrden AND'
      '        bp.sNumeroActividad = ba.sNumeroActividad AND'
      '        bp.dIdFecha = ba.dIdFecha)'
      ''
      'INNER JOIN'
      '  personal mo'
      '    ON (mo.sIdPersonal = bp.sIdPersonal)'
      ''
      'WHERE'
      '  ot.sContrato = :sContrato AND'
      '  ot.sNumeroOrden = :sNumeroOrden'
      ''
      'GROUP BY'
      '  sContrato,'
      '  sNumeroOrden,'
      '  sTipoAnexo,'
      '  dIdFecha,'
      '  sNumeroActividad,'
      '  sIdPartidaAnexo'
      ''
      'UNION'
      ''
      '/* HERRAMIENTA Y EQUIPO */'
      'SELECT'
      '  1 AS iIdOrdenTipoAnexo,'
      '  ot.sContrato,'
      '  ot.sNumeroOrden,'
      '  ot.sIdFolio,'
      '  ot.sDescripcionCorta AS mDescripcionCortaOrden,'
      '  ot.mDescripcion AS mDescripcionOrden,'
      '  ba.mDescripcion AS mDescripcionBitacora,'
      '  "EQ" AS sTipoAnexo,'
      '  "" AS Frente,'
      '  ba.sIdClasificacion,'
      '  ba.dIdFecha,'
      '  eq.iItemOrden,'
      '  ba.sHoraInicio,'
      '  ba.sHoraFinal,'
      '  ba.sWbs,'
      '  ba.sNumeroActividad,'
      '  SUM(be.dCantidad) AS dCantidad,'
      '  SUM(be.dCantHH) AS dJornada,'
      '  ba.mDescripcion AS mDescripcionBitacoraActividades,'
      '  eq.dVentaMN,'
      '  eq.dVentaDLL,'
      '  eq.sIdEquipo AS sIdPartidaAnexo,'
      '  CAST(eq.sDescripcion AS CHAR) AS sTituloPartidaAnexo,'
      '  ROUND(SUM(be.dCantHH) * eq.dVentaMN, 2) AS dCostoMN,'
      '  ROUND(SUM(be.dCantHH) * eq.dVentaDLL, 2) AS dCostoDLL'
      ''
      'FROM'
      '  ordenesdetrabajo ot'
      ''
      'INNER JOIN'
      '  bitacoradeactividades ba'
      '    ON (ba.sContrato = ot.sContrato AND'
      '        ba.sNumeroOrden = ot.sNumeroOrden AND'
      '        ba.dCantidad > 0 AND'
      '        ba.dIdFecha BETWEEN :FechaInicio AND :FechaFinal)'
      ''
      'INNER JOIN'
      '  bitacoradeequipos be'
      '    ON (be.sContrato = ba.sContrato AND'
      '        be.sNumeroOrden = ba.sNumeroOrden AND'
      '        be.sNumeroActividad = ba.sNumeroActividad AND'
      '        be.dIdFecha = ba.dIdFecha)'
      ''
      'INNER JOIN'
      '  equipos eq'
      '    ON (eq.sIdEquipo = be.sIdEquipo)'
      ''
      'WHERE'
      '  ot.sContrato = :sContrato AND'
      '  ot.sNumeroOrden = :sNumeroOrden'
      ''
      'GROUP BY'
      '  sContrato,'
      '  sNumeroOrden,'
      '  sTipoAnexo,'
      '  dIdFecha,'
      '  sNumeroActividad,'
      '  sIdPartidaAnexo'
      ''
      'UNION'
      ''
      '/* PRECIO UNITARIO */'
      'SELECT'
      '  2 AS iIdOrdenTipoAnexo,'
      '  ot.sContrato,'
      '  ot.sNumeroOrden,'
      '  ot.sIdFolio,'
      '  ot.sDescripcionCorta AS mDescripcionCortaOrden,'
      '  ot.mDescripcion AS mDescripcionOrden,'
      '  ba.mDescripcion AS mDescripcionBitacora,'
      '  "PU" AS sTipoAnexo,'
      '  "" AS Frente,'
      '  ba.sIdClasificacion,'
      '  ba.dIdFecha,'
      '  axa.iItemOrden,'
      '  ba.sHoraInicio,'
      '  ba.sHoraFinal,'
      '  ba.sWbs,'
      '  ba.sNumeroActividad,'
      '  SUM(bm.dCantidad) AS dCantidad,'
      '  axa.sMedida AS dJornada,'
      '  ba.mDescripcion AS mDescripcionBitacoraActividades,'
      '  axa.dVentaMN,'
      '  axa.dVentaDLL,'
      '  axa.sNumeroActividad AS sIdPartidaAnexo,'
      '  CAST(axa.mDescripcion AS CHAR) AS sTituloPartidaAnexo,'
      '  ROUND(SUM(bm.dCantidad) * axa.dVentaMN, 2) AS dCostoMN,'
      '  ROUND(SUM(bm.dCantidad) * axa.dVentaDLL, 2) AS dCostoDLL'
      ''
      'FROM'
      '  ordenesdetrabajo ot'
      ''
      'INNER JOIN'
      '  bitacoradeactividades ba'
      '    ON (ba.sContrato = ot.sContrato AND'
      '        ba.sNumeroOrden = ot.sNumeroOrden AND'
      '        /*ba.dCantidad > 0 AND*/'
      '        ba.dIdFecha BETWEEN :FechaInicio AND :FechaFinal)'
      ''
      'INNER JOIN'
      '  bitacorademateriales bm'
      '    ON (bm.sContrato = ba.sContrato AND'
      '        bm.sNumeroOrden = ba.sNumeroOrden AND'
      '        bm.iIdDiario = ba.iIdDiario AND'
      '        bm.dIdFecha = ba.dIdFecha)'
      ''
      'INNER JOIN'
      '  actividadesxanexo axa'
      '    ON (axa.sNumeroActividad = bm.sIdMaterial)'
      ''
      'WHERE'
      '  ot.sContrato = :sContrato AND'
      '  ot.sNumeroOrden = :sNumeroOrden'
      ''
      'GROUP BY'
      '  sContrato,'
      '  sNumeroOrden,'
      '  sTipoAnexo,'
      '  dIdFecha,'
      '  sNumeroActividad,'
      '  sIdPartidaAnexo'
      ''
      'ORDER BY'
      '  iIdOrdenTipoAnexo,'
      '  dIdFecha DESC,'
      '  iItemOrden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'FechaInicio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'FechaFinal'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sContrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sNumeroOrden'
        ParamType = ptUnknown
      end>
    Left = 328
    Top = 104
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'FechaInicio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'FechaFinal'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sContrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sNumeroOrden'
        ParamType = ptUnknown
      end>
  end
  object SaveExcel: TSaveDialog
    Left = 512
    Top = 56
  end
  object roqContratos: TZReadOnlyQuery
    Connection = connection.zConnection
    AfterScroll = roqContratosAfterScroll
    SQL.Strings = (
      'SELECT'
      '  ot.sContrato,'
      '  CAST(cont.mDescripcion AS CHAR(500)) AS mDescripcion,'
      '  MIN(ba.dIdFecha) AS FechaInicio,'
      '  MAX(ba.dIdFecha) AS FechaFinal'
      ''
      'FROM'
      '  ordenesdetrabajo ot'
      ''
      'INNER JOIN'
      '  bitacoradeactividades ba'
      '    ON (ba.sContrato = ot.sContrato AND'
      '        ba.sNumeroOrden = ot.sNumeroOrden AND'
      '        ba.dCantidad > 0)'
      ''
      'INNER JOIN'
      '  contratos cont'
      '    ON (cont.sContrato = ot.sContrato)'
      ''
      'WHERE'
      '  ba.dIdFecha BETWEEN :FechaInicio AND :FechaFinal'
      ''
      'GROUP BY'
      '  sContrato'
      ''
      'ORDER BY'
      '  sContrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'FechaInicio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'FechaFinal'
        ParamType = ptUnknown
      end>
    Left = 136
    Top = 136
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'FechaInicio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'FechaFinal'
        ParamType = ptUnknown
      end>
  end
  object roqOrdenes: TZReadOnlyQuery
    Connection = connection.zConnection
    AfterOpen = roqOrdenesAfterOpen
    AfterClose = roqOrdenesAfterClose
    AfterRefresh = roqOrdenesAfterRefresh
    AfterScroll = roqOrdenesAfterScroll
    SQL.Strings = (
      'SELECT'
      '  ot.sContrato,'
      '  ot.sNumeroOrden,'
      '  ot.sIdFolio,'
      '  CAST(ot.mDescripcion AS CHAR(1000)) AS mDescripcion,'
      '  MIN(ba.dIdFecha) AS FechaInicio,'
      '  MAX(ba.dIdFecha) AS FechaFinal'
      ''
      'FROM'
      '  ordenesdetrabajo ot'
      ''
      'INNER JOIN'
      '  bitacoradeactividades ba'
      '    ON (ba.sContrato = ot.sContrato AND'
      '        ba.sNumeroOrden = ot.sNumeroOrden AND'
      '        ba.dCantidad > 0)'
      ''
      'WHERE'
      '  ba.dIdFecha BETWEEN :FechaInicio AND :FechaFinal'
      ''
      'GROUP BY'
      '  sContrato,'
      '  sNumeroOrden'
      ''
      'ORDER BY'
      '  sContrato,'
      '  sNumeroOrden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'FechaInicio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'FechaFinal'
        ParamType = ptUnknown
      end>
    Left = 536
    Top = 216
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'FechaInicio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'FechaFinal'
        ParamType = ptUnknown
      end>
  end
  object dsContratos: TDataSource
    DataSet = roqContratos
    Left = 264
    Top = 112
  end
  object dsOrdenes: TDataSource
    DataSet = roqOrdenes
    Left = 472
    Top = 224
  end
end
