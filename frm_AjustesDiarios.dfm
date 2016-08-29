object frmAjustesDiarios: TfrmAjustesDiarios
  Left = 0
  Top = 0
  Caption = 'Ajustes Diarios'
  ClientHeight = 295
  ClientWidth = 1003
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1003
    Height = 254
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object Splitter1: TSplitter
      Left = 329
      Top = 0
      Height = 254
      ExplicitLeft = 296
      ExplicitTop = 80
      ExplicitHeight = 100
    end
    object Panel3: TPanel
      Left = 0
      Top = 0
      Width = 329
      Height = 254
      Align = alLeft
      BevelOuter = bvNone
      Padding.Left = 4
      Padding.Right = 4
      TabOrder = 0
      object DBGrid1: TDBGrid
        Left = 4
        Top = 0
        Width = 321
        Height = 254
        Align = alClient
        DataSource = dsOrdenes
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
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
            Width = 79
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'sIdFolio'
            Width = 134
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'mDescripcion'
            Width = 163
            Visible = True
          end>
      end
    end
    object Panel4: TPanel
      Left = 332
      Top = 0
      Width = 671
      Height = 254
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object Panel5: TPanel
        Left = 0
        Top = 0
        Width = 671
        Height = 121
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 0
        object Label1: TLabel
          Left = 16
          Top = 51
          Width = 48
          Height = 13
          Caption = 'Actividad:'
        end
        object Label2: TLabel
          Left = 20
          Top = 97
          Width = 44
          Height = 13
          Caption = 'Horarios:'
        end
        object rgAnexo: TRadioGroup
          Left = 0
          Top = 0
          Width = 671
          Height = 41
          Align = alTop
          Columns = 3
          ItemIndex = 0
          Items.Strings = (
            'Mano de Obra'
            'Equipo de Trabajo'
            'Pernoctas')
          TabOrder = 0
          OnClick = rgAnexoClick
        end
        object cbActividad: TComboBox
          Left = 70
          Top = 47
          Width = 107
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 1
          OnChange = cbActividadChange
        end
        object cbHorarios: TComboBox
          Left = 70
          Top = 94
          Width = 145
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 2
          OnChange = cbHorariosChange
        end
        object DBMemo1: TDBMemo
          Left = 176
          Top = 47
          Width = 417
          Height = 41
          DataField = 'mDescripcionBitacora'
          DataSource = dsDatos
          ReadOnly = True
          TabOrder = 3
        end
      end
      object gridActividades: TDBGrid
        Left = 0
        Top = 121
        Width = 671
        Height = 133
        Align = alClient
        DataSource = dsDatos
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = gridActividadesDrawColumnCell
        OnDblClick = gridActividadesDblClick
        Columns = <
          item
            Expanded = False
            FieldName = 'sNumeroActividad'
            Title.Caption = 'No. Actividad'
            Width = 70
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'sHoraInicio'
            Title.Caption = 'Inicio'
            Width = 35
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'sHoraFinal'
            Title.Caption = 'Final'
            Width = 35
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'sIdPartidaAnexo'
            Title.Caption = 'Partida'
            Width = 80
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'dJornadaRed'
            Title.Caption = 'Jornada/Cant'
            Width = 150
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'dJornadaAjustada'
            Title.Caption = 'Ajustado'
            Width = 150
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'sTituloPartidaAnexo'
            Title.Caption = 'T'#237'tulo Partida'
            Visible = True
          end>
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 254
    Width = 1003
    Height = 41
    Align = alBottom
    Padding.Left = 6
    Padding.Top = 4
    Padding.Right = 6
    Padding.Bottom = 4
    TabOrder = 1
    object btnGrabar: TButton
      Left = 846
      Top = 5
      Width = 75
      Height = 31
      Align = alRight
      Caption = '&Grabar'
      Enabled = False
      TabOrder = 0
      OnClick = btnGrabarClick
    end
    object btnCerrar: TButton
      Left = 921
      Top = 5
      Width = 75
      Height = 31
      Align = alRight
      Caption = '&Cerrar'
      TabOrder = 1
      OnClick = btnCerrarClick
    end
  end
  object pnlEditar: TPanel
    Left = 680
    Top = 8
    Width = 257
    Height = 126
    TabOrder = 2
    Visible = False
    object Label4: TLabel
      Left = 16
      Top = 51
      Width = 81
      Height = 13
      Caption = 'Nueva Cantidad:'
    end
    object dJornadaAjustadaEdit: TJvCalcEdit
      Left = 105
      Top = 48
      Width = 144
      Height = 21
      DecimalPlaces = 6
      DisplayFormat = ',0.######'
      TabOrder = 0
      DecimalPlacesAlwaysShown = False
    end
    object btnAceptar: TButton
      Left = 95
      Top = 96
      Width = 75
      Height = 25
      Caption = '&Aceptar'
      Default = True
      ModalResult = 1
      TabOrder = 1
    end
    object btnCancelar: TButton
      Left = 176
      Top = 96
      Width = 75
      Height = 25
      Cancel = True
      Caption = '&Cancelar'
      ModalResult = 2
      TabOrder = 2
    end
    object Panel6: TPanel
      Left = 1
      Top = 1
      Width = 255
      Height = 41
      Align = alTop
      BevelOuter = bvNone
      Enabled = False
      TabOrder = 3
      object Label3: TLabel
        Left = 16
        Top = 19
        Width = 48
        Height = 13
        Caption = 'Jornadas:'
      end
      object dJornadaEdit: TJvCalcEdit
        Left = 72
        Top = 16
        Width = 144
        Height = 21
        DecimalPlaces = 6
        DisplayFormat = ',0.######'
        TabOrder = 0
        DecimalPlacesAlwaysShown = False
      end
    end
  end
  object roqOrdenes: TZReadOnlyQuery
    Connection = connection.zConnection
    AfterScroll = roqOrdenesAfterScroll
    SQL.Strings = (
      'SELECT'
      '  ot.sContrato,'
      '  ot.sNumeroOrden,'
      '  ot.sIdFolio,'
      '  CAST(ot.mDescripcion AS CHAR(1000)) AS mDescripcion,'
      '  ba.dIdFecha'
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
      '  ot.sContrato = :sContrato AND'
      '  ba.dIdFecha = :Fecha'
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
        Name = 'sContrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    Left = 64
    Top = 56
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'sContrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
  end
  object roqReporte: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      '/* MANO DE OBRA */'
      'SELECT'
      '  0 AS iIdOrdenTipoAnexo,'
      '  ba.sContrato,'
      '  ba.sNumeroOrden,'
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
      '  bp.dCantidad AS dCantidad,'
      '  bp.dCantHH AS dJornada,'
      '  ROUND(bp.dCantHH, 6) AS dJornadaRed,'
      ''
      '  IFNULL((SELECT'
      '            CAST(bj.dAjuste AS CHAR)'
      '          FROM'
      '            bitacoradeajustes bj'
      '          WHERE'
      '            bj.sContrato = ba.sContrato AND'
      '            bj.sNumeroOrden = ba.sNumeroOrden AND'
      '            bj.sNumeroActividad = ba.sNumeroActividad AND'
      '            bj.dFecha = ba.dIdFecha AND'
      '            bj.sIdPartidaAnexo = bp.sIdPersonal), 0) AS dAjuste,'
      '  (ROUND(bp.dCantHH, 6) + IFNULL((SELECT'
      '                                    CAST(bj.dAjuste AS CHAR)'
      '                                  FROM'
      '                                    bitacoradeajustes bj'
      '                                  WHERE'
      
        '                                    bj.sContrato = ba.sContrato ' +
        'AND'
      
        '                                    bj.sNumeroOrden = ba.sNumero' +
        'Orden AND'
      
        '                                    bj.sNumeroActividad = ba.sNu' +
        'meroActividad AND'
      '                                    bj.dFecha = ba.dIdFecha AND'
      
        '                                    bj.sIdPartidaAnexo = bp.sIdP' +
        'ersonal), 0)) AS dJornadaAjustada,'
      ''
      '  ba.mDescripcion AS mDescripcionBitacoraActividades,'
      '  mo.dVentaMN,'
      '  mo.dVentaDLL,'
      '  mo.sIdPersonal AS sIdPartidaAnexo,'
      '  CAST(mo.sDescripcion AS CHAR) AS sTituloPartidaAnexo,'
      '  Null AS sIdCategoria'
      ''
      'FROM'
      '  bitacoradeactividades ba'
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
      '  ba.sContrato = :sContrato AND'
      '  ba.sNumeroOrden = :sNumeroOrden AND'
      '  ba.dCantidad > 0 AND'
      '  ba.dIdFecha = :Fecha'
      ''
      'GROUP BY'
      '  ba.sContrato,'
      '  ba.sNumeroOrden,'
      '  sTipoAnexo,'
      '  dIdFecha,'
      '  sNumeroActividad,'
      '  sIdPartidaAnexo,'
      '  sIdCategoria'
      ''
      'UNION'
      ''
      '/* HERRAMIENTA Y EQUIPO */'
      'SELECT'
      '  1 AS iIdOrdenTipoAnexo,'
      '  ba.sContrato,'
      '  ba.sNumeroOrden,'
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
      '  be.dCantidad AS dCantidad,'
      '  be.dCantHH AS dJornada,'
      '  ROUND(be.dCantHH, 6) AS dJornadaRed,'
      ''
      '  IFNULL((SELECT'
      '            CAST(bj.dAjuste AS CHAR)'
      '          FROM'
      '            bitacoradeajustes bj'
      '          WHERE'
      '            bj.sContrato = ba.sContrato AND'
      '            bj.sNumeroOrden = ba.sNumeroOrden AND'
      '            bj.sNumeroActividad = ba.sNumeroActividad AND'
      '            bj.dFecha = ba.dIdFecha AND'
      '            bj.sIdPartidaAnexo = be.sIdEquipo), 0) AS dAjuste,'
      '  (ROUND(be.dCantHH, 6) + IFNULL((SELECT'
      '                                    CAST(bj.dAjuste AS CHAR)'
      '                                  FROM'
      '                                    bitacoradeajustes bj'
      '                                  WHERE'
      
        '                                    bj.sContrato = ba.sContrato ' +
        'AND'
      
        '                                    bj.sNumeroOrden = ba.sNumero' +
        'Orden AND'
      
        '                                    bj.sNumeroActividad = ba.sNu' +
        'meroActividad AND'
      '                                    bj.dFecha = ba.dIdFecha AND'
      
        '                                    bj.sIdPartidaAnexo = be.sIdE' +
        'quipo), 0)) AS dJornadaAjustada,'
      ''
      '  ba.mDescripcion AS mDescripcionBitacoraActividades,'
      '  eq.dVentaMN,'
      '  eq.dVentaDLL,'
      '  eq.sIdEquipo AS sIdPartidaAnexo,'
      '  CAST(eq.sDescripcion AS CHAR) AS sTituloPartidaAnexo,'
      '  Null AS sIdCategoria'
      ''
      'FROM'
      '  bitacoradeactividades ba'
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
      '  ba.sContrato = :sContrato AND'
      '  ba.sNumeroOrden = :sNumeroOrden AND'
      '  ba.dCantidad > 0 AND'
      '  ba.dIdFecha = :Fecha'
      ''
      'GROUP BY'
      '  ba.sContrato,'
      '  ba.sNumeroOrden,'
      '  sTipoAnexo,'
      '  dIdFecha,'
      '  sNumeroActividad,'
      '  sIdPartidaAnexo,'
      '  sIdCategoria'
      ''
      'UNION'
      ''
      '/* PERNOCTAS */'
      'SELECT'
      '  2 AS iIdOrdenTipoAnexo,'
      '  bp.sContrato,'
      '  bp.sNumeroOrden,'
      '  cat.sDescripcion AS mDescripcionBitacora,'
      '  "PN" AS sTipoAnexo,'
      '  "" AS Frente,'
      '  "" AS sIdClasificacion,'
      '  bp.dIdFecha,'
      '  CAST(bp.sIdCategoria AS UNSIGNED) AS iItemOrden,'
      '  "" AS sHoraInicio,'
      '  "" AS sHoraFinal,'
      '  "" AS sWbs,'
      '  "" AS sNumeroActividad,'
      '  SUM(bp.dCantidad) AS dCantidad,'
      '  SUM(bp.dCantidad) AS dJornada,'
      '  SUM(ROUND(bp.dCantidad, 6)) AS dJornadaRed,'
      '  IFNULL((SELECT'
      '            SUM(CAST(bj.dAjuste AS CHAR))'
      '          FROM'
      '            bitacoradeajustes bj'
      '          WHERE'
      '            bj.sContrato = bp.sContrato AND'
      '            bj.sNumeroOrden = bp.sNumeroOrden AND'
      '            bj.sIdCategoria = bp.sIdCategoria AND'
      '            bj.dFecha = bp.dIdFecha AND'
      '            bj.sIdPartidaAnexo = bp.sIdCuenta), 0) AS dAjuste,'
      '  (IFNULL(SUM(ROUND(bp.dCantidad, 6)), 0) + '
      '   IFNULL((SELECT'
      '             SUM(CAST(bj.dAjuste AS CHAR))'
      '           FROM'
      '             bitacoradeajustes bj'
      '           WHERE'
      '             bj.sContrato = bp.sContrato AND'
      '             bj.sNumeroOrden = bp.sNumeroOrden AND'
      '             bj.sIdCategoria = bp.sIdCategoria AND'
      '             bj.dFecha = bp.dIdFecha AND'
      
        '             bj.sIdPartidaAnexo = bp.sIdCuenta), 0)) AS dJornada' +
        'Ajustada,'
      '  "" AS mDescripcionBitacoraActividades,'
      '  cta.dVentaMN,'
      '  cta.dVentaDLL,'
      '  bp.sIdCuenta AS sIdPartidaAnexo,'
      '  CAST(cat.sDescripcion AS CHAR) AS sTituloPartidaAnexo,'
      '  bp.sIdCategoria'
      '  '
      'FROM'
      '  bitacoradepernocta bp'
      ''
      'INNER JOIN'
      '  cuentas cta'
      '    ON (cta.sIdCuenta = bp.sIdCuenta)'
      ''
      'INNER JOIN'
      '  categorias cat'
      '    ON (cat.sIdCategoria = bp.sIdCategoria)'
      ''
      'WHERE'
      '  bp.sContrato = :sContrato AND'
      '  bp.sNumeroOrden = :sNumeroOrden AND'
      '  bp.dCantidad > 0 AND'
      '  bp.dIdFecha = :Fecha'
      ''
      'GROUP BY'
      '  bp.sContrato,'
      '  bp.sNumeroOrden,'
      '  sTipoAnexo,'
      '  dIdFecha,'
      '  sNumeroActividad,'
      '  sIdPartidaAnexo,'
      '  sIdCategoria'
      ''
      'ORDER BY'
      '  iIdOrdenTipoAnexo,'
      '  dIdFecha DESC,'
      '  iItemOrden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'sContrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sNumeroOrden'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
    Left = 304
    Top = 64
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'sContrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sNumeroOrden'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end>
  end
  object dsOrdenes: TDataSource
    DataSet = roqOrdenes
    Left = 88
    Top = 104
  end
  object dsDatos: TDataSource
    DataSet = memDatos
    Left = 416
    Top = 152
  end
  object memDatos: TJvMemoryData
    FieldDefs = <>
    Left = 352
    Top = 152
  end
  object roqExec: TZReadOnlyQuery
    Connection = connection.zConnection
    Params = <>
    Left = 328
    Top = 200
  end
end
