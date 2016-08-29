object FrmGeneradores: TFrmGeneradores
  Left = 0
  Top = 0
  BorderStyle = bsToolWindow
  Caption = 'Impresion de Generadores'
  ClientHeight = 291
  ClientWidth = 506
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Position = poMainFormCenter
  Scaled = False
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GBx1: TcxGroupBox
    Left = 0
    Top = 0
    Align = alTop
    Caption = 'Parametros de Filtrado'
    TabOrder = 0
    Height = 105
    Width = 506
    object dxLayoutControl1: TdxLayoutControl
      Left = 2
      Top = 18
      Width = 502
      Height = 85
      Align = alClient
      TabOrder = 0
      object DtEdtFechaInicio: TcxDateEdit
        Left = 72
        Top = 10
        Properties.OnCloseUp = DtEdtFechaFinPropertiesCloseUp
        Properties.OnEditValueChanged = DtEdtFechaInicioPropertiesEditValueChanged
        Style.BorderColor = clWindowFrame
        Style.BorderStyle = ebs3D
        Style.HotTrack = False
        Style.ButtonStyle = bts3D
        Style.PopupBorderStyle = epbsFrame3D
        TabOrder = 0
        Width = 97
      end
      object DtEdtFechaFin: TcxDateEdit
        Left = 72
        Top = 37
        Properties.OnCloseUp = DtEdtFechaFinPropertiesCloseUp
        Properties.OnEditValueChanged = DtEdtFechaFinPropertiesEditValueChanged
        Style.BorderColor = clWindowFrame
        Style.BorderStyle = ebs3D
        Style.HotTrack = False
        Style.ButtonStyle = bts3D
        Style.PopupBorderStyle = epbsFrame3D
        TabOrder = 2
        Width = 97
      end
      object LCmbOrdenes: TcxLookupComboBox
        Left = 265
        Top = 10
        Properties.DropDownListStyle = lsFixedList
        Properties.KeyFieldNames = 'sContrato'
        Properties.ListColumns = <
          item
            FieldName = 'sContrato'
          end>
        Properties.ListOptions.ShowHeader = False
        Properties.ListOptions.SyncMode = True
        Properties.ListSource = dsOrdenes
        Style.BorderColor = clWindowFrame
        Style.BorderStyle = ebs3D
        Style.HotTrack = False
        Style.ButtonStyle = bts3D
        Style.PopupBorderStyle = epbsFrame3D
        TabOrder = 1
        Width = 241
      end
      object LCmbFolios: TcxLookupComboBox
        Left = 265
        Top = 37
        Properties.DropDownListStyle = lsFixedList
        Properties.KeyFieldNames = 'sNumeroOrden'
        Properties.ListColumns = <
          item
            FieldName = 'sNumeroOrden'
          end>
        Properties.ListOptions.ShowHeader = False
        Properties.ListOptions.SyncMode = True
        Properties.ListSource = dsFolios
        Style.BorderColor = clWindowFrame
        Style.BorderStyle = ebs3D
        Style.HotTrack = False
        Style.ButtonStyle = bts3D
        Style.PopupBorderStyle = epbsFrame3D
        TabOrder = 3
        Width = 145
      end
      object dxLayoutControl1Group_Root: TdxLayoutGroup
        AlignHorz = ahClient
        AlignVert = avTop
        ButtonOptions.Buttons = <>
        Hidden = True
        ShowBorder = False
        Index = -1
      end
      object dxLayoutControl1Item1: TdxLayoutItem
        Parent = dxLayoutControl1Group1
        AlignHorz = ahLeft
        CaptionOptions.Text = 'Fecha Inicio'
        Control = DtEdtFechaInicio
        ControlOptions.ShowBorder = False
        Index = 0
      end
      object dxLayoutControl1Item2: TdxLayoutItem
        Parent = dxLayoutControl1Group2
        AlignHorz = ahLeft
        CaptionOptions.Text = 'Fecha Final'
        Control = DtEdtFechaFin
        ControlOptions.ShowBorder = False
        Index = 0
      end
      object dxLayoutControl1Item3: TdxLayoutItem
        Parent = dxLayoutControl1Group1
        AlignHorz = ahClient
        CaptionOptions.Text = 'Orden de Trabajo'
        Control = LCmbOrdenes
        ControlOptions.ShowBorder = False
        Index = 1
      end
      object dxLayoutControl1Group1: TdxLayoutAutoCreatedGroup
        Parent = dxLayoutControl1Group_Root
        LayoutDirection = ldHorizontal
        Index = 0
        AutoCreated = True
      end
      object dxLayoutControl1Item4: TdxLayoutItem
        Parent = dxLayoutControl1Group2
        AlignHorz = ahClient
        CaptionOptions.Text = 'Folio'
        Control = LCmbFolios
        ControlOptions.ShowBorder = False
        Index = 1
      end
      object dxLayoutControl1Group2: TdxLayoutAutoCreatedGroup
        Parent = dxLayoutControl1Group_Root
        LayoutDirection = ldHorizontal
        Index = 1
        AutoCreated = True
      end
    end
  end
  object GBx2: TcxGroupBox
    Left = 0
    Top = 105
    Align = alClient
    Caption = 'Opciones de Generador'
    TabOrder = 1
    Height = 186
    Width = 506
    object dxLayoutControl2: TdxLayoutControl
      Left = 2
      Top = 18
      Width = 502
      Height = 166
      Align = alClient
      TabOrder = 0
      object RdGpGeneradores: TcxRadioGroup
        Left = 10
        Top = 10
        Properties.Items = <
          item
            Caption = 'Personal'
          end
          item
            Caption = 'Horas Extra'
          end
          item
            Caption = 'Equipos'
          end>
        ItemIndex = 0
        Style.BorderColor = clWindowFrame
        Style.BorderStyle = ebs3D
        TabOrder = 0
        Height = 105
        Width = 471
      end
      object btnIMprimir: TcxButton
        Left = 208
        Top = 121
        Width = 75
        Height = 25
        Caption = '&Imprimir'
        TabOrder = 1
        OnClick = btnIMprimirClick
      end
      object dxLayoutControl2Group_Root: TdxLayoutGroup
        AlignHorz = ahLeft
        AlignVert = avTop
        ButtonOptions.Buttons = <>
        Hidden = True
        ShowBorder = False
        Index = -1
      end
      object dxLayoutControl2Item1: TdxLayoutItem
        Parent = dxLayoutControl2Group_Root
        CaptionOptions.Text = 'cxRadioGroup1'
        CaptionOptions.Visible = False
        Control = RdGpGeneradores
        ControlOptions.ShowBorder = False
        Index = 0
      end
      object dxLayoutControl2Item2: TdxLayoutItem
        Parent = dxLayoutControl2Group_Root
        AlignHorz = ahCenter
        CaptionOptions.Text = 'cxButton1'
        CaptionOptions.Visible = False
        Control = btnIMprimir
        ControlOptions.ShowBorder = False
        Index = 1
      end
    end
  end
  object QrOrdenes: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select c.sContrato, c.sTipoObra from bitacoradeactividades b'
      
        'inner join reportediario r on (r.sOrden = b.sContrato and r.dIdF' +
        'echa = b.dIdFecha )'
      'inner join contratos c on (b.sContrato = c.sContrato)'
      
        'where b.dIdFecha between :fechaI and :fechaF group by b.sContrat' +
        'o')
    Params = <
      item
        DataType = ftUnknown
        Name = 'fechaI'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'fechaF'
        ParamType = ptUnknown
      end>
    Left = 152
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'fechaI'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'fechaF'
        ParamType = ptUnknown
      end>
  end
  object QrFolios: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'Select sContrato,sNumeroOrden from ordenesdetrabajo Order By sCo' +
        'ntrato,sNumeroOrden')
    Params = <>
    MasterFields = 'sContrato'
    MasterSource = dsOrdenes
    LinkedFields = 'sContrato'
    Left = 200
    Top = 72
  end
  object dsOrdenes: TDataSource
    DataSet = QrOrdenes
    Left = 200
  end
  object dsFolios: TDataSource
    DataSet = QrFolios
    Left = 240
    Top = 72
  end
  object FrReporte: TfrxReport
    Version = '4.10.3'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Por defecto'
    PrintOptions.PrintOnSheet = 0
    ReportOptions.CreateDate = 38372.938800231500000000
    ReportOptions.LastChange = 41928.829860925900000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'var'
      '   xCaracter : byte ;'
      '   dJornadasNormales,'
      '   dJornadasExtras : Double ;'
      
        '   Separador : string;                                          ' +
        '   '
      ''
      'procedure GroupFooter1OnAfterPrint(Sender: TfrxComponent);'
      'begin'
      '   dJornadasNormales := 0 ;                 '
      '   dJornadasExtras := 0 ;                '
      'end;'
      ''
      'procedure MasterData2OnBeforePrint(Sender: TfrxComponent);'
      'begin'
      ' {  if <dsGeneradorDia."sIdTipoPersonal"> = '#39'EXT'#39' then'
      
        '        dJornadasExtras := dJornadasExtras + <dsGeneradorDia."dT' +
        'otal">'
      '   Else'
      
        '        dJornadasNormales := dJornadasNormales + <dsGeneradorDia' +
        '."dTotal"> ;  }    '
      'end;'
      ''
      
        'procedure ReporteDiarioBarcoOnBeforePrint(Sender: TfrxComponent)' +
        ';'
      'begin'
      '     { if <dsGeneradorDia."sAnexo"> <> '#39#39' then'
      '        Separador := '#39'.'#39';  }  '
      
        '      //  showmessage(<dsConfiguracion."iFirmasGeneradores">);  ' +
        '                                          '
      '      if <dsConfiguracion."iFirmasGeneradores"> = 2 then'
      '      begin'
      '           memo36.Visible := False;'
      '           memo37.Visible := False;'
      
        '           memo39.Visible := False;                             ' +
        '                              '
      '      end;                        '
      'end;'
      ''
      'Begin'
      '  '
      
        '    memo100.text := '#39'1'#39' ;                                       ' +
        '                                                     '
      
        '    memo101.text := '#39'2'#39' ;                                       ' +
        '                                                     '
      
        '    memo102.text := '#39'3'#39' ;                                       ' +
        '                                                     '
      
        '    memo103.text := '#39'4'#39' ;                                       ' +
        '                                                     '
      
        '    memo104.text := '#39'5'#39' ;                                       ' +
        '                                                     '
      
        '    memo105.text := '#39'6'#39' ;                                       ' +
        '                                                     '
      
        '    memo106.text := '#39'7'#39' ;                                       ' +
        '                                                     '
      
        '    memo107.text := '#39'8'#39' ;                                       ' +
        '                                                     '
      
        '    memo108.text := '#39'9'#39' ;                                       ' +
        '                                                     '
      
        '    memo109.text := '#39'10'#39' ;                                      ' +
        '                                                      '
      
        '    memo110.text := '#39'11'#39' ;                                      ' +
        '                                                      '
      
        '    memo111.text := '#39'12'#39' ;                                      ' +
        '                                                      '
      
        '    memo112.text := '#39'13'#39' ;                                      ' +
        '                                                      '
      
        '    memo113.text := '#39'14'#39' ;                                      ' +
        '                                                      '
      
        '    memo114.text := '#39'15'#39' ;                                      ' +
        '                                                      '
      
        '    memo115.text := '#39'16'#39' ;                                      ' +
        '                                                      '
      
        '    memo116.text := '#39'17'#39' ;                                      ' +
        '                                                      '
      
        '    memo117.text := '#39'18'#39' ;                                      ' +
        '                                                      '
      
        '    memo118.text := '#39'19'#39' ;                                      ' +
        '                                                      '
      
        '    memo119.text := '#39'20'#39' ;                                      ' +
        '                                                      '
      
        '    memo120.text := '#39'21'#39' ;                                      ' +
        '                                                      '
      
        '    memo121.text := '#39'22'#39' ;                                      ' +
        '                                                      '
      
        '    memo122.text := '#39'23'#39' ;                                      ' +
        '                                                      '
      
        '    memo123.text := '#39'24'#39' ;                                      ' +
        '                                                      '
      
        '    memo124.text := '#39'25'#39' ;                                      ' +
        '                                                      '
      
        '    memo125.text := '#39'26'#39' ;                                      ' +
        '                                                      '
      
        '    memo126.text := '#39'27'#39' ;                                      ' +
        '                                                      '
      '    memo127.text := '#39'28'#39' ;'
      
        '    memo128.text := '#39'29'#39' ;                                      ' +
        '                                                      '
      
        '    memo129.text := '#39'30'#39' ;                                      ' +
        '                                                      '
      
        '    memo130.text := '#39'31'#39' ;                                      ' +
        '                                                      '
      'End.')
    OnGetValue = FrReporteGetValue
    Left = 256
    Top = 168
    Datasets = <
      item
        DataSet = frmReportePeriodo.dsConfiguracion
        DataSetName = 'dsConfiguracion'
      end
      item
        DataSet = frmDiarioTurno2.dsGeneradorGeneral
        DataSetName = 'dsGeneradorGeneral'
      end
      item
        DataSet = frmDiarioTurno2.TD_ConfigOTBarco
        DataSetName = 'TD_ConfigOTBarco'
      end
      item
        DataSet = frmDiarioTurno2.Td_contrato
        DataSetName = 'Td_contrato'
      end>
    Variables = <>
    Style = <
      item
        Name = 'Title'
        Color = clNavy
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -16
        Font.Name = 'Arial'
        Font.Style = [fsBold]
      end
      item
        Name = 'Header'
        Color = clNone
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = [fsBold]
      end
      item
        Name = 'Group header'
        Color = 15790320
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = [fsBold]
      end
      item
        Name = 'Data'
        Color = clNone
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = []
      end
      item
        Name = 'Group footer'
        Color = clNone
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = [fsBold]
      end
      item
        Name = 'Header line'
        Color = clNone
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = []
        Frame.Typ = [ftBottom]
        Frame.Width = 2.000000000000000000
      end>
    object Data: TfrxDataPage
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      Height = 223.000000000000000000
      Left = 613.000000000000000000
      Top = 186.000000000000000000
      Width = 336.000000000000000000
    end
    object ReporteDiarioBarco: TfrxReportPage
      Orientation = poLandscape
      PaperWidth = 279.400000000000000000
      PaperHeight = 215.900000000000000000
      PaperSize = 256
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      OnBeforePrint = 'ReporteDiarioBarcoOnBeforePrint'
      object GroupHeader1: TfrxGroupHeader
        Height = 238.110375350000000000
        Top = 18.897650000000000000
        Width = 980.410082000000000000
        AllowSplit = True
        Condition = 'Td_contrato."sContrato"'
        ReprintOnNewPage = True
        OutlineText = 'Td_contrato."sContrato"'
        object Picture2: TfrxPictureView
          Top = 7.559060000000000000
          Width = 151.181102360000000000
          Height = 56.692913390000000000
          ShowHint = False
          Center = True
          DataField = 'bImagen'
          DataSet = frmDiarioTurno2.Td_contrato
          DataSetName = 'Td_contrato'
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Picture4: TfrxPictureView
          Left = 827.717070000000000000
          Top = 7.559060000000000000
          Width = 151.181102360000000000
          Height = 56.692913390000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmReportePeriodo.dsConfiguracion
          DataSetName = 'dsConfiguracion'
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Memo100: TfrxMemoView
          ShiftMode = smDontShift
          Left = 37.795275590000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo26: TfrxMemoView
          ShiftMode = smDontShift
          Left = 880.629952990000000000
          Top = 201.433112360000000000
          Width = 49.133858270000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'TOTAL')
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo101: TfrxMemoView
          ShiftMode = smDontShift
          Left = 65.007874020000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo103: TfrxMemoView
          ShiftMode = smDontShift
          Left = 119.433070870000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo104: TfrxMemoView
          ShiftMode = smDontShift
          Left = 146.645669290000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo105: TfrxMemoView
          ShiftMode = smDontShift
          Left = 173.858267720000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo106: TfrxMemoView
          ShiftMode = smDontShift
          Left = 201.070866140000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo107: TfrxMemoView
          ShiftMode = smDontShift
          Left = 228.283464570000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo108: TfrxMemoView
          ShiftMode = smDontShift
          Left = 255.496062990000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo109: TfrxMemoView
          ShiftMode = smDontShift
          Left = 282.708661420000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo110: TfrxMemoView
          ShiftMode = smDontShift
          Left = 309.921259840000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo111: TfrxMemoView
          ShiftMode = smDontShift
          Left = 337.133858270000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo112: TfrxMemoView
          ShiftMode = smDontShift
          Left = 364.346456690000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo113: TfrxMemoView
          ShiftMode = smDontShift
          Left = 391.559055120000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo114: TfrxMemoView
          ShiftMode = smDontShift
          Left = 418.771653540000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo115: TfrxMemoView
          ShiftMode = smDontShift
          Left = 445.984251970000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo116: TfrxMemoView
          ShiftMode = smDontShift
          Left = 473.196850390000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo117: TfrxMemoView
          ShiftMode = smDontShift
          Left = 500.409448820000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo118: TfrxMemoView
          ShiftMode = smDontShift
          Left = 527.622047240000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo119: TfrxMemoView
          ShiftMode = smDontShift
          Left = 554.834645670000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo120: TfrxMemoView
          ShiftMode = smDontShift
          Left = 582.047244090000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo121: TfrxMemoView
          ShiftMode = smDontShift
          Left = 609.259842520000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo122: TfrxMemoView
          ShiftMode = smDontShift
          Left = 636.472440940000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo123: TfrxMemoView
          ShiftMode = smDontShift
          Left = 663.685039370000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo124: TfrxMemoView
          ShiftMode = smDontShift
          Left = 690.897637800000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo125: TfrxMemoView
          ShiftMode = smDontShift
          Left = 718.110236220000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo126: TfrxMemoView
          ShiftMode = smDontShift
          Left = 745.322834650000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo127: TfrxMemoView
          ShiftMode = smDontShift
          Left = 772.535433070000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo128: TfrxMemoView
          ShiftMode = smDontShift
          Left = 799.748031500000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo129: TfrxMemoView
          ShiftMode = smDontShift
          Left = 826.960629920000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo130: TfrxMemoView
          ShiftMode = smDontShift
          Left = 854.173228350000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo102: TfrxMemoView
          ShiftMode = smDontShift
          Left = 92.220472440000000000
          Top = 201.433063540000000000
          Width = 27.212598430000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo20: TfrxMemoView
          Left = 151.181200000000000000
          Top = 7.559060000000000000
          Width = 676.535870000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Arial Black'
          Font.Style = [fsBold, fsItalic]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_contrato."mCliente"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo23: TfrxMemoView
          Left = 151.181200000000000000
          Top = 26.456710000000000000
          Width = 676.535870000000000000
          Height = 49.133890000000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."mClienteBarco"]')
          ParentFont = False
        end
        object Rich1: TfrxRichView
          ShiftMode = smDontShift
          Left = 1.889763780000000000
          Top = 83.149660000000000000
          Width = 721.890230000000000000
          Height = 26.456702680000000000
          ShowHint = False
          StretchMode = smActualHeight
          GapX = 2.000000000000000000
          GapY = 1.000000000000000000
          RichEdit = {
            7B5C727466315C616E73695C616E7369637067313235325C64656666305C6E6F
            7569636F6D7061745C6465666C616E67323035387B5C666F6E7474626C7B5C66
            305C666E696C5C66636861727365743020417269616C3B7D7D0D0A7B5C2A5C67
            656E657261746F722052696368656432302031302E302E31303538367D5C7669
            65776B696E64345C756331200D0A5C706172645C716A5C625C66733134204F42
            52413A205C6230205B6473436F6E66696775726163696F6E2E226D4465736372
            697063696F6E426172636F225D5C667331325C7061720D0A7D0D0A00}
        end
        object Memo98: TfrxMemoView
          ShiftMode = smDontShift
          Left = 729.449290000000000000
          Top = 83.149660000000000000
          Width = 132.283464570000000000
          Height = 26.456692910000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight]
          Memo.UTF8W = (
            'ESTIMACION No. ')
          ParentFont = False
        end
        object Memo99: TfrxMemoView
          ShiftMode = smDontShift
          Left = 861.732840000000000000
          Top = 83.149660000000000000
          Width = 105.826766770000000000
          Height = 26.456692910000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[ESTIMACION] [MONEDA]')
          ParentFont = False
        end
        object Memo131: TfrxMemoView
          ShiftMode = smDontShift
          Left = 729.449290000000000000
          Top = 109.606299210000000000
          Width = 132.283464570000000000
          Height = 15.118110240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight]
          Memo.UTF8W = (
            'CONTRATO')
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo132: TfrxMemoView
          ShiftMode = smDontShift
          Left = 861.732840000000000000
          Top = 109.606299210000000000
          Width = 105.826766770000000000
          Height = 15.118110240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haRight
          Memo.UTF8W = (
            '[dsConfiguracion."sContratoBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo133: TfrxMemoView
          ShiftMode = smDontShift
          Left = 729.449290000000000000
          Top = 124.724409450000000000
          Width = 132.283464570000000000
          Height = 15.118110240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight]
          Memo.UTF8W = (
            'MONTO DE CONTRATO M.N')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo134: TfrxMemoView
          ShiftMode = smDontShift
          Left = 729.449290000000000000
          Top = 140.401574800000000000
          Width = 132.283464570000000000
          Height = 30.236230240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight]
          Memo.UTF8W = (
            'MONTO DE CONTRATO USD.')
          ParentFont = False
        end
        object Memo19: TfrxMemoView
          ShiftMode = smDontShift
          Left = 861.732840000000000000
          Top = 124.724409450000000000
          Width = 105.826766770000000000
          Height = 15.118110240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haRight
          Memo.UTF8W = (
            ' [TD_ConfigOTBarco."dMontoMn"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo21: TfrxMemoView
          ShiftMode = smDontShift
          Left = 861.732840000000000000
          Top = 140.401574800000000000
          Width = 105.826766770000000000
          Height = 30.236230240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haRight
          Memo.UTF8W = (
            ' [TD_ConfigOTBarco."dMontoDll"]')
          ParentFont = False
        end
        object Memo22: TfrxMemoView
          ShiftMode = smDontShift
          Left = 929.764380000000000000
          Top = 201.433210000000000000
          Width = 49.133858270000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'UNIDAD')
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo24: TfrxMemoView
          ShiftMode = smDontShift
          Top = 201.433210000000000000
          Width = 37.795275590000000000
          Height = 22.677165350000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PART.')
          ParentFont = False
          WordWrap = False
          VAlign = vaCenter
        end
        object Memo135: TfrxMemoView
          ShiftMode = smDontShift
          Top = 180.535464800000000000
          Width = 978.898245590000000000
          Height = 18.897632910000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            'GENERADOR DE ESTIMACION POR PARTIDAS DEL ANEXO C')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo3: TfrxMemoView
          Left = 1.889763780000000000
          Top = 110.047310000000000000
          Width = 75.590600000000000000
          Height = 15.118120000000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Memo.UTF8W = (
            'CONTRATISTA:')
          ParentFont = False
        end
        object Memo12: TfrxMemoView
          Left = 1.889763780000000000
          Top = 124.724490000000000000
          Width = 113.385900000000000000
          Height = 15.118120000000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Memo.UTF8W = (
            'PLAZO DE EJECUCI'#211'N:')
          ParentFont = False
        end
        object Memo18: TfrxMemoView
          Left = 1.889763780000000000
          Top = 147.401670000000000000
          Width = 75.590600000000000000
          Height = 11.338590000000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Memo.UTF8W = (
            'PERIODO:')
          ParentFont = False
        end
        object Memo45: TfrxMemoView
          Left = 77.149660000000000000
          Top = 109.984251970000000000
          Width = 370.393940000000000000
          Height = 15.118110240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Memo.UTF8W = (
            '[dsConfiguracion."sNombre"]')
          ParentFont = False
        end
        object Memo50: TfrxMemoView
          Left = 115.165430000000000000
          Top = 124.724490000000000000
          Width = 336.378170000000000000
          Height = 15.118110240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Memo.UTF8W = (
            '[TD_ConfigOTBarco."dDuracion"] DIAS NATURALES')
          ParentFont = False
        end
        object Memo52: TfrxMemoView
          Left = 78.370130000000000000
          Top = 146.401670000000000000
          Width = 374.173470000000000000
          Height = 11.338580240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Memo.UTF8W = (
            '[PERGENOPT]')
          ParentFont = False
        end
        object Memo53: TfrxMemoView
          Left = 457.323130000000000000
          Top = 109.984251968504000000
          Width = 86.929190000000000000
          Height = 15.118120000000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Memo.UTF8W = (
            'FECHA DE INICIO:')
          ParentFont = False
        end
        object Memo54: TfrxMemoView
          Left = 457.323130000000000000
          Top = 124.724490000000000000
          Width = 128.504020000000000000
          Height = 15.118120000000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Memo.UTF8W = (
            'FECHA DE TERMINACION:')
          ParentFont = False
        end
        object Memo55: TfrxMemoView
          Left = 544.252320000000000000
          Top = 109.984251968504000000
          Width = 132.283550000000000000
          Height = 15.118110240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Memo.UTF8W = (
            '[TD_ConfigOTBarco."dFechaInicio"]')
          ParentFont = False
        end
        object Memo56: TfrxMemoView
          Left = 586.047620000000000000
          Top = 124.724490000000000000
          Width = 98.267780000000000000
          Height = 15.118110240000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Memo.UTF8W = (
            '[TD_ConfigOTBarco."dFechaFinal"]')
          ParentFont = False
        end
        object Shape1: TfrxShapeView
          Top = 81.149660000000000000
          Width = 978.898270000000000000
          Height = 90.330691570000000000
          ShowHint = False
        end
      end
      object GroupHeader2: TfrxGroupHeader
        Height = 22.677180000000000000
        Top = 279.685220000000000000
        Width = 980.410082000000000000
        Condition = 'dsGeneradorGeneral."sAnexo"'
        KeepTogether = True
        Stretched = True
        object Memo136: TfrxMemoView
          Width = 37.795300000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsGeneradorGeneral."sAnexo"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo137: TfrxMemoView
          Left = 37.795300000000000000
          Width = 937.323440000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Memo.UTF8W = (
            '[dsGeneradorGeneral."sTitulo"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object MasterData2: TfrxMasterData
        Height = 26.456702680000000000
        Top = 325.039580000000000000
        Width = 980.410082000000000000
        OnBeforePrint = 'MasterData2OnBeforePrint'
        DataSet = frmDiarioTurno2.dsGeneradorGeneral
        DataSetName = 'dsGeneradorGeneral'
        KeepHeader = True
        RowCount = 0
        Stretched = True
        object Memo17: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Top = 3.779530000000020000
          Width = 980.788016690000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haBlock
          Memo.UTF8W = (
            '[dsGeneradorGeneral."sDescripcion"]')
          ParentFont = False
        end
        object Memo1: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 37.795275590000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia1"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo2: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 65.007874020000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia2'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia2"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo4: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 119.433070870000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia4'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia4"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo5: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 146.645669290000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia5'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia5"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo7: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 173.858267720000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia6'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia6"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo8: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 201.070866140000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia7'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia7"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo9: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 228.283464570000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia8'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia8"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo10: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 255.496062990000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia9'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia9"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo11: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 282.708661420000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia10'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia10"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo13: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 309.921259840000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia11'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia11"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo14: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 337.133858270000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia12'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia12"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo15: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 364.346456690000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia13'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia13"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo16: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 391.559055120000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia14'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia14"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo25: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 418.771653540000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia15'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia15"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo27: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 445.984251970000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia16'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia16"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo28: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 473.196850390000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia17'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia17"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo29: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 500.409448820000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia18'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia18"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo30: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 527.622047240000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia19'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia19"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo31: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 554.834645670000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia20'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia20"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo32: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 582.047244090000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia21'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia21"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo33: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 609.259842520000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia22'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia22"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo34: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 636.472440940000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia23'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia23"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo35: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 663.685039370000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia24'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia24"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo38: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 690.897637800000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia25'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia25"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo40: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 718.110236220000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia26'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia26"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo41: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 745.322834650000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia27'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia27"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo42: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 772.535433070000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia28'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia28"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo43: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 799.748031500000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia29'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia29"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo44: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 826.960629920000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia30'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia30"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo46: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 854.173228350000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia31'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia31"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo47: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 92.220472440000000000
          Top = 15.118120000000000000
          Width = 27.212598430000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataField = 'dia3'
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dia3"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo48: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 880.629921260000000000
          Top = 15.118120000000000000
          Width = 49.133858270000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Color = clWhite
          DataSet = frmDiarioTurno2.dsGeneradorGeneral
          DataSetName = 'dsGeneradorGeneral'
          DisplayFormat.FormatStr = '%5.4n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haRight
          HideZeros = True
          Memo.UTF8W = (
            '[dsGeneradorGeneral."dTotal"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo49: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Left = 929.764380000000000000
          Top = 15.118120000000000000
          Width = 49.133858270000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsGeneradorGeneral."sMedida"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo51: TfrxMemoView
          ShiftMode = smWhenOverlapped
          Top = 15.118120000000000000
          Width = 37.795275590000000000
          Height = 11.338582680000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsGeneradorGeneral."sIdRecurso"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object PageFooter1: TfrxPageFooter
        Height = 105.826840000000000000
        Top = 442.205010000000000000
        Width = 980.410082000000000000
        object Memo91: TfrxMemoView
          Left = 3.779530000000000000
          Top = 15.118120000000000000
          Width = 226.771653540000000000
          Height = 25.456710000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sLeyenda1"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo92: TfrxMemoView
          Left = 4.000000000000000000
          Top = 52.252010000000000000
          Width = 226.771653540000000000
          Height = 26.677180000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftTop]
          HAlign = haCenter
          Memo.UTF8W = (
            '[PUESTO_SUPERINTENDENTE]')
          ParentFont = False
        end
        object Memo93: TfrxMemoView
          Left = 4.000000000000000000
          Top = 39.354359999999900000
          Width = 226.771653540000000000
          Height = 13.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUPERINTENDENTE]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo94: TfrxMemoView
          Left = 746.126470000000000000
          Top = 52.157487640000000000
          Width = 226.771653540000000000
          Height = 26.677180000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftTop]
          HAlign = haCenter
          Memo.UTF8W = (
            '[PUESTO_SUPERVISOR_TIERRA]')
          ParentFont = False
        end
        object Memo95: TfrxMemoView
          Left = 746.126470000000000000
          Top = 39.307093940000000000
          Width = 226.771653540000000000
          Height = 13.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUPERVISOR_TIERRA]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo144: TfrxMemoView
          Left = 746.346940000000000000
          Top = 15.118117560000000000
          Width = 226.771653540000000000
          Height = 24.456710000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sLeyenda2"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo6: TfrxMemoView
          Left = 336.378170000000000000
          Top = 73.252010000000000000
          Width = 321.260050000000000000
          Height = 15.118120000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[Page#] de [TotalPages#]')
          ParentFont = False
        end
        object Memo36: TfrxMemoView
          Left = 377.953000000000000000
          Top = 12.338590000000000000
          Width = 226.771653540000000000
          Height = 25.456710000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sLeyenda3"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo37: TfrxMemoView
          Left = 378.173470000000000000
          Top = 49.472480000000000000
          Width = 226.771653540000000000
          Height = 26.677180000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftTop]
          HAlign = haCenter
          Memo.UTF8W = (
            '[PUESTO_SUPERVISOR]')
          ParentFont = False
        end
        object Memo39: TfrxMemoView
          Left = 378.173470000000000000
          Top = 36.574830000000000000
          Width = 226.771653540000000000
          Height = 13.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUPERVISOR]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupFooter1: TfrxGroupFooter
        Height = 7.559060000000000000
        Top = 374.173470000000000000
        Width = 980.410082000000000000
      end
    end
  end
  object fDbDts1: TfrxDBDataset
    UserName = 'fDbDts1'
    CloseDataSource = False
    BCDToCurrency = False
    Left = 240
    Top = 144
  end
end
