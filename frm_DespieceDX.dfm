object frmDespieceDX: TfrmDespieceDX
  Left = 0
  Top = 0
  BorderIcons = [biMinimize, biMaximize]
  Caption = 'Despiece de partidas'
  ClientHeight = 289
  ClientWidth = 695
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 161
    Width = 695
    Height = 6
    Cursor = crVSplit
    Align = alTop
    Beveled = True
    ExplicitLeft = 1
    ExplicitTop = 128
    ExplicitWidth = 737
  end
  object Label7: TLabel
    Left = 13
    Top = 6
    Width = 15
    Height = 13
    Caption = 'Eje'
  end
  object Label14: TLabel
    Left = 485
    Top = 6
    Width = 36
    Height = 13
    Caption = 'ANCHO'
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 695
    Height = 161
    Align = alTop
    TabOrder = 0
    object Grid_Despiece: TDBGrid
      Left = 1
      Top = 1
      Width = 828
      Height = 160
      TabStop = False
      DataSource = dsDespiece
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      Options = [dgTitles, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
      ParentFont = False
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
      Columns = <
        item
          Expanded = False
          FieldName = 'iIdOrden'
          ReadOnly = True
          Title.Alignment = taCenter
          Title.Caption = 'Orden'
          Width = 63
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'sEje'
          ReadOnly = True
          Title.Alignment = taCenter
          Title.Caption = 'Tipo'
          Width = 86
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'sEje1'
          Title.Alignment = taCenter
          Title.Caption = 'Eje1'
          Width = 51
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'sEje2'
          Title.Alignment = taCenter
          Title.Caption = 'Eje2'
          Width = 57
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dLargo'
          Title.Alignment = taCenter
          Title.Caption = 'Largo'
          Width = 90
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dAncho'
          Title.Caption = 'Alto'
          Width = 91
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dAncho'
          Title.Caption = 'Ancho'
          Width = 98
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dSubtotal'
          Title.Caption = 'Subtotal'
          Width = 96
          Visible = True
        end>
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 167
    Width = 695
    Height = 122
    Align = alClient
    TabOrder = 1
    object lblEje: TLabel
      Left = 13
      Top = 19
      Width = 18
      Height = 13
      Caption = 'EJE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblEntre: TLabel
      Left = 99
      Top = 6
      Width = 34
      Height = 13
      Caption = 'ENTRE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lbleje1: TLabel
      Left = 77
      Top = 18
      Width = 17
      Height = 13
      Caption = 'EJE'
    end
    object lbleje2: TLabel
      Left = 139
      Top = 18
      Width = 17
      Height = 13
      Caption = 'EJE'
    end
    object lblveces: TLabel
      Left = 182
      Top = 6
      Width = 54
      Height = 13
      Alignment = taCenter
      Caption = 'NO. VECES'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label9: TLabel
      Left = 287
      Top = 6
      Width = 38
      Height = 13
      Caption = 'CARAS'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblLargo: TLabel
      Left = 352
      Top = 6
      Width = 38
      Height = 13
      Alignment = taCenter
      Caption = 'LARGO'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label11: TLabel
      Left = 451
      Top = 6
      Width = 29
      Height = 13
      Caption = 'ALTO'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label12: TLabel
      Left = 534
      Top = 6
      Width = 38
      Height = 13
      Caption = 'ANCHO'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label13: TLabel
      Left = 621
      Top = 6
      Width = 58
      Height = 13
      Caption = 'SUBTOTAL'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object tdNumVeces: TRxDBCalcEdit
      Left = 182
      Top = 37
      Width = 78
      Height = 22
      Margins.Left = 4
      Margins.Top = 1
      DataField = 'dNumVeces'
      DataSource = dsDespiece
      DecimalPlaces = 4
      DisplayFormat = ',0.####'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      NumGlyphs = 2
      ParentFont = False
      TabOrder = 0
      OnEnter = tdNumVecesEnter
      OnExit = tdNumVecesExit
      OnKeyPress = tdNumVecesKeyPress
    end
    object tdCaras: TRxDBCalcEdit
      Left = 267
      Top = 37
      Width = 78
      Height = 22
      Margins.Left = 4
      Margins.Top = 1
      DataField = 'dCaras'
      DataSource = dsDespiece
      DecimalPlaces = 4
      DisplayFormat = ',0.####'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      NumGlyphs = 2
      ParentFont = False
      TabOrder = 1
      OnEnter = tdCarasEnter
      OnExit = tdCarasExit
      OnKeyPress = tdCarasKeyPress
    end
    object tdLargo: TRxDBCalcEdit
      Left = 352
      Top = 37
      Width = 78
      Height = 22
      Margins.Left = 4
      Margins.Top = 1
      DataField = 'dLargo'
      DataSource = dsDespiece
      DecimalPlaces = 4
      DisplayFormat = ',0.####'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      NumGlyphs = 2
      ParentFont = False
      TabOrder = 2
      OnEnter = tdLargoEnter
      OnExit = tdLargoExit
      OnKeyPress = tdLargoKeyPress
    end
    object tdAlto: TRxDBCalcEdit
      Left = 437
      Top = 37
      Width = 78
      Height = 22
      Margins.Left = 4
      Margins.Top = 1
      DataField = 'dAlto'
      DataSource = dsDespiece
      DecimalPlaces = 4
      DisplayFormat = ',0.####'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      NumGlyphs = 2
      ParentFont = False
      TabOrder = 3
      OnEnter = tdAltoEnter
      OnExit = tdAltoExit
      OnKeyPress = tdAltoKeyPress
    end
    object tdAncho: TRxDBCalcEdit
      Left = 520
      Top = 37
      Width = 78
      Height = 22
      Margins.Left = 4
      Margins.Top = 1
      DataField = 'dAncho'
      DataSource = dsDespiece
      DecimalPlaces = 4
      DisplayFormat = ',0.####'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      NumGlyphs = 2
      ParentFont = False
      TabOrder = 4
      OnEnter = tdAnchoEnter
      OnExit = tdAnchoExit
      OnKeyPress = tdAnchoKeyPress
    end
    object tdSubtotal: TRxDBCalcEdit
      Left = 605
      Top = 37
      Width = 78
      Height = 22
      Margins.Left = 4
      Margins.Top = 1
      DataField = 'dSubtotal'
      DataSource = dsDespiece
      ReadOnly = True
      DecimalPlaces = 4
      DisplayFormat = ',0.####'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      NumGlyphs = 2
      ParentFont = False
      TabOrder = 5
    end
    object txtEje: TDBEdit
      Left = 7
      Top = 37
      Width = 50
      Height = 22
      DataField = 'sEje'
      DataSource = dsDespiece
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 6
      OnEnter = txtEjeEnter
      OnExit = txtEjeExit
      OnKeyPress = txtEjeKeyPress
    end
    object txtEje1: TDBEdit
      Left = 64
      Top = 37
      Width = 50
      Height = 22
      DataField = 'sEje1'
      DataSource = dsDespiece
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 7
      OnEnter = txtEje1Enter
      OnExit = txtEje1Exit
      OnKeyPress = txtEje1KeyPress
    end
    object txtEje2: TDBEdit
      Left = 124
      Top = 37
      Width = 50
      Height = 22
      DataField = 'sEje2'
      DataSource = dsDespiece
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 8
      OnEnter = txtEje2Enter
      OnExit = txtEje2Exit
      OnKeyPress = txtEje2KeyPress
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 240
    Width = 832
    Height = 106
    TabOrder = 2
    object Total: TLabel
      Left = 23
      Top = 11
      Width = 32
      Height = 13
      Caption = 'TOTAL'
    end
    object cmdAdd: TButton
      Left = 209
      Top = 9
      Width = 57
      Height = 25
      Caption = '&Agregar'
      TabOrder = 0
      TabStop = False
      OnClick = cmdAddClick
    end
    object cmdEdit: TButton
      Left = 272
      Top = 9
      Width = 57
      Height = 25
      Caption = '&Editar'
      TabOrder = 1
      TabStop = False
      OnClick = cmdEditClick
    end
    object cmdDelete: TButton
      Left = 449
      Top = 9
      Width = 54
      Height = 25
      Caption = '&Borrar'
      TabOrder = 2
      TabStop = False
      OnClick = cmdDeleteClick
    end
    object cmdPost: TButton
      Left = 331
      Top = 9
      Width = 57
      Height = 25
      Caption = '&Grabar'
      Enabled = False
      TabOrder = 3
      TabStop = False
      OnClick = cmdPostClick
    end
    object cmdCancel: TButton
      Left = 390
      Top = 9
      Width = 57
      Height = 25
      Caption = '&Cancelar'
      Enabled = False
      TabOrder = 4
      TabStop = False
      OnClick = cmdCancelClick
    end
    object tdTotal: TRxCalcEdit
      Left = 70
      Top = 9
      Width = 104
      Height = 22
      Margins.Left = 4
      Margins.Top = 1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      NumGlyphs = 2
      ParentFont = False
      ReadOnly = True
      TabOrder = 5
    end
    object cmdOk: TBitBtn
      Left = 527
      Top = 8
      Width = 75
      Height = 25
      Caption = 'OK'
      ModalResult = 1
      TabOrder = 6
      OnClick = cmdOkClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333333333333333333333330000333333333333333333333333F33333333333
        00003333344333333333333333388F3333333333000033334224333333333333
        338338F3333333330000333422224333333333333833338F3333333300003342
        222224333333333383333338F3333333000034222A22224333333338F338F333
        8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
        33333338F83338F338F33333000033A33333A222433333338333338F338F3333
        0000333333333A222433333333333338F338F33300003333333333A222433333
        333333338F338F33000033333333333A222433333333333338F338F300003333
        33333333A222433333333333338F338F00003333333333333A22433333333333
        3338F38F000033333333333333A223333333333333338F830000333333333333
        333A333333333333333338330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
      SkinData.SkinSection = 'BUTTON'
    end
    object cmdExit: TBitBtn
      Left = 608
      Top = 7
      Width = 75
      Height = 25
      TabOrder = 7
      OnClick = cmdExitClick
      Kind = bkCancel
      SkinData.SkinSection = 'BUTTON'
    end
    object txtInicio: TEdit
      Left = 7
      Top = 30
      Width = 23
      Height = 21
      TabOrder = 8
      Visible = False
    end
  end
  object Panel: tNewGroupBox
    Left = 144
    Top = 34
    Width = 433
    Height = 193
    Align = alCustom
    Caption = ' CATALOGO DE PERIMETROS'
    Color = clSilver
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindow
    Font.Height = -12
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
    TabOrder = 3
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
      Width = 425
      Height = 163
      Hint = 'Doble Click para Seleccionar'
      Align = alCustom
      Anchors = [akLeft, akTop, akRight, akBottom]
      Color = 15138559
      Ctl3D = False
      DataSource = ds_perimetros
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
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
      OnKeyPress = ListaObjetoKeyPress
      TitleButtons = True
      Columns = <
        item
          Expanded = False
          FieldName = 'sTipo'
          Width = 40
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'sDimension'
          Width = 56
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dExtPulgada'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dExtMts'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dPerimMts'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dInxIn'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dKg_mts'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dArea_m2'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'dM_mts2'
          Visible = True
        end>
    end
  end
  object dsDespiece: TDataSource
    AutoEdit = False
    DataSet = Q_Despiece
    Left = 624
    Top = 128
  end
  object Q_Despiece: TZQuery
    Connection = connection.zConnection
    AfterPost = Q_DespieceAfterPost
    SQL.Strings = (
      'select'
      '  a.*'
      'from'
      '  estimaciondespiece a'
      'where'
      '  a.scontrato = :contrato and'
      '  a.snumeroorden = :orden and'
      '  a.snumerogenerador = :generador and'
      '  a.swbs = :wbs and'
      '  a.snumeroactividad = :actividad and'
      '  a.sIsometrico = :isometrico'
      'order by'
      '  iIdOrden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'orden'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'generador'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'wbs'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'actividad'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'isometrico'
        ParamType = ptUnknown
      end>
    Left = 584
    Top = 128
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'orden'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'generador'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'wbs'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'actividad'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'isometrico'
        ParamType = ptUnknown
      end>
  end
  object Q_perimetros: TZQuery
    Connection = connection.zConnection
    AfterPost = Q_DespieceAfterPost
    Params = <>
    Left = 584
    Top = 88
  end
  object ds_perimetros: TDataSource
    AutoEdit = False
    DataSet = Q_perimetros
    Left = 616
    Top = 88
  end
end
