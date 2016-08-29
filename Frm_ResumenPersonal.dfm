object FrmResumenPersonal: TFrmResumenPersonal
  Left = 0
  Top = 0
  Caption = 'Resumen de personal'
  ClientHeight = 489
  ClientWidth = 1211
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  Visible = True
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1211
    Height = 432
    Align = alClient
    TabOrder = 0
    object LTituloGrid: TLabel
      Left = 1
      Top = 73
      Width = 1209
      Height = 16
      Align = alTop
      Alignment = taCenter
      Caption = 
        'RESUMEN DE PERSONAL DEL DIA:  PERTENECIENTE A LA CATEGORIA:  EN ' +
        'FOLIO:  '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clHighlight
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
      ExplicitWidth = 504
    end
    object GroupBox1: TGroupBox
      Left = 1
      Top = 1
      Width = 1209
      Height = 72
      Align = alTop
      TabOrder = 0
      object Label2: TLabel
        Left = 271
        Top = 11
        Width = 42
        Height = 18
        Caption = 'Folio:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label1: TLabel
        Left = 54
        Top = 11
        Width = 49
        Height = 18
        Caption = 'Fecha:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label3: TLabel
        Left = 25
        Top = 39
        Width = 78
        Height = 18
        Caption = 'Categor'#237'a:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object dIdFecha: TDateTimePicker
        Left = 111
        Top = 8
        Width = 146
        Height = 24
        Date = 41509.931811944440000000
        Time = 41509.931811944440000000
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnChange = dIdFechaChange
      end
      object ComboBox: TComboBox
        Left = 319
        Top = 11
        Width = 250
        Height = 24
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnChange = ComboBoxChange
      end
      object cmbCategoria: TComboBox
        Left = 111
        Top = 38
        Width = 458
        Height = 24
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        OnChange = cmbCategoriaChange
      end
    end
    object NextGrid: TNextGrid
      Left = 1
      Top = 105
      Width = 1209
      Height = 326
      Align = alClient
      AppearanceOptions = [ao3DGridLines, aoAlphaBlendedSelection, aoBoldTextSelection, aoHideSelection, aoHighlightSlideCells, aoIndicateSelectedCell, aoIndicateSortedColumn]
      HighlightedTextColor = clBlack
      Options = [goFooter, goGrid, goHeader, goSelectFullRow]
      PopupMenu = PopupMenu1
      TabOrder = 2
      TabStop = True
      OnAfterEdit = NextGridAfterEdit
      object Procesar: TNxCheckBoxColumn
        DefaultWidth = 28
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        Header.Alignment = taCenter
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing]
        ParentFont = False
        Position = 0
        SortType = stBoolean
        Width = 28
      end
      object sidPersonal: TNxTextColumn
        DefaultWidth = 126
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Arial'
        Font.Style = []
        Header.Caption = 'Categor'#237'a'
        Header.Alignment = taCenter
        ParentFont = False
        Position = 1
        SortType = stAlphabetic
        Width = 126
      end
      object sTipoObra: TNxTextColumn
        DefaultWidth = 101
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Arial'
        Font.Style = []
        Header.Caption = 'Tipo de Obra'
        Header.Alignment = taCenter
        ParentFont = False
        Position = 2
        SortType = stAlphabetic
        Visible = False
        Width = 101
      end
      object sDescripcion: TNxTextColumn
        DefaultWidth = 371
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Caption = 'Descripci'#243'n'
        Header.Alignment = taCenter
        ParentFont = False
        Position = 3
        SortType = stAlphabetic
        Width = 371
      end
      object ClsIdPernocta: TNxComboBoxColumn
        DefaultWidth = 147
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Caption = 'Pernocta'
        Header.Alignment = taCenter
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing, coShowTextFitHint]
        ParentFont = False
        Position = 4
        SortType = stAlphabetic
        Visible = False
        Width = 147
      end
      object ClsIdPlataforma: TNxComboBoxColumn
        DefaultWidth = 125
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Caption = 'Plataforma'
        Header.Alignment = taCenter
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing, coShowTextFitHint]
        ParentFont = False
        Position = 5
        SortType = stAlphabetic
        Visible = False
        Width = 125
      end
      object ClFolio: TNxComboBoxColumn
        DefaultWidth = 133
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Caption = 'Folio'
        Header.Alignment = taCenter
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing, coShowTextFitHint]
        ParentFont = False
        Position = 6
        SortType = stAlphabetic
        Width = 133
      end
      object dCantidad: TNxNumberColumn
        Color = 16774120
        DefaultValue = '0'
        DefaultWidth = 58
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Arial'
        Font.Style = []
        Footer.Color = clMoneyGreen
        Footer.FormulaKind = fkSum
        Footer.FormatMask = '0.00'
        Footer.FormatMaskKind = mkFloat
        Header.Caption = 'Cantidad'
        Header.Alignment = taCenter
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing]
        ParentFont = False
        Position = 7
        SortType = stNumeric
        Width = 58
        Increment = 1.000000000000000000
        Precision = 0
      end
      object sAgrupaPersonal: TNxTextColumn
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Caption = 'Agrupa Personal'
        ParentFont = False
        Position = 8
        SortType = stAlphabetic
        Visible = False
      end
      object dCantHH: TNxNumberColumn
        DefaultValue = '0'
        DefaultWidth = 60
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Alignment = taCenter
        ParentFont = False
        Position = 9
        SortType = stNumeric
        Visible = False
        Width = 60
        Increment = 1.000000000000000000
        Precision = 0
      end
      object dSolicitado: TNxNumberColumn
        Color = 16774120
        DefaultValue = '0'
        DefaultWidth = 60
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Caption = 'Cant. Sol.'
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing]
        ParentFont = False
        Position = 10
        SortType = stNumeric
        Width = 60
        Increment = 1.000000000000000000
        Precision = 0
      end
      object sHoraInicio: TNxTextColumn
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        Position = 11
        SortType = stAlphabetic
        Visible = False
      end
      object ShoraFinal: TNxTextColumn
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'ShoraFinal'
        Font.Style = []
        ParentFont = False
        Position = 12
        SortType = stAlphabetic
        Visible = False
      end
      object ClsTipoPernocta: TNxComboBoxColumn
        DefaultWidth = 334
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Caption = 'Tipo pernocta'
        Header.Alignment = taCenter
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing, coShowTextFitHint]
        ParentFont = False
        Position = 13
        SortType = stAlphabetic
        Width = 334
      end
      object NxComboBoxColumn1: TNxComboBoxColumn
        Alignment = taCenter
        DefaultValue = 'Si'
        DefaultWidth = 60
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Header.Caption = 'Imp. Res'
        Header.Alignment = taCenter
        Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing, coShowTextFitHint]
        ParentFont = False
        Position = 14
        SortType = stAlphabetic
        Width = 60
        AutoDropDown = True
        Items.Strings = (
          'Si'
          'No')
        Style = cbsDropDownList
      end
    end
    object ChbHD: TCheckBox
      Left = 1
      Top = 89
      Width = 1209
      Height = 16
      Align = alTop
      Caption = 'Habilitar/Deshabilitar.'
      TabOrder = 1
      OnClick = ChbHDClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 432
    Width = 1211
    Height = 57
    Align = alBottom
    TabOrder = 2
    object Label4: TLabel
      Left = 16
      Top = 6
      Width = 90
      Height = 13
      Caption = 'Ajustar Pernoctas:'
    end
    object pBarraAvance: TProgressBar
      Left = 1
      Top = 39
      Width = 1209
      Height = 17
      Align = alBottom
      TabOrder = 2
    end
    object btnPost: TAdvGlowButton
      Left = 724
      Top = 8
      Width = 137
      Height = 25
      Hint = 'Salvar cambios (F10)'
      Caption = 'Guardar cambios'
      ImageIndex = 8
      Images = connection.ImageList1
      NotesFont.Charset = DEFAULT_CHARSET
      NotesFont.Color = clWindowText
      NotesFont.Height = -11
      NotesFont.Name = 'Tahoma'
      NotesFont.Style = []
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = btnPostClick
      Appearance.ColorChecked = 16111818
      Appearance.ColorCheckedTo = 16367008
      Appearance.ColorDisabled = 15921906
      Appearance.ColorDisabledTo = 15921906
      Appearance.ColorDown = 16111818
      Appearance.ColorDownTo = 16367008
      Appearance.ColorHot = 16117985
      Appearance.ColorHotTo = 16372402
      Appearance.ColorMirrorHot = 16107693
      Appearance.ColorMirrorHotTo = 16775412
      Appearance.ColorMirrorDown = 16102556
      Appearance.ColorMirrorDownTo = 16768988
      Appearance.ColorMirrorChecked = 16102556
      Appearance.ColorMirrorCheckedTo = 16768988
      Appearance.ColorMirrorDisabled = 11974326
      Appearance.ColorMirrorDisabledTo = 15921906
    end
    object EditAjustaPernocta: TNxNumberEdit
      Left = 112
      Top = 5
      Width = 121
      Height = 21
      TabOrder = 0
      Text = '0.00000000'
      Precision = 8
    end
  end
  object PnlDiaAnterior: TPanel
    Left = 92
    Top = 88
    Width = 896
    Height = 338
    TabOrder = 1
    Visible = False
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 894
      Height = 15
      Align = alTop
      TabOrder = 0
    end
    object Panel5: TPanel
      Left = 1
      Top = 16
      Width = 894
      Height = 280
      Align = alClient
      TabOrder = 1
      object GrdPersonal: TNextGrid
        Left = 1
        Top = 1
        Width = 892
        Height = 278
        Align = alClient
        TabOrder = 0
        TabStop = True
        object NxTreeColumn1: TNxTreeColumn
          DefaultWidth = 56
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Position = 0
          SortType = stAlphabetic
          Width = 56
        end
        object NxTextColumn2: TNxTextColumn
          DefaultWidth = 94
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          Header.Caption = 'Folio'
          Header.Alignment = taCenter
          ParentFont = False
          Position = 1
          SortType = stAlphabetic
          Width = 94
        end
        object CbxProcesar: TNxCheckBoxColumn
          DefaultWidth = 28
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          Header.Caption = 'Sel.'
          Header.Alignment = taCenter
          Options = [coCanClick, coCanInput, coCanSort, coEditing, coPublicUsing]
          ParentFont = False
          Position = 2
          SortType = stBoolean
          Width = 28
          OnChange = CbxProcesarChange
        end
        object NxTextColumn3: TNxTextColumn
          DefaultWidth = 51
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          Header.Caption = 'Cant.'
          ParentFont = False
          Position = 3
          SortType = stAlphabetic
          Width = 51
        end
        object NxTextColumn1: TNxTextColumn
          DefaultWidth = 61
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          Header.Caption = 'Id'
          Header.Alignment = taCenter
          ParentFont = False
          Position = 4
          SortType = stAlphabetic
          Width = 61
        end
        object NxTextColumn4: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          Header.Caption = 'Tipo'
          Header.Alignment = taCenter
          ParentFont = False
          Position = 5
          SortType = stAlphabetic
        end
        object NxMemoColumn1: TNxMemoColumn
          DefaultWidth = 462
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          Header.Caption = 'Descripci'#243'n'
          Header.Alignment = taCenter
          ParentFont = False
          Position = 6
          SortType = stAlphabetic
          Width = 462
        end
        object NxTextColumn5: TNxTextColumn
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          Header.Caption = 'tipopernocta'
          ParentFont = False
          Position = 7
          SortType = stAlphabetic
          Visible = False
        end
      end
    end
    object Panel6: TPanel
      Left = 1
      Top = 296
      Width = 894
      Height = 41
      Align = alBottom
      TabOrder = 2
      object LbPlataforma: TLabel
        Left = 413
        Top = 1
        Width = 4
        Height = 13
        Caption = '-'
      end
      object LbPErnocta: TLabel
        Left = 413
        Top = 20
        Width = 4
        Height = 13
        Caption = '-'
      end
      object Label5: TLabel
        Left = 351
        Top = 0
        Width = 56
        Height = 13
        Caption = 'Plataforma:'
      end
      object Label6: TLabel
        Left = 358
        Top = 19
        Width = 47
        Height = 13
        Caption = 'Pernocta:'
      end
      object AdvGlowButton1: TAdvGlowButton
        Left = 684
        Top = 6
        Width = 137
        Height = 25
        Hint = 'Salvar cambios (F10)'
        Caption = 'Guardar cambios'
        ImageIndex = 8
        Images = connection.ImageList1
        NotesFont.Charset = DEFAULT_CHARSET
        NotesFont.Color = clWindowText
        NotesFont.Height = -11
        NotesFont.Name = 'Tahoma'
        NotesFont.Style = []
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = AdvGlowButton1Click
        Appearance.ColorChecked = 16111818
        Appearance.ColorCheckedTo = 16367008
        Appearance.ColorDisabled = 15921906
        Appearance.ColorDisabledTo = 15921906
        Appearance.ColorDown = 16111818
        Appearance.ColorDownTo = 16367008
        Appearance.ColorHot = 16117985
        Appearance.ColorHotTo = 16372402
        Appearance.ColorMirrorHot = 16107693
        Appearance.ColorMirrorHotTo = 16775412
        Appearance.ColorMirrorDown = 16102556
        Appearance.ColorMirrorDownTo = 16768988
        Appearance.ColorMirrorChecked = 16102556
        Appearance.ColorMirrorCheckedTo = 16768988
        Appearance.ColorMirrorDisabled = 11974326
        Appearance.ColorMirrorDisabledTo = 15921906
      end
      object cmbChanFolio: TAdvComboBox
        Left = 81
        Top = 6
        Width = 200
        Height = 21
        Color = clWindow
        Version = '1.3.2.2'
        Visible = True
        DropWidth = 0
        Enabled = True
        ItemIndex = -1
        LabelCaption = 'Aplicar a folio:'
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clWindowText
        LabelFont.Height = -11
        LabelFont.Name = 'Tahoma'
        LabelFont.Style = []
        TabOrder = 0
        Text = 'cmbChanFolio'
        OnChange = cmbChanFolioChange
      end
      object AdvGlowButton2: TAdvGlowButton
        Left = 827
        Top = 6
        Width = 56
        Height = 25
        Caption = 'Cancelar'
        ImageIndex = 8
        NotesFont.Charset = DEFAULT_CHARSET
        NotesFont.Color = clWindowText
        NotesFont.Height = -11
        NotesFont.Name = 'Tahoma'
        NotesFont.Style = []
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = AdvGlowButton2Click
        Appearance.ColorChecked = 16111818
        Appearance.ColorCheckedTo = 16367008
        Appearance.ColorDisabled = 15921906
        Appearance.ColorDisabledTo = 15921906
        Appearance.ColorDown = 16111818
        Appearance.ColorDownTo = 16367008
        Appearance.ColorHot = 16117985
        Appearance.ColorHotTo = 16372402
        Appearance.ColorMirrorHot = 16107693
        Appearance.ColorMirrorHotTo = 16775412
        Appearance.ColorMirrorDown = 16102556
        Appearance.ColorMirrorDownTo = 16768988
        Appearance.ColorMirrorChecked = 16102556
        Appearance.ColorMirrorCheckedTo = 16768988
        Appearance.ColorMirrorDisabled = 11974326
        Appearance.ColorMirrorDisabledTo = 15921906
      end
      object Sustituircambio: TCheckBox
        Left = 280
        Top = 6
        Width = 65
        Height = 17
        Caption = 'Sustituir'
        TabOrder = 1
      end
    end
  end
  object PopupMenu1: TPopupMenu
    OnPopup = PopupMenu1Popup
    Left = 560
    Top = 240
    object DesSeleccionartodo1: TMenuItem
      Caption = 'Des-Seleccionar todo'
      OnClick = DesSeleccionartodo1Click
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object Importardesdecuadre1: TMenuItem
      Caption = 'Importar todo desde cuadre'
      Visible = False
      object Sustituyendo1: TMenuItem
        Caption = 'Sustituyendo cantidades de personal'
        OnClick = Sustituyendo1Click
      end
      object Sumando1: TMenuItem
        Caption = 'Sumando cantidades de personal'
        OnClick = Sumando1Click
      end
      object Agregando1: TMenuItem
        Caption = 'Agregando personal que no exista.'
        OnClick = Agregando1Click
      end
    end
    object Importardesdedaanterior1: TMenuItem
      Caption = 'Importar todo el personal desde d'#237'a anterior'
      OnClick = Importardesdedaanterior1Click
    end
    object Imppersonaldediaanteriorcondiferentefolio1: TMenuItem
      Caption = 'Imp. personal de dia anterior asignando folio.'
      OnClick = Imppersonaldediaanteriorcondiferentefolio1Click
    end
    object mniCambio: TMenuItem
      Caption = 'Cambiar Personal a Folio'
      object TMenuItem
      end
    end
  end
end
