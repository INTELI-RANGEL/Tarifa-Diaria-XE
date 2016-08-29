object frmCuadreCategoria: TfrmCuadreCategoria
  Left = 0
  Top = 0
  Caption = 'Cuadre'
  ClientHeight = 503
  ClientWidth = 933
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object dxDock: TdxDockSite
    Left = 0
    Top = 0
    Width = 933
    Height = 503
    Align = alClient
    DockingType = 5
    OriginalWidth = 933
    OriginalHeight = 503
    object dxPanelCuadre: TdxLayoutDockSite
      Left = 201
      Top = 0
      Width = 732
      Height = 503
      DockingType = 0
      OriginalWidth = 300
      OriginalHeight = 200
      object dxLayoutDockSite1: TdxLayoutDockSite
        Left = 0
        Top = 0
        Width = 732
        Height = 503
        DockingType = 0
        OriginalWidth = 300
        OriginalHeight = 200
      end
      object dxDockPanel: TdxDockPanel
        Left = 0
        Top = 0
        Width = 732
        Height = 503
        PopupMenu = popPrincipal
        AllowFloating = True
        AutoHide = False
        CustomCaptionButtons.Buttons = <>
        ShowCaption = False
        TabsProperties.CustomButtons.Buttons = <>
        TabsProperties.TabPosition = tpLeft
        DockingType = 0
        OriginalWidth = 185
        OriginalHeight = 140
        object cxCuadre: TcxSpreadSheet
          Left = 0
          Top = 100
          Width = 724
          Height = 395
          Visible = False
          Align = alClient
          DefaultStyle.Font.Name = 'Tahoma'
          HeaderFont.Charset = DEFAULT_CHARSET
          HeaderFont.Color = clWindowText
          HeaderFont.Height = -11
          HeaderFont.Name = 'Tahoma'
          HeaderFont.Style = []
          PainterType = ptOfficeXPStyle
          Precision = 10
          RowHeaderWidth = 30
          OnClearCells = cxCuadreClearCells
          OnEndEdit = cxCuadreEndEdit
          OnSetSelection = cxCuadreSetSelection
          OnKeyPress = cxCuadreKeyPress
          ExplicitTop = 95
          ExplicitHeight = 400
        end
        object grpCategoria: TcxGroupBox
          AlignWithMargins = True
          Left = 3
          Top = 3
          Align = alTop
          Caption = 'Estado'
          Style.LookAndFeel.NativeStyle = False
          Style.LookAndFeel.SkinName = 'DevExpressStyle'
          StyleDisabled.LookAndFeel.NativeStyle = False
          StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
          StyleFocused.LookAndFeel.NativeStyle = False
          StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
          StyleHot.LookAndFeel.NativeStyle = False
          StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
          TabOrder = 1
          ExplicitLeft = 6
          ExplicitTop = 0
          DesignSize = (
            718
            87)
          Height = 94
          Width = 718
          object dbLblbCategoria: TcxDBLabel
            Left = 135
            Top = 13
            DataBinding.DataField = 'IdCategoria'
            DataBinding.DataSource = dsCategorias
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -16
            Style.Font.Name = 'Tahoma'
            Style.Font.Style = []
            Style.IsFontAssigned = True
            Transparent = True
            Height = 27
            Width = 138
          end
          object cxLabel8: TcxLabel
            AlignWithMargins = True
            Left = 6
            Top = 13
            AutoSize = False
            Caption = 'CATEGORIA'
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -16
            Style.Font.Name = 'Tahoma'
            Style.Font.Style = []
            Style.LookAndFeel.NativeStyle = False
            Style.LookAndFeel.SkinName = 'DevExpressStyle'
            Style.IsFontAssigned = True
            StyleDisabled.LookAndFeel.NativeStyle = False
            StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
            StyleFocused.LookAndFeel.NativeStyle = False
            StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
            StyleHot.LookAndFeel.NativeStyle = False
            StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
            Properties.LineOptions.Visible = True
            Transparent = True
            Height = 27
            Width = 126
          end
          object cxDBLabel1: TcxDBLabel
            AlignWithMargins = True
            Left = 6
            Top = 34
            Margins.Top = 20
            Anchors = [akLeft, akTop, akRight]
            DataBinding.DataField = 'Descripcion'
            DataBinding.DataSource = dsCategorias
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'Tahoma'
            Style.Font.Style = []
            Style.IsFontAssigned = True
            Transparent = True
            Height = 28
            Width = 614
          end
          object dbSolicitado: TcxDBLabel
            Left = 73
            Top = 61
            DataBinding.DataField = 'Solicitado'
            DataBinding.DataSource = dsCategorias
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'Tahoma'
            Style.Font.Style = [fsBold]
            Style.IsFontAssigned = True
            Transparent = True
            Height = 21
            Width = 72
          end
          object cxLabel9: TcxLabel
            Left = 6
            Top = 61
            Caption = 'SOLICITADO'
            Transparent = True
          end
          object dbA_Bordo: TcxDBLabel
            Left = 205
            Top = 61
            DataBinding.DataField = 'A_Bordo'
            DataBinding.DataSource = dsCategorias
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'Tahoma'
            Style.Font.Style = [fsBold]
            Style.IsFontAssigned = True
            Transparent = True
            Height = 21
            Width = 68
          end
          object cxLabel10: TcxLabel
            Left = 155
            Top = 61
            Caption = 'A BORDO'
            Transparent = True
          end
          object cxLabel11: TcxLabel
            Left = 282
            Top = 61
            Caption = 'SUMA'
            Transparent = True
          end
          object dbCantidad: TcxDBLabel
            Left = 398
            Top = 61
            DataBinding.DataField = 'Suma'
            DataBinding.DataSource = dsFolios
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'Tahoma'
            Style.Font.Style = [fsBold]
            Style.IsFontAssigned = True
            Transparent = True
            Height = 21
            Width = 11
          end
          object grpProgreso: TcxGroupBox
            Left = 3
            Top = 15
            Align = alTop
            PanelStyle.Active = True
            Style.BorderStyle = ebsNone
            Style.LookAndFeel.NativeStyle = False
            Style.LookAndFeel.SkinName = 'DevExpressStyle'
            StyleDisabled.LookAndFeel.NativeStyle = False
            StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
            StyleFocused.LookAndFeel.NativeStyle = False
            StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
            StyleHot.LookAndFeel.NativeStyle = False
            StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
            TabOrder = 0
            Height = 40
            Width = 712
            object prgActividades: TProgressBar
              Left = 81
              Top = 20
              Width = 400
              Height = 8
              Color = 10975747
              Position = 0
            end
            object prgFolios: TProgressBar
              Left = 81
              Top = 5
              Width = 400
              Height = 8
              Color = 10975747
              Position = 0
            end
            object prgHorarios: TProgressBar
              Left = 81
              Top = 35
              Width = 400
              Height = 8
              Color = 10975747
              Position = 0
            end
            object cxLabel1: TcxLabel
              Left = 3
              Top = 0
              Caption = 'FOLIOS'
              Transparent = True
            end
            object cxLabel2: TcxLabel
              Left = 3
              Top = 15
              Caption = 'ACTIVIDADES'
              Transparent = True
            end
            object cxLabel3: TcxLabel
              Left = 3
              Top = 30
              Caption = 'HORARIOS'
              Transparent = True
            end
          end
          object lblSuma: TcxLabel
            Left = 319
            Top = 61
            ParentFont = False
            Style.Font.Charset = DEFAULT_CHARSET
            Style.Font.Color = clWindowText
            Style.Font.Height = -11
            Style.Font.Name = 'Tahoma'
            Style.Font.Style = [fsBold]
            Style.IsFontAssigned = True
          end
        end
      end
    end
    object dxpnlCategorias: TdxDockPanel
      Left = 0
      Top = 0
      Width = 201
      Height = 503
      AllowFloating = True
      AutoHide = False
      Caption = 'Categorias'
      CaptionButtons = [cbMaximize]
      CustomCaptionButtons.Buttons = <>
      TabsProperties.CustomButtons.Buttons = <>
      DockingType = 1
      OriginalWidth = 201
      OriginalHeight = 140
      object cxGroupBox1: TcxGroupBox
        AlignWithMargins = True
        Left = 3
        Top = 427
        Align = alBottom
        PanelStyle.Active = True
        Style.BorderStyle = ebsNone
        Style.LookAndFeel.NativeStyle = False
        Style.LookAndFeel.SkinName = 'DevExpressStyle'
        StyleDisabled.LookAndFeel.NativeStyle = False
        StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
        StyleFocused.LookAndFeel.NativeStyle = False
        StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
        StyleHot.LookAndFeel.NativeStyle = False
        StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
        TabOrder = 0
        Height = 43
        Width = 187
        object btnGuardar: TcxButton
          Left = 2
          Top = 2
          Width = 84
          Height = 39
          Align = alLeft
          Caption = 'Guardar'
          TabOrder = 0
          OnClick = btnGuardarClick
          OptionsImage.ImageIndex = 1
          OptionsImage.Images = cxIconosBotones32
        end
      end
      object treeCategorias: TcxTreeView
        AlignWithMargins = True
        Left = 3
        Top = 30
        Width = 187
        Height = 11
        Align = alClient
        TabOrder = 1
        Images = tsIconos
        Items.NodeData = {
          0102000000290000000000000000000000FFFFFFFFFFFFFFFF00000000020000
          000850006500720073006F006E0061006C001F0000000200000000000000FFFF
          FFFFFFFFFFFF00000000000000000331002E0031001F00000002000000000000
          00FFFFFFFFFFFFFFFF00000000000000000331002E0032002700000001000000
          00000000FFFFFFFFFFFFFFFF000000000300000007450071007500690070006F
          0073001F0000000200000000000000FFFFFFFFFFFFFFFF000000000000000003
          34002E0031001F0000000200000000000000FFFFFFFFFFFFFFFF000000000000
          00000334002E0032001F0000000200000000000000FFFFFFFFFFFFFFFF000000
          00000000000334002E003300}
        Indent = 27
        ReadOnly = True
        OnChange = CambioDeNodo
      end
      object cbbCategorias: TcxImageComboBox
        AlignWithMargins = True
        Left = 3
        Top = 3
        Align = alTop
        EditValue = 0
        Properties.Images = cximgCombo
        Properties.Items = <
          item
            Description = 'Personal'
            ImageIndex = 0
            Value = 0
          end
          item
            Description = 'Equipos'
            ImageIndex = 1
            Value = 1
          end>
        Properties.OnChange = cbbCategoriasPropertiesChange
        TabOrder = 2
        Width = 187
      end
      object grpDatosCategoria: TcxGroupBox
        AlignWithMargins = True
        Left = 3
        Top = 245
        Align = alBottom
        PanelStyle.Active = True
        Style.BorderStyle = ebsNone
        Style.LookAndFeel.NativeStyle = False
        Style.LookAndFeel.SkinName = 'DevExpressStyle'
        StyleDisabled.LookAndFeel.NativeStyle = False
        StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
        StyleFocused.LookAndFeel.NativeStyle = False
        StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
        StyleHot.LookAndFeel.NativeStyle = False
        StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
        TabOrder = 3
        Height = 176
        Width = 187
        object dbPlataforma: TcxDBLookupComboBox
          Left = 2
          Top = 64
          Align = alTop
          DataBinding.DataField = 'Plataforma'
          DataBinding.DataSource = dsFolios
          Properties.KeyFieldNames = 'sIdPlataforma'
          Properties.ListColumns = <
            item
              Caption = 'Plataforma'
              FieldName = 'sIdPlataforma'
            end>
          Properties.ListSource = dsPlataformas
          TabOrder = 0
          Width = 183
        end
        object cxLabel4: TcxLabel
          AlignWithMargins = True
          Left = 2
          Top = 47
          Margins.Left = 0
          Margins.Top = 7
          Margins.Right = 0
          Margins.Bottom = 0
          Align = alTop
          Caption = 'PLATAFORMA'
          Transparent = True
        end
        object cxLabel5: TcxLabel
          Left = 2
          Top = 2
          Align = alTop
          Caption = 'PERNOCTA'
          Transparent = True
        end
        object dbPernocta: TcxDBLookupComboBox
          Left = 2
          Top = 19
          Align = alTop
          DataBinding.DataField = 'Pernocta'
          DataBinding.DataSource = dsFolios
          Properties.KeyFieldNames = 'sIdPernocta'
          Properties.ListColumns = <
            item
              Caption = 'Pernocta'
              FieldName = 'sIdPernocta'
            end>
          Properties.ListSource = dsPernoctas
          TabOrder = 3
          Width = 183
        end
      end
      object cxGroupBox2: TcxGroupBox
        AlignWithMargins = True
        Left = 3
        Top = 47
        Align = alBottom
        PanelStyle.Active = True
        Style.BorderStyle = ebsNone
        TabOrder = 4
        Height = 192
        Width = 187
        object cxLabel6: TcxLabel
          Left = 2
          Top = 2
          Align = alTop
          Caption = 'ACTIVIDAD'
          Transparent = True
        end
        object dbActividadDescripcion: TcxDBMemo
          Left = 2
          Top = 19
          Align = alClient
          DataBinding.DataField = 'Descripcion'
          DataBinding.DataSource = dsActividades
          Properties.ReadOnly = True
          Style.LookAndFeel.NativeStyle = False
          Style.LookAndFeel.SkinName = 'Foggy'
          StyleDisabled.LookAndFeel.NativeStyle = False
          StyleDisabled.LookAndFeel.SkinName = 'Foggy'
          StyleFocused.LookAndFeel.NativeStyle = False
          StyleFocused.LookAndFeel.SkinName = 'Foggy'
          StyleHot.LookAndFeel.NativeStyle = False
          StyleHot.LookAndFeel.SkinName = 'Foggy'
          TabOrder = 1
          Height = 171
          Width = 183
        end
      end
    end
  end
  object grpDias: TcxGroupBox
    Left = 575
    Top = 136
    Caption = 'Cuadres Existentes'
    Style.LookAndFeel.NativeStyle = False
    Style.LookAndFeel.SkinName = 'DevExpressStyle'
    StyleDisabled.LookAndFeel.NativeStyle = False
    StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
    StyleFocused.LookAndFeel.NativeStyle = False
    StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
    StyleHot.LookAndFeel.NativeStyle = False
    StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
    TabOrder = 1
    Visible = False
    Height = 122
    Width = 156
    object treeCuadres: TcxTreeView
      Left = 3
      Top = 15
      Width = 150
      Height = 45
      Align = alClient
      TabOrder = 0
      Images = tsIconos
      Items.NodeData = {
        0102000000290000000000000000000000FFFFFFFFFFFFFFFF00000000020000
        000850006500720073006F006E0061006C001F0000000200000000000000FFFF
        FFFFFFFFFFFF00000000000000000331002E0031001F00000002000000000000
        00FFFFFFFFFFFFFFFF00000000000000000331002E0032002700000001000000
        00000000FFFFFFFFFFFFFFFF000000000300000007450071007500690070006F
        0073001F0000000200000000000000FFFFFFFFFFFFFFFF000000000000000003
        34002E0031001F0000000200000000000000FFFFFFFFFFFFFFFF000000000000
        00000334002E0032001F0000000200000000000000FFFFFFFFFFFFFFFF000000
        00000000000334002E003300}
      Indent = 27
      ReadOnly = True
      OnChange = CambioDeFecha
    end
    object btnSeleccionarFecha: TcxButton
      Left = 3
      Top = 86
      Width = 150
      Height = 26
      Align = alBottom
      Caption = 'Seleccionar'
      ModalResult = 1
      TabOrder = 1
      OptionsImage.ImageIndex = 0
      OptionsImage.Images = cxIconosBotones16
    end
    object cxButton2: TcxButton
      Left = 3
      Top = 60
      Width = 150
      Height = 26
      Align = alBottom
      Caption = 'Nuevo Cuadre'
      TabOrder = 2
      OnClick = cxButton2Click
      OptionsImage.ImageIndex = 1
      OptionsImage.Images = cxIconosBotones16
    end
  end
  object grpCrearHorario: TcxGroupBox
    Left = 250
    Top = 276
    PanelStyle.Active = True
    TabOrder = 2
    Visible = False
    Height = 142
    Width = 295
    object lblCaption: TcxLabel
      AlignWithMargins = True
      Left = 5
      Top = 5
      Align = alTop
      AutoSize = False
      Caption = 'NUEVO HORARIO'
      ParentFont = False
      Style.Font.Charset = DEFAULT_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -19
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.LookAndFeel.NativeStyle = False
      Style.LookAndFeel.SkinName = 'DevExpressStyle'
      Style.IsFontAssigned = True
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
      StyleHot.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
      Properties.LineOptions.Visible = True
      Transparent = True
      Height = 27
      Width = 285
    end
    object btnAceptarHorario: TcxButton
      Left = 103
      Top = 94
      Width = 75
      Height = 25
      Caption = 'Aceptar'
      ModalResult = 1
      TabOrder = 1
      OptionsImage.ImageIndex = 0
      OptionsImage.Images = cxIconosBotones16
    end
    object btnCancelarHorario: TcxButton
      Left = 22
      Top = 94
      Width = 75
      Height = 25
      Caption = 'Cancelar'
      ModalResult = 2
      TabOrder = 2
      OptionsImage.ImageIndex = 2
      OptionsImage.Images = cxIconosBotones16
    end
    object cxHoraInicio: TcxTimeEdit
      Left = 22
      Top = 46
      EditValue = 0d
      ParentFont = False
      Properties.SpinButtons.Visible = False
      Properties.TimeFormat = tfHourMin
      Style.Font.Charset = DEFAULT_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.LookAndFeel.NativeStyle = False
      Style.LookAndFeel.SkinName = 'Foggy'
      Style.IsFontAssigned = True
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.SkinName = 'Foggy'
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.SkinName = 'Foggy'
      StyleHot.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.SkinName = 'Foggy'
      TabOrder = 3
      Width = 44
    end
    object cxHoraFinal: TcxMaskEdit
      Left = 72
      Top = 46
      ParentFont = False
      Properties.CharCase = ecUpperCase
      Properties.EditMask = '!90:00;1;_'
      Style.Font.Charset = DEFAULT_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
      TabOrder = 4
      Text = '00:00'
      Width = 42
    end
  end
  object grpAjuste: TcxGroupBox
    Left = 575
    Top = 246
    PanelStyle.Active = True
    TabOrder = 3
    Visible = False
    Height = 177
    Width = 282
    object cxLabel12: TcxLabel
      AlignWithMargins = True
      Left = 5
      Top = 5
      Align = alTop
      AutoSize = False
      Caption = 'AJUSTE'
      ParentFont = False
      Style.Font.Charset = DEFAULT_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -19
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.LookAndFeel.NativeStyle = False
      Style.LookAndFeel.SkinName = 'DevExpressStyle'
      Style.IsFontAssigned = True
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
      StyleHot.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
      Properties.LineOptions.Visible = True
      Transparent = True
      Height = 27
      Width = 272
    end
    object cxLabel13: TcxLabel
      Left = 132
      Top = 56
      Caption = 'Cantidad de Ajuste'
    end
    object btnAplicaAjuste: TcxButton
      Left = 132
      Top = 101
      Width = 75
      Height = 25
      Caption = 'Aplicar Ajuste'
      TabOrder = 2
      OnClick = btnAplicaAjusteClick
    end
    object lstCuentas: TcxListBox
      AlignWithMargins = True
      Left = 5
      Top = 38
      Width = 121
      Height = 134
      Align = alLeft
      ItemHeight = 13
      TabOrder = 3
      OnClick = lstCuentasClick
    end
    object clcAjuste: TcxCalcEdit
      Left = 132
      Top = 76
      EditValue = 0.000000000000000000
      Properties.OnChange = clcAjustePropertiesChange
      TabOrder = 4
      Width = 121
    end
  end
  object dxDockManager: TdxDockingManager
    Color = clBtnFace
    DefaultHorizContainerSiteProperties.CustomCaptionButtons.Buttons = <>
    DefaultHorizContainerSiteProperties.Dockable = True
    DefaultHorizContainerSiteProperties.ImageIndex = -1
    DefaultVertContainerSiteProperties.CustomCaptionButtons.Buttons = <>
    DefaultVertContainerSiteProperties.Dockable = True
    DefaultVertContainerSiteProperties.ImageIndex = -1
    DefaultTabContainerSiteProperties.CustomCaptionButtons.Buttons = <>
    DefaultTabContainerSiteProperties.Dockable = True
    DefaultTabContainerSiteProperties.ImageIndex = -1
    DefaultTabContainerSiteProperties.TabsProperties.CustomButtons.Buttons = <>
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    LookAndFeel.NativeStyle = False
    LookAndFeel.SkinName = 'DevExpressStyle'
    Left = 312
    Top = 120
    PixelsPerInch = 96
  end
  object tsIconos: TcxImageList
    DrawingStyle = dsTransparent
    Items = <
      item
        ImageFormat = ifPNG
        ImgData = {
          89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
          61000000017352474200AECE1CE90000000467414D410000B18F0BFC61050000
          00097048597300000EC300000EC301C76FA864000002C049444154384F85526B
          48D3511CD5565AF4A40F518898364DE6730FE666E2CC547C904ED6CC0796A4CB
          CC576E73F940660AEA74BAF998AE504442A669A46899A369F620CC6998A96405
          25F4F043D2831E6874FA5F1D7D31EDC0E172EFEF9CC3BDBFDFB55A07D662B198
          C666B3B7100A0482CD644FCE57CB1B40A9546E329588DCCCEAA89E914BE14FFB
          1481684DE14FAB63BCEB9343187B29C9FA21C47CAB20F8F87855D4E7F7834A2C
          98CAB0F8A01A731DB9B85D1A034590F330B98D45BE16A4D8230B78F9BA598481
          C2308C6B4598D2446244E187DE0C3EA40227902759E46B418A6D12FE5247861F
          16668D18A849436D3C0B65423768A25D7FA7F3ED97FE1B507DC26B5E779A87E5
          4FD358FE3889AF4F1AF0CA9089A2103A529807E6370C209D9607B9546A137C7E
          7E7F63C4AF77267C1955E3457B1A15E0F42396B14F6799C6BA201DDEA549E0DD
          9FBB918F6F662D3EF45FC4586D1C8AC3E88F49CDA2D910B6F248467AB73418B3
          AD12CC349F41FB393E24FE76D9546DEBAAE4DF20C9B47495072B5FCF33775588
          D110C78246E48E9A240EA41AE6F8A9020697D2D8129D45FF17B42C2DDB4FA1E7
          F495770643DB2D426F6732CC861C0CA813A1927321AB7147618B2FD2AA3C8D09
          72D710CA4382564268991A6661B92114FD13728CBD2D45D7C449C89ABC905874
          1071797638AB72847E3810D727E3D1381801792307F179872B28AF0D09B04955
          792C989ECBD0F32C1E6DE66368190DC09547FED03F3C02DD3D1EEAEE72A11DE2
          A0C6C442FDB02F74438188CD7759A4BC3B48806D520903C5062ECA7A3950DD64
          A36E8403ED301BD52626D477BC5169F4826AD01345D7DC91779581EC26064432
          6750DE9D2480162E71382ACC72AA13E6D04DC21CE72951AE3325A0235A4A3187
          E2854388A21871DE712634D5612428C9FEB28F707F18E55D79C24AF72D1B32A6
          ED14493299F96E0BF758567246AEBD8D22F523ADACFF007FFC4301360EB27800
          00000049454E44AE426082}
      end
      item
        ImageFormat = ifPNG
        ImgData = {
          89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
          61000000017352474200AECE1CE90000000467414D410000B18F0BFC61050000
          00097048597300000EC300000EC301C76FA8640000033A49444154384F7D937D
          50CB7100C67FE814966AE1AAEBBC9C5CA4F4A28CA94B544BCB9026476B975EA8
          4D6D5236D696BB92ACD2B2B763425EF2562DD7B4AB562A7274A5BC5BDC71FED1
          F9873B969778FC563B57EE78EE9E7FBECFF3B9EF3BF197A6FCE589FA772693C9
          A6D6A406A634F0829FE9F7AD1CB99C11F042BD73452E9BCD9EF6BFCC8613444C
          4C8CBD2EC9F7CB4D11E3D78026053D252C54267883C7A27BFC2FB3E104C162B1
          1CCBB72DFDFEA0B61886631C1815FB71B58837FAEEEDDB9CE1E1E1CC668578F4
          79472D9AE529A8DE43470163C18FAD5151F36C384190CBA128D336DC7F615462
          C4DC84FE1A31DA6591963BD2B59F4C7981A3D704615FBB6B8A6131DFC2DD1349
          90C6F97631994C171B4E105C2ED7412DDC21E8D608F0F9512DEA794118ACE6A3
          B32C098F2FE5A3BD781B8A188BF0B1FF12F4F9D1C867D1B2E3E2E266DA70628A
          8F8FCF740663EDF2CA03D116A3BA00EAED4BD028D98C06C9561CDBE20B2D2708
          E2F0F930A824284C0BB68486D256787979D95B59C2D3D373466E6558AA4CB7E1
          59EB4009065ED7A32A3302FAC22458DE3FC0708F122A0E0D727E2CFACC37D070
          4F849CD290E7BB843E7B5C5D5D1D89C3DAF0F6538664989E14E17A6F3A0C0F0B
          51D771147AE9767C7963C2872E3934C9C1A86B294263AF18A7DBD8D01939C855
          AC4182605127917792FEF37A4F3654CDF128D747A3AC8181F30631147931E8BB
          58804EF90E1CE106E2728B742C939EA7234F1B0291960616CFE327917C68B992
          2B5936B85BE2FF52A64AFC5C754108B92EF39BB0943E5A91BF0E25D9AB912609
          FA5176266B4471418083159B47E2F72D1962A4BA3D8ED839573B7606542A75B6
          9393934B2C774140E609DF5A912EE46B63FF5E340D66A1FA2E132AD37A708F2C
          FECECC72BFE6174E5D4DA150E6927D677777F73F3761951DFFA4DFD3B3B713D1
          36C447B3391DE77AA3A0E95903D51D1AD4DD74482FAE4278B28B99ECCE184726
          6B7AEA716F685BC9A7DA1A8AA3067F9477F8436EF24341BD374457BCC03FB510
          216C47905DCA38325976F1424FE9A61C8FA658BEDBA3C88C394311695484719C
          B12AC1E955D016CA13BF8D0EC6A591F6A564D7611C992CEB17B5236D7D20D6BD
          CD22ED38C1D659ADE3D67C2A695204F11B62C69FC73AF8BBC70000000049454E
          44AE426082}
      end
      item
        ImageFormat = ifPNG
        ImgData = {
          89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
          61000000017352474200AECE1CE90000000467414D410000B18F0BFC61050000
          00097048597300000EC300000EC301C76FA8640000020E49444154384F9D935D
          4893511CC6D7CC395DDA45E8242234A2B0615343F163947A6757DD48D0740A4B
          F1222311518CD01421ECA29B3E2CFC80A928DB4A6528F891396683CCD08B6899
          82370A6EF373DA94DA7C3AE73D9B74F4B58B7EF0E35CBCFFE7E170CE7925229C
          204A892122D26FFF446AAD4D7EBCF026170BAF7330DF928DF957D9987B790373
          2FAE63A452D54066689128D2F16A75C38F160D36866E06CCC326515807F3E078
          9E81A1FB971AE92C8BF09C9C6E4AC7FEDE2CF67767E0DFFD02BF770AFE9F9FE0
          DBB6C3E7B1C1B765C5E883CB20B3A12CC2133AD598460293F04C96124BB01DD4
          56028FED2E7EAF59305C7E9116C8588447D6AF8F87C3740FB34F9398CD7FADCD
          6A7CEB2A42CFED585A10C6223C327BDD356C7D6FC37267E611970C9958FFDC84
          C1B2385A2067119EB08F8F52F06BB9176B264DC02CAC1A3501B3E0FDFA0496D2
          F3B4209C4578E4B65A35F6165BB1DA9B0137B52768BAE0CEF4430CE8CFD18208
          16E109B7D624C2EB780677771A5CDDA982EEA05DA9E460CBD1577CF6D882880F
          552AECCCD4C365488193E832241FE8246E8CE9F0AE483844058BF028C62B13E0
          B15760A5430D67C755DEF644AC5B6EC15CA0A405A7588447411FC9E67B3D565A
          5587BC22E836E6C0A48D39B640DE5F7661826EF1AD2E16669D12E642A2560953
          410C8CDA6818EF44C3907F6682CC8ADE027DDFF47E238951229E0EAC342CFA2F
          FC0712C91F749C5F89DC2816090000000049454E44AE426082}
      end
      item
        ImageFormat = ifPNG
        ImgData = {
          89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
          61000000017352474200AECE1CE90000000467414D410000B18F0BFC61050000
          00097048597300000EC300000EC301C76FA864000002F149444154384F85937B
          48935118C63F0BB13B15F457692A0415891425895A59E625B2C8BC5422045219
          69A2660816A2782935AD696ECE29DE9D66A6A299979CF7CD4D4DE7A5FE4A539B
          A6E5A59CDB9CEEE99C6F1225450FFC38E77CEFF33E1FDFCB77983532F807EB7E
          839EFF2A83C42C69E3932C191233A5481074E27186048FF812C4A58B11CBED40
          0CB71D919C6611F1D2A05FA287F504C3F88C4EFC4F114F9BB1EAD7AB2FDD4ED4
          CF3F813EAE1D7943076B5AD42C43C9A2C5828AA0D6E2876A89AD852736D20043
          7D37116DD42D8F6249FD119129ADAC4950DC850C6117F84532A417CAC02B9482
          5720656BF7E3EA698091BE9BA88763039DA61F9A994A3C4C12B1268D7605EA25
          827699AC7A5404AAA0E81A1AB045DF4D244BB2864EDD0BF55409C2E21B581337
          5F8AB4BC4E3C27A4E64A90922B06275BCCD60222AA68C0767D3791E4911574AA
          6EA814390889A9634D2AF2FD143A8B453287453559C91CA8FCC22B68C01EC226
          8211D316750C3AA5180B1F3910A6846328FB3C0633CF6140E00C39DF09729E23
          FAD2CEE25DEA19F4704EA32EC402B5C11678137408D5770F4E31A20747A01C7B
          0545731086F23CB1A29443B73800DD821C2B3FBAB1322FC1F24C0BB4D30D589A
          A88666EC25D4C305503445A1F4C63E19D3106A09853819BD02578C364410732B
          BE4BC3302FB987B9F660CCB50562B62500334DB731F3F626BED5FB62BAF63AA4
          F1A7C0F7348B646A020F415EE48FB6585BCC0E64636130F50FE3F46B1F4C557B
          63AAD20B5FCA3D30517609132F2EA2EAD601F8DBECB2642AFCF64314ED0CF163
          07283F0848932F268971B2CC0D13A517A0283E8FCF42178C173A613CDF0163B9
          F678FFCC1A39EEA68364885B19A18F39CAEF58A29FEF85AF8D41C4E888F13C62
          CCB1C768F6498C66D9E293C0062319C731C2B7C230EF289A43F723D965770209
          D8C008DC8CA5C2AB6628F13647C9357314937DF1153308BD4C5144F1A0EC45A1
          3BE1B209F2DD4C90E56A3CE67378E70112C05EAA8D846D04FA73AC65C71AE8B3
          CD84D5CBC4303F01AB734B0C4DEC8B0E0000000049454E44AE426082}
      end
      item
        ImageFormat = ifPNG
        ImgData = {
          89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
          61000000017352474200AECE1CE90000000467414D410000B18F0BFC61050000
          00097048597300000EC300000EC301C76FA8640000029849444154384F7D92F9
          4B14611C87272A3BB4FB4FA94843A442A45289488D0E138F36152DB3434DCD14
          F3C823CBD514F16833CDAC3C4124D35CCD3214833052F7D0DD9DDD71D7DD2D49
          F770EDD3FBBEBBB5FDA2030FCF7786F93ECC30C3916313C1E33FB6ACC3561774
          DE4CD8C0494382063F840663E84C20064F9FC4C0A9E3E80F0EC0DB407F46EF89
          63E80938826E7F3F741DF545BBDF61BCF1F5C10BEF035212F0E046C282616EAE
          84B9A902C6C67218256530D63F84A1A688A1AF2A80BEE23E748F72A02BC9025F
          98093E3F0DAF7C0E8204B671C3A141309145EDBD0468EEC643931E0B55AA0873
          B7A219B3C991505CBB0479C245C8E2CE635A7416F2E418B47AB3C0768E3EFA42
          5D09D419AEC5DB316C499914C150248493C50B98119DC3744C18A6A242204B8C
          C4CB43FB69C093133F9F80C3F11B36DBEA3FAC7682CB169BC38DD57D5E2619A3
          012FAEECE9380B3CEBF802A5DA0C49FB0473439BD3F5AFC79985821B3068E7A1
          CBBB8E25CB0A8A6A47696007574C068763153295695DF45A017A5EC0BC46C0E2
          B21D795523CE407ED547D8575651DB3A869959136A5A3E3357378F3257367E62
          E6B31321A8056832E3F1E3971DD9E5C334B093CB110FC16E77604A69C477C25A
          D6AAB40C9E605AB421A3F43D0DECE232C9602301B16404DF14463C6E18662EAD
          933217D70C32CFA55C866696275F280A869F56A414F6D1C06E2E950C56129894
          1B30295B58D36A250F1545C163DE6445726EAF33408765EB0A1E540FE0EBB401
          054FFA99F32ADE31E78AFB98E589E198936BD8CFA4335A7035AB8706F6707458
          22015AA508660B04931B7AB396B10CED82139E1097DEED0C44DF6C91C6657493
          0B5D88A5DCF94BA7933432A775E20A2595D201112122A96988043C09E47FA6EF
          426B6EF6AEC13E1774F6E2386EE31F190999F96A7C329F0000000049454E44AE
          426082}
      end
      item
        ImageFormat = ifPNG
        ImgData = {
          89504E470D0A1A0A0000000D49484452000000100000001008060000001FF3FF
          61000000017352474200AECE1CE90000000467414D410000B18F0BFC61050000
          00097048597300000EC300000EC301C76FA8640000023F49444154384F63C002
          18AF749B39DDEE33DB7CA7CFF4E9AD5EE3A7D73B0DB61CAFD57105CA3183E4C1
          AA7000C69BFDE6D58FE67BFEFF7A73C3FFBF5F5FFEFF03C45F80EC7B339DFE9F
          6ED0E806AA6105A903AB46038C173B4D9D1F2EF0F8FBEDE6FCFF7FBE3CFDFFE7
          F313207E0C6683C4EECD71FDBFBB54311AA816E4120CC07CB3D7E4D887E3B9FF
          BF5CACFEFFF3F96E20DE05C75F2E54FD7F7738F3FFE96A955340B51C102DA880
          E5668FF1C7AFB72AFF7F391FF8FFEB8DC2FFDFEF7783F1D79B45FF3F9F0BF8FF
          FD6ED5FFF3F58ABF806AB9215A5001EB8D4E838FDF1F15031506FDFFF5B4FCFF
          EFD75D40DCF9FFE7B38AFFDFEF04FEFFF1A4F4FFB95A7990013C102DA880F552
          BBEE898FA723FF7FBFE1FFFFE7BD0420ED07C610B6FFFFF727C2FE1F29953B03
          548BD505CCDB0A55BDEE4E37FEFFE584D7FF4F072C50F0E7E31EFF6F4F37F8BF
          22453C17A8960DA2051580A286FD60B9F2C487B3F5FEBF5BABF7FFED0A35307E
          B746F7FFC3593AFFB7E7482C01AAE185AAC50A40127C27AB95AE3E9D22F7FFCD
          22B5FF8F5B84FE3F992CFBFF4091F40DA09C3010638D4264C0BE3D5F76D2AD46
          B1FF2FA78AFDBF5FC6F9FF6683C8FFB5A962D380725C1025F8014B746A437CCF
          CCCDFF8F9DB9F97FF2BC6DFF732A67FE77F54D4F06CA61F53B3A60B2770DD329
          6B5E0434E016D880F8ECDEFF5AFAF60620398812FC00140EFC2985135EED3E74
          F17FCFCC4DFF8362EB5E03C584A072440176F7A09C84C8D4B6B741B10D6F2D1C
          A240CEC79A7C71019053411A04A118C4C6E27C06060047813872450D77820000
          000049454E44AE426082}
      end>
    Left = 312
    Top = 152
    Bitmap = {}
  end
  object cxIconosBotones32: TcxImageList
    Height = 32
    Width = 32
    FormatVersion = 1
    DesignInfo = 9961816
    ImageInfo = <
      item
        Image.Data = {
          36100000424D3610000000000000360000002800000020000000200000000100
          2000000000000010000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0001000000020000000400000006000000070000000800000009000000080000
          0007000000060000000400000002000000010000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000001000000030000
          00070000000C00000012000000180000001D0000002000000021000000200000
          001E00000019000000130000000D000000070000000300000001000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000000000000000000100000002000000060000000E0000
          001803020C350E09357B1B1163C0201474D9241683ED271890FF241583EE2014
          74D91A1162C10F09367F03020C380000001A0000001000000007000000030000
          0001000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000001000000040000000B00000016060415451C13
          63BD2C1F96FF2D27A3FF2D31B2FF2D35BAFF2E39BFFF2E3BC4FF2D38BFFF2C35
          BAFF2B2FB3FF2A24A2FF291C93FF1C1162BE06041447000000190000000C0000
          0004000000010000000000000000000000000000000000000000000000000000
          00000000000000000001000000050000000D020105251A1155A4332799FD3336
          B3FF323FC4FF4351CCFF5564D6FF6372DBFF6979DDFF6E7EE0FF6877DDFF6170
          DBFF5261D4FF3F4DCCFF2D3BC3FF2C30B0FF2B2096FD191055AA010105290000
          0010000000060000000100000000000000000000000000000000000000000000
          000000000001000000040000000E05041037281D7DDB3D38ACFF3542C4FF4754
          CEFF6877DCFF6771D3FF5B5EC4FF5150B8FF4D4AB3FF4A44AEFF4F4BB4FF5351
          BAFF5B5EC4FF666FD3FF6373DBFF414ECCFF2F3CC3FF2E29A5FF25197ADD0503
          103B000000100000000500000001000000000000000000000000000000000000
          0001000000030000000C04030D302D2286E54849BAFF3B49C9FF5E6CD7FF6B74
          D3FF4F4CB5FF6A64B9FFA6A2D1FFD5D1E4FFEAE8ECFFF9F6F3FFBEBCC0FFD6D3
          E5FFA8A4D2FF6E69BCFF514FB8FF6770D3FF5665D6FF3240C5FF3131AFFF291E
          84E604030C350000000E00000004000000010000000000000000000000000000
          0002000000090101041D2A2178CE5050BDFF3F4DCAFF6976DAFF5B5DC2FF5F58
          B4FFC0BBD9FFF7F4EEFFF8F5F0FFF8F6F1FFF9F6F2FF72564DFFCAC9C4FFFAF7
          F2FFF9F6F2FFF9F5F1FFC3BEDBFF635CB7FF5A5DC2FF606ED8FF3442C6FF3334
          B0FF261D75D0010104230000000B000000020000000000000000000000010000
          000500000012211B5CA3514FB6FF4755CEFF6673D8FF5453BAFF746DBAFFEEE8
          E9FFF6F3ECFFF8F4EEFFF8F4EFFFF8F4F0FFF9F6F1FF74584FFFCAC7C4FFF9F6
          F1FFF8F5F0FFF8F5F0FFF8F4EFFFEFEAEBFF7871BEFF5555BCFF5B69D6FF3645
          C7FF3734ABFF1F185AA700000016000000060000000100000000000000020000
          000A0C0A1F47443CA7FD616ED4FF5766D3FF5D5DC1FF746CB8FFF3EEE9FFF6F1
          EBFFF6F1ECFFF7F2EEFFF8F3EEFFF8F4EFFFF9F5F0FF765A51FFC9C7C4FFF8F5
          F0FFF8F4F0FFF8F4EFFFF7F3EEFFF7F2EDFFF5F0EBFF7872BCFF5B5DC2FF4A59
          CFFF3945C5FF3A31A1FD0A081E4C0000000D0000000200000000000000040000
          00102C2673BA686BC9FF4E5DD0FF6C72D0FF564DACFFE9E3E3FFF5F0E8FF775B
          51FFF6F2EBFFF6F2ECFFF7F2EDFFF7F3EEFFF7F4EFFF785C53FFC9C6C2FFF8F4
          EFFFF8F4EEFFF7F3EEFFF7F3EDFF755950FFF7F1EBFFEBE5E6FF5B54B0FF656C
          CFFF3D4CC9FF3C40B8FF28216EBD000000140000000500000001000000070706
          122F4943ABFD7783DDFF5E6BD5FF4944AEFFC3BCD3FFF5EEE6FFF5EEE7FFF5F0
          E9FFF6F1EAFFF6F1EBFFF6F2ECFFF7F2EDFFF7F3EDFF7A5F55FFC9C6C1FFF7F3
          EEFFF7F2EDFFF7F2EDFFF7F2ECFFF6F1ECFFF6F0EAFFF5F0E9FFC5C0D6FF4C48
          B2FF4F5ED1FF3F4CCAFF3F38A5FD04040C2C00000009000000010000000A1E1B
          48796868C4FF6A78DAFF666FD0FF5C52ABFFF4ECE4FFF4EEE6FFF5EEE7FFF5EF
          E8FFF6EFE9FFF6F1EAFFF6F1EBFFF6F1ECFFF7F2ECFF7C6157FFC9C5C0FFF7F2
          EDFFF6F1ECFFF6F1ECFFF6F1EBFFF5F0EAFFF5F0E9FFF5EFE8FFF4EFE6FF6158
          B0FF5E65CDFF4351CCFF4A49B8FF17143B710000000D000000030000000D3530
          79B88187D6FF6271D7FF5C5ABDFF998FC0FFF3ECE3FFF4ECE5FFF4EEE5FFF4EE
          E7FFF5EFE8FFF5EFE9FFF5F0E9FFF6F1EAFFF6F1EBFF7E6258FFC9C4BFFFF6F1
          EAFFF6F1EAFFF6F0EAFFF6F0E9FFF5EFE8FFF5EEE8FFF5EEE6FFF4EEE6FF9C93
          C3FF595ABFFF4755CEFF4E54C3FF2E2972B900000011000000040000000E433F
          96DA909AE2FF616FD7FF4641ABFFC9C0D2FFF3EBE2FFF3ECE4FFF3ECE5FFF4ED
          E5FFF4EEE7FFF4EEE7FFF5EFE8FFF6F0E8FFF5F0E9FF5D4941FF9E9A95FFDED8
          D2FFF5EFE9FFF5EFE8FFF5EFE8FFF4EEE7FFF4EEE7FFF4EDE6FFF4EDE4FFCDC6
          D5FF4844AEFF4B59CFFF4D58CAFF3D3792DC00000013000000050000000E4F4C
          ADF29AA5E8FF6170D8FF3C32A2FFE0D8D9FFF3EAE2FFC5BFB8FFC5C0B9FFC5C0
          BAFFC6C0BBFFC6C0BBFFC9C4BFFFC9C7C3FF9A8B83FF74574EFF6E5E58FFC0BE
          BAFFF6F2ECFFF4EFE7FFF4EEE7FFF4EDE6FFF4EDE5FFF3ECE4FFF3ECE4FFE6DE
          DDFF3D34A3FF4F5ED1FF4F5CCFFF4844A7F300000014000000050000000E5856
          BBFFA3AFECFF6878DAFF32289BFFEBE1DBFF6162C8FF5A59C1FF534FB9FF4944
          B2FF413AAAFF3930A3FF32269CFF2B1F96FF775E59FFF1ECE6FF7B6156FFBAB7
          B5FFEDE9E7FFF9F5F1FFF5EDE7FFF4ECE5FFF3ECE4FFF3EBE3FF785C53FFF1E9
          DFFF372B9EFF5462D3FF5766D3FF504DB4FF00000013000000050000000C5654
          B3F2A7B2EAFF7482DEFF392FA0FFDCD2D2FFF2E9E0FFF2EAE1FFF2EAE1FFF2EB
          E3FFF9F6F2FFFAF7F4FFFAF8F5FFFAF8F5FFB8A8A1FF7A5F54FF6F554BFFC6BC
          B8FFCFCCCAFFECEAE8FFF9F6F2FFF3EDE4FFF2EBE2FFF2EBE2FFF2EAE1FFE1D8
          D7FF3A32A1FF5867D4FF606ED5FF504CAEF200000012000000040000000B504F
          A2D9A9B0E8FF8391E3FF4039A6FFC1B5C4FFF1E8DFFFF1E9DFFFF1E9E0FFF6F2
          ECFFFBF8F5FFFBF8F6FFFBF9F6FFFBF9F6FFFBF9F6FFFBF9F6FFD6CECAFF674E
          46FFC6BDB9FFCFCDCBFFEDEBE8FFF7F2ECFFF2EAE1FFF2E9E1FFF1EAE0FFC6B9
          C7FF413BA8FF5C6BD6FF6D77D7FF4B489DDB0000000F00000004000000084444
          87B4A2A7E2FF98A4E9FF5757BAFF8E80AEFFF1E7DDFFF1E8DFFFF1E8DEFFFAF7
          F4FFFCFAF7FFFCFAF7FFFCFAF7FFFCFAF7FFFCFAF7FFFCFAF7FFFCFAF7FFD7CF
          CBFF6A5048FFC8BEBAFFCFCDCBFFECE9E7FFF2E9E0FFF2E8E0FFF1E8DEFF9183
          B1FF5455BAFF6170D7FF7880D6FF3E3C80B50000000C00000003000000062A2B
          52728C8FD8FFB6BFF0FF777ED2FF4D409CFFECE1D5FFF0E7DDFFF1E8DEFFFCFA
          F8FFFCFBF8FFFCFBF8FFFCFBF9FFFCFBF9FFFCFBF9FFFCFBF9FFFCFBF9FFFCFB
          F9FFD8D1CDFF6D534BFFC8C0BCFFD0CECCFFE4DCD3FFF1E8DEFFECE2D6FF5346
          A0FF6972CFFF7483DEFF767AD0FF222145680000000800000002000000030B0C
          15257174CCFCCCD3F5FFA8B4EDFF3F37A3FFAFA0B5FFEFE5DAFFF1E7DDFFFBF9
          F6FFFDFBFAFFFDFBFAFFFDFBFAFFFDFCFAFFFDFBFAFFFDFBFAFFFDFBFAFFFDFC
          FAFFFDFCFAFFD9D2CEFF6F554CFFC8BFBAFFC6BEB6FFE9E0D5FFB2A2B6FF4039
          A6FF8491E2FF9AA5E8FF6A6AC6FC07070E210000000500000001000000020000
          0007545590B3B0B4E7FFC4CDF3FF7D83D3FF423498FFD9CBC2FFEFE6DCFF9C88
          7FFFFDFCFBFFFEFCFBFFFDFCFBFFFEFCFBFFFDFDFBFFFEFCFBFFFEFDFCFFFDFC
          FBFFFEFCFCFFFEFDFBFFDAD2CFFF6B4F45FFD7CDC2FFDBCBC4FF46389BFF7379
          D0FF8E9BE6FF9BA0E1FF4E4F8DB50000000A0000000300000000000000010000
          0004181929388589D6FCD4DAF6FFBBC5F2FF5D5BBAFF5E4F9EFFE2D2C4FFF0E6
          DCFFFCFAF8FFFEFDFDFFFEFDFDFFFEFDFDFFFEFDFDFFFEFDFDFFFEFDFDFFFEFD
          FDFFFEFDFDFFFEFDFDFFFCFAF8FFE0D4CBFFE4D3C4FF6152A0FF5A58B9FF93A0
          E7FFAEB8EDFF7F82D2FC1718283C000000060000000100000000000000000000
          0002000000054B4D7B95A3A8E4FFD9DFF8FFB6C1F1FF4E49AEFF5B4C9DFFD9C8
          BEFFECE1D5FFFBF9F7FFFFFEFEFFFFFEFEFFFFFEFEFFFFFEFEFFFFFEFEFFFFFE
          FEFFFFFEFEFFFBF9F7FFECE2D7FFD9C8BEFF5D4E9FFF4C47ADFF99A5E8FFB0BB
          EFFF9CA0E1FF494B789800000009000000020000000000000000000000000000
          0000000000020303050C676BA6C3B7BBEBFFDBE1F8FFBBC5F1FF605DB9FF4133
          97FFA595ADFFE3D3C1FFF0E7DFFFF9F6F2FFFEFDFCFFBFB2ADFFFEFDFCFFF9F6
          F2FFF0E8DDFFE3D3C1FFA695ADFF413398FF5D5AB9FFA1ABEBFFB6BFF0FFAEB3
          E8FF6468A4C50303051000000004000000010000000000000000000000000000
          000000000001000000030C0C121B787EBFDDB9BEECFFE1E6FAFFC9D2F5FF8E92
          D7FF4038A3FF4A3B99FF8B7BA7FFB9A8B1FFCAB9B5FFE0CEBAFFCAB9B5FFB9A8
          B1FF8B7BA7FF4A3B99FF3F36A3FF878DD5FFABB6EEFFC3CCF3FFB4BAEBFF777A
          BCDD0B0C121E0000000400000001000000000000000000000000000000000000
          00000000000000000001000000030E0E151D757AB6D1AEB3E9FFE3E6F9FFDDE2
          F9FFC4CCF4FF8E92D7FF625FBAFF433AA4FF372C9CFF271991FF372C9CFF423A
          A4FF615EB9FF8B8FD6FFB0BCF0FFBFCAF3FFD6DBF7FFAAB0E8FF7277B4D20D0E
          1520000000040000000100000000000000000000000000000000000000000000
          0000000000000000000000000001000000020505080E53578093949ADEF9C4C8
          F0FFE7EAFAFFE4E8FAFFD6DCF7FFC9D0F4FFC1CAF3FFBAC3F2FFB9C3F2FFBFC9
          F3FFC7D0F4FFD4DBF7FFE0E4F9FFC2C7EFFF9297DEF952557E95050508100000
          0003000000010000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000100000003151620286166
          93A79AA1E5FFB1B7ECFFCED2F4FFDBDFF7FFE2E5F9FFECEFFCFFE1E5F9FFDADE
          F6FFCCD1F3FFB0B5EBFF989FE4FF606493A91516202A00000005000000020000
          0001000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000001000000010000
          00030B0B10163638515D686C9CB07A80B7CE888ECDE5989FE5FF888DCDE57A7F
          B7CE676D9BB13538515F0B0B1017000000040000000200000001000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0001000000010000000200000002000000030000000300000003000000030000
          0003000000030000000200000001000000010000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
      end
      item
        Image.Data = {
          36100000424D3610000000000000360000002800000020000000200000000100
          2000000000000010000000000000000000000000000000000000000000000000
          0000000000000000000100000001000000010000000100000001000000010000
          0001000000010000000100000001000000010000000100000001000000010000
          0001000000010000000100000001000000010000000100000001000000010000
          0001000000010000000100000001000000000000000000000000000000000000
          0001000000010000000300000005000000070000000700000007000000070000
          0007000000070000000700000007000000070000000700000007000000070000
          0007000000070000000700000007000000070000000700000007000000070000
          0007000000070000000600000004000000010000000100000000000000000000
          0001000000060000000E000000160000001A0000001B0000001B0000001B0000
          001B0000001B0000001B0000001B0000001B0000001B0000001C0000001C0000
          001C0000001C0000001C0000001C0000001C0000001C0000001D0000001D0000
          001D0000001C0000001800000010000000070000000100000000000000010000
          00030000000E2D1F198E583E33FD593D34FF583C32FF583C32FF5A3F37FFC58D
          5DFFC18656FFBD8151FFBB7D4DFFB97A4BFFB77748FFB57444FFB37141FFB06D
          3DFFAF6B3AFFAC6837FFAB6535FFA96333FFA76030FF50342DFF4F342CFF4F34
          2CFF4F342DFF4E342DFD2519158B000000100000000400000001000000010000
          0005000000145C3F36FD806357FF745449FF73534AFF735349FF61473EFFE8C5
          94FFE5BE89FFE3B981FFE3B87FFFE3B67DFFE2B57AFFE1B278FFE0B175FFDFAF
          73FFDEAD71FFDEAC6FFFDDAB6DFFDDAA6CFFDDA86AFF523731FF6B4C42FF6B4C
          42FF6B4C42FF6B4D43FF4E342BFB000000170000000600000001000000010000
          0006000000175F4337FF83655AFF77574CFF76574BFF76574BFF654B41FFEAC8
          98FFE6C08CFFE5BC85FFE3B981FFE2B77EFFE2B57CFFE1B479FFE0B277FFE0B0
          75FFDFAF72FFDEAC70FFDEAC6FFFDDAA6CFFDCA96BFF553932FF6B4C43FF6C4D
          42FF6B4C42FF6D4F45FF50362DFF0000001B0000000700000001000000010000
          000600000017624539FF87695DFF79594EFF795A4EFF795A4EFF684E44FFE1C2
          96FFDCB988FFDBB684FFD7B17BFFD7AF79FFD6AD76FFD6AC75FFD6AA72FFD5A8
          70FFD4A66EFFD4A56BFFD3A369FFD3A267FFD3A166FF583D35FF6D4F44FF6D4E
          45FF6C4D44FF6E5047FF51362FFF0000001B0000000700000001000000010000
          00060000001766483CFF8C6E63FF7D5D51FF7D5D50FF7C5C50FF6C5247FFFAF5
          F2FFF9F3F0FFF9F2EFFFF7F1ECFFF7F0EBFFF6F0EAFFF6EEE9FFF5EEE9FFF5EE
          E8FFF4EDE7FFF4ECE7FFF3ECE6FFF4EBE5FFF3EBE4FF5A3F38FF6F5046FF6E50
          46FF6E4F45FF715348FF523830FF0000001A0000000700000001000000010000
          000500000016684B3EFF917366FF816153FF816153FF7F6154FF6E554AFFFAF7
          F3FFF9F4F0FFF9F3EFFFF8F3EDFFF7F0ECFFF6F0EBFFF6EFEBFFF6EFEAFFF5EE
          E9FFF5EDE8FFF5EDE7FFF4ECE7FFF4EBE6FFF4EBE5FF5E4339FF705248FF7052
          48FF6F5147FF72554BFF523A31FF000000190000000600000001000000010000
          0005000000156C4F42FF95786AFF846456FF836557FF836456FF72584DFFFAF7
          F5FFFAF5F2FFD5B8A9FFD2B4A6FFCCA899FFC9A492FFC6A090FFC39C8BFFC097
          87FFBE9283FFBA8F7FFFB88C7BFFF4EDE6FFF4ECE6FF60463DFF72544BFF7053
          4BFF705248FF75574DFF543A32FF000000190000000600000001000000010000
          0005000000156F5346FF997D6FFF87675AFF876859FF876859FF755C50FFFBF8
          F6FFFAF6F2FFFAF5F1FFF9F5F1FFF9F3EFFFF7F2EDFFF7F1EDFFF6F0EBFFF6EF
          EAFFF5EFEAFFF5EEE9FFF5EDE8FFF5ECE8FFF4EDE6FF634940FF73564CFF7355
          4CFF72544BFF775A4FFF563D34FF000000180000000600000001000000000000
          000500000014735649FF9E8275FF8B6B5CFF8A6D5EFF8B6C5DFF796053FFFCF9
          F7FFFAF7F3FFD9BEB1FFD8BCAEFFD5BAABFFCDAC9DFFCBA797FFC8A392FFC6A0
          8FFFC39B8BFFC09787FFBD9283FFF5EEE8FFF4EDE8FF664C43FF75594EFF7458
          4DFF74574CFF795C53FF573D35FF000000180000000600000001000000000000
          00050000001376594BFFA28878FF8E7061FF8F7061FF8E7062FF7C6155FFFCF9
          F8FFFBF7F4FFFAF7F4FFF9F6F3FFF9F5F2FFF9F4F1FFF8F2EEFFF7F1EDFFF7F0
          ECFFF6F0ECFFF6F0EBFFF5EFEAFFF6EEE9FFF4EEE8FF694F46FF775B50FF7759
          4FFF75594EFF7B6056FF594037FF000000170000000600000001000000000000
          0005000000137B5D4EFFA68D7FFF937464FF937464FF917465FF7F6659FFFCFB
          F9FFFCF9F8FFFCF9F8FFFCF9F7FFFBF9F6FFFCF8F5FFF9F6F3FFF9F5F2FFF9F5
          F1FFF9F4F1FFF8F4F0FFF7F2EFFFF7F2EDFFF7F1EDFF6C5248FF785E52FF775D
          51FF775A50FF7F6359FF5C4239FF000000160000000600000001000000000000
          0005000000127E6152FFAB9182FF967767FF957A67FF957968FF876D5EFF8268
          5BFF82695AFF82685AFF81685AFF80675AFF7F6759FF7A6155FF755D51FF745B
          50FF735B50FF725A4FFF72594FFF71594EFF70574CFF72594EFF7B6055FF7A60
          54FF795E52FF81675BFF5D433AFF000000160000000500000001000000000000
          000400000012836654FFB09687FF997B6BFF9A7D6BFF9A7E6BFF9A7D6CFF997D
          6BFF997E6CFF997E6BFF997E6CFF987E6DFF997E6CFF967B6BFF82675BFF7F66
          5AFF806659FF7F6659FF7F655AFF7F6559FF7F6458FF7D6257FF7C6255FF7C61
          55FF7A6054FF826A5EFF5E463CFF000000150000000500000001000000000000
          000400000011866958FFB39B8CFF9D816FFF9D7F6EFF9D816EFF9D816FFF9D81
          70FF9D826FFF9D8270FF9D8270FF9C8170FF9B8170FF9C826FFF8D7364FF8369
          5CFF82695DFF82685CFF81675CFF81675BFF7F6659FF7F6559FF7F6458FF7D63
          58FF7D6257FF866D61FF61483EFF000000140000000500000001000000000000
          0004000000108A6C5BFFB79F91FFA08471FFA08471FFA08571FFA08573FFA085
          73FFA08574FFA08673FFA18673FFA08574FFA08673FF9F8574FF9C8271FF836A
          60FF836B60FF836A5FFF836A5DFF83695DFF82695DFF81685BFF81675AFF7F65
          59FF7F6359FF886E64FF624940FF000000140000000500000000000000000000
          0004000000108D705EFFBBA494FFA48774FFA48874FFAA927FFFAC9582FFAD94
          83FFAD9482FFAC9382FFAC9482FFAC9483FFAB9381FFAA9282FFAA9080FF9880
          71FF8E766BFF8D756AFF8D7568FF8C7468FF897167FF887065FF866E61FF8167
          5CFF806659FF8B7267FF634A40FF000000130000000500000000000000000000
          00040000000F907462FFBFA898FFA68A77FFA78B79FFB09785FF6E5449FF5439
          31FF543831FF60463EFF644A42FF634A40FF61483FFF5F473CFF5E453CFF5C44
          3BFF5B4339FF5A4139FF594137FF583F36FF563E36FF654D43FF887164FF836A
          5DFF82685CFF8D766AFF654C42FF000000120000000500000000000000000000
          00040000000E957764FFC2AD9DFFA98D7BFFAA8F7BFFB29986FF563A33FF5944
          3DFF644E47FF75594EFFE8DAD0FFDCC5B5FFDBC4B5FFDBC4B3FFDAC3B3FFDAC3
          B2FFD9C2B1FFE3D2C6FFE2D1C5FFE2D0C3FFE1CFC2FF573F37FF8A7266FF856C
          5EFF836A5EFF8F796EFF664F44FF000000120000000400000000000000000000
          00030000000E977A67FFC5B0A0FFAC907CFFAD927DFFB39C88FF573B34FF5A45
          3EFF654F48FF775B50FFEBDFD5FFDEC9BBFFDEC9B9FFDDC8B8FFDDC7B7FFDCC6
          B6FF584139FF705448FF705347FF6F5246FFE4D3C7FF594138FF8B7368FF856D
          62FF856B5FFF937B70FF685146FF000000110000000400000000000000000000
          00030000000D9B7D6AFFC8B4A3FFAE947FFFAF9480FFB79E8BFF583C34FF5B46
          3FFF655049FF795D53FFEFE4DBFFE1CEC0FFE1CDBFFFE1CDBDFFE0CBBBFFDFCA
          BBFF554038FF61473FFF654B42FF715548FFE7D8CDFF5B433AFF8C7468FF876F
          62FF876D60FF957E72FF695147FF000000100000000400000000000000000000
          00030000000C9C7F6BFFCAB6A7FFB29681FFB29782FFB89F8CFF593C35FF5B46
          40FF66514AFF7B5F55FFF2E8E0FFE5D3C5FFE4D3C4FFE3D1C2FFE3D0C1FFE2CF
          C0FF523E36FF5B413AFF5E433BFF73564BFFEBDDD4FF5E463CFF8E7469FF8871
          65FF866E64FF968075FF6B5348FF000000100000000400000000000000000000
          00030000000C9E826CFFCDBAAAFFB39983FFB49984FFBBA18DFF593D35FF5C47
          40FF68524BFF7D6157FFF4ECE5FFE8D8CAFFE8D7C9FFE7D6C8FFE6D5C7FFE6D4
          C6FF4E3C34FF553B34FF583E36FF75584DFFEEE2D9FF60483EFF8F766AFF8A72
          66FF896F65FF998377FF6C544AFF0000000F0000000400000000000000000000
          00030000000BA2836EFFCFBCACFFB69B86FFB59B86FFBBA28FFF5A3E36FF5C47
          41FF68524BFF7E6459FFF6F0EAFFEBDCCFFFEADCCEFFEADBCDFFE9DACCFFE8D9
          CBFF4B3A33FF523731FF533832FF765A4FFFF1E6DEFF624A41FF8E786AFF8A73
          67FF897165FF9B8579FF6D564BFF0000000E0000000400000000000000000000
          00030000000AA3876FFFD1BEB0FFB89D89FFB89E87FFBDA390FF5B3E37FF5C48
          42FF69524CFF80665CFFF8F3EEFFEEE1D4FFEDE0D3FFEDDFD2FFECDED1FFEBDD
          D0FF473832FF493832FF493932FF493832FFF3EBE4FF654D43FF90766CFF8D73
          68FF8A7166FF9E877DFF6D564BFF0000000D0000000300000000000000000000
          000200000008A1866EF9D2BFB0FFD3C0B2FFD3C2B1FFD6C5B6FF5B3E37FF5D48
          42FF69534CFF82675DFFF9F5F1FFF9F5F0FFF9F4EFFFF9F4EFFFF8F3EEFFF8F2
          EDFFF7F2ECFFF7F1EBFFF7F1EBFFF6F0EAFFF6EEE9FF664F46FFB4A399FFB3A0
          97FFB19D95FFB09C93FF6E564CFC0000000B0000000300000000000000000000
          00010000000552443984A48873FCA88C75FFA98D76FFA88F78FF836859FF765B
          4DFF765A4DFF80655BFFA9948BFFA89289FFA59087FFA48E85FFA28C82FFA08A
          80FF9E877EFF9C847BFF998278FF987F76FF947C73FF695148FF745F52FF745E
          51FF725B50FF6F584DFC453730A5000000070000000200000000000000000000
          0000000000020000000500000007000000090000000900000009000000090000
          000A0000000A0000000A0000000A0000000A0000000A0000000A0000000A0000
          000B0000000B0000000B0000000B0000000B0000000B0000000B0000000B0000
          000B0000000B0000000A00000007000000030000000100000000000000000000
          0000000000000000000100000002000000020000000200000002000000020000
          0002000000020000000200000002000000020000000200000002000000020000
          0002000000030000000300000003000000030000000300000003000000030000
          0003000000030000000200000002000000010000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
      end>
  end
  object cxImageList1: TcxImageList
    FormatVersion = 1
    DesignInfo = 7864664
    ImageInfo = <
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          20000000000000040000000000000000000000000000000000000000000A0000
          0010000000110000001100000011000000120000001200000012000000120000
          0012000000120000001300000013000000120000000C0000000381594CC2B47C
          69FFB37B69FFB37B68FFB37A68FFB37A68FFB27A68FFB27A68FFB37968FFB279
          68FFB27967FFB27867FFB17867FFB17866FF7F5649C30000000BB77F6EFFFBF8
          F5FFF8EEE9FFF8EEE9FFF7EFE8FFF7EEE8FFF7EEE8FFF7EEE8FFF7EDE7FFF7ED
          E6FFF6EDE6FFF6ECE6FFF6ECE6FFF6ECE5FFB47B69FF00000011B98472FFFBF8
          F6FFBF998AFFEBDAD3FFBE9788FFEBDAD3FFBD9586FFEBDAD3FFBC9484FFEBDA
          D3FF5D6DDDFFE4DDE1FF5A69DCFFF7EDE7FFB77F6EFF00000011BC8978FFFCFA
          F8FFEBDDD5FFECDCD5FFEBDDD5FFECDCD5FFEBDDD5FFECDCD5FFEBDDD5FFECDC
          D5FFE5DFE3FFE5DFE2FFE5DEE2FFF8EEE9FFB98472FF00000010C08E7DFFFCFA
          F9FFC6A294FFEDDED6FFC4A092FFEDDED6FFC29E8EFFEDDED6FFC19B8CFFEDDE
          D6FF6577E1FFE5E0E4FF6272E0FFF8F1EBFFBC8977FF00000010C39482FFFCFA
          FAFFEDDFD9FFEDDFD8FFEDDFD9FFEDDFD8FFEDDFD9FFEDDFD8FFEDDFD9FFEDDF
          D8FFE6E2E6FFE6E2E6FFE6E2E5FFF9F2EEFFC08E7CFF0000000FC79887FFFDFB
          FAFFCCAB9DFFEEE0DBFFCAA99BFFEEE0DBFFC9A799FFEEE0DBFFC8A496FFEEE0
          DBFF6D81E5FFE8E3E8FF6A7DE4FFFAF4F0FFC49381FF0000000EC99D8CFFFDFC
          FCFFEEE2DCFFEEE2DCFFEEE2DCFFEEE2DCFFEEE2DCFFEEE2DCFFEEE2DCFFEEE2
          DCFFE8E6EAFFE8E5EAFFE8E4E9FFFAF6F2FFC69886FF0000000DCDA190FFFEFC
          FCFFD0B1A3FFEFE3DFFFCFB0A2FFEFE3DFFFCFAFA0FFEFE3DFFFCDAD9FFFEFE3
          DFFF7388E8FFE9E6EBFF7186E7FFFBF7F5FFC99D8BFF0000000DCFA594FFFEFC
          FCFFFDF9F9FFFDF9F9FFFDF9F9FFFDFAF8FFFDF9F8FFFDFAF8FFFCF9F7FFFCF9
          F7FFFCF9F7FFFDF8F7FFFCF9F7FFFCF9F7FFCCA290FF0000000C4B53C3FF8D9E
          ECFF687CE3FF6678E2FF6476E1FF6172E0FF5F70DFFF5F70DFFF5D6CDEFF5B69
          DCFF5966DBFF5664DAFF5462D9FF616DDCFF3337AAFF0000000B4C55C4FF93A4
          EEFF6C80E6FF6A7EE4FF687BE4FF6678E2FF6375E1FF6375E1FF6172E0FF5E6F
          DEFF5C6CDDFF5A69DCFF5766DAFF6472DDFF3538ABFF0000000A4D56C6FF96A7
          EFFF95A6EFFF93A4EDFF90A2EDFF8F9FEDFF8B9BEBFF8B9BEBFF8898EAFF8595
          EAFF8291E7FF7F8DE7FF7D89E5FF7987E5FF3539ACFF000000093A4093C14D55
          C5FF4B53C3FF4A51C1FF484FBFFF464DBEFF444BBBFF444BBBFF4249B9FF4046
          B7FF3E44B4FF3C41B3FF3A3EB0FF393CAEFF282B80C200000006000000040000
          0006000000060000000600000007000000070000000700000007000000070000
          0007000000070000000800000008000000070000000500000001}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000090000
          000E000000100000001000000010000000100000001000000011000000110000
          001100000011000000100000000B00000003000000000000000017417CCA2159
          A8FF225BAAFF225AAAFF2159A9FF2158A9FF2057A8FF2057A7FF2055A7FF1F55
          A7FF1F54A6FF1E53A6FF1E52A4FF1A458DE303080F2900000002225DA8FF2F6B
          B0FF579AD3FF71BEECFF46A6E4FF44A3E4FF41A1E3FF3FA0E2FF3C9EE2FF3B9C
          E1FF389BE0FF369AE0FF3498DFFF2875C1FF0F284E8B000000082868B1FF4884
          BFFF4489C7FF9CD8F5FF63B9EBFF55B0E8FF52AEE8FF4EACE7FF4CA9E6FF49A8
          E5FF47A6E4FF44A4E4FF41A2E3FF3991D7FF1B4787D50000000D2C6FB7FF6CA7
          D2FF3C86C4FFA0D4EFFF94D5F4FF66BDEEFF63BBEDFF60B9EBFF5DB6EBFF5BB5
          EAFF57B2EAFF55B0E9FF51AEE7FF4FABE7FF2967B4FF040B142F2F75BCFF8FC7
          E6FF4D9CD0FF7FBCE2FFC3EEFCFF78CAF2FF74C8F1FF72C5F0FF6FC4F0FF6DC2
          EFFF69C0EEFF66BDEEFF63BBEDFF60B9EBFF408ACAFF112C4E81327CBFFFAFE3
          F5FF71C1E6FF56A3D6FFD2F5FDFFD4F6FEFFD2F4FEFFCDF3FDFFC8F1FDFFC2EE
          FCFFBCEBFBFFB5E7FAFFADE3F9FFA5DFF8FF82C0E6FF1E5189CB3582C4FFC7F5
          FEFF92DEF4FF7B93A8FF4CA0D6FF4A9DD5FF489CD4FF479AD2FF4698D2FF4596
          D1FF4394CFFF4292CEFF2D73BAFF2D72B9FF2C71B8FF2765A7EA3688C8FFCDF7
          FEFFA1E6F7FFBA8573FFFFFFFFFFFCF9F7FFFCF9F7FFFCF9F6FFFBF9F6FFFCF8
          F6FFFBF8F6FFFFFFFFFFB17B68FF0000001C0000000A00000007398ECBFFD0F8
          FEFFAAEAF8FFBC8A78FFFFFFFFFFCAA497FFC9A396FFC9A395FFC8A294FFC7A2
          93FFC7A092FFFFFFFFFFB47F6DFF0000001000000000000000003B92CEFFD3F9
          FEFFB2EEF9FFBF8E7DFFFFFFFFFFFDFBF9FFFDFAF8FFFCFBF8FFFCFAF8FFFCFA
          F8FFFCFAF7FFFFFFFFFFB78471FF0000000C00000000000000003D97D1FFE2FC
          FEFFDEF8FAFFC39381FFFFFFFFFFCCA99CFFCCA89BFFCBA79AFFCBA699FFCAA6
          98FFCAA598FFFFFFFFFFBB8776FF0000000700000000000000002E739DBF3E9A
          D3FF3D97D1FFC69785FFFFFFFFFFFCF9F6FFFCF9F5FFFBF9F5FFFBF7F4FFFBF8
          F4FFFAF7F3FFFFFFFFFFBE8C7BFF000000050000000000000000000000020000
          000300000005C99B8AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFC1907FFF000000030000000000000000000000000000
          000000000001977669BECB9E8DFFCA9D8CFFC99C8BFFC89B89FFC89A88FFC799
          87FFC69786FFC59785FF916E61BF000000020000000000000000000000000000
          0000000000000000000100000001000000010000000100000001000000010000
          0002000000020000000200000001000000000000000000000000}
      end>
  end
  object cxIconosBotones16: TcxImageList
    FormatVersion = 1
    DesignInfo = 7864696
    ImageInfo = <
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          00000000000000000002000000070000000C0000001000000012000000110000
          000E000000080000000200000000000000000000000000000000000000000000
          000100000004000101120D2A1D79184E36C6216B4BFF216B4BFF216C4BFF1A53
          3AD20F2F21840001011500000005000000010000000000000000000000000000
          0005050F0A351C5B40DC24805CFF29AC7EFF2CC592FF2DC894FF2DC693FF2AAE
          80FF258560FF1A563DD405110C3D00000007000000010000000000000003040E
          0A31206548ED299D74FF2FC896FF2EC996FF56D4ACFF68DAB5FF3BCD9DFF30C9
          96FF32CA99FF2BA479FF227050F805110C3D00000005000000000000000A1A57
          3DD02EA57CFF33CA99FF2EC896FF4CD2A8FF20835CFF00673BFF45BE96FF31CB
          99FF31CB98FF34CC9CFF31AD83FF1B5C41D300010113000000020B23185E2E8A
          66FF3BCD9EFF30CA97FF4BD3A9FF349571FF87AF9DFFB1CFC1FF238A60FF45D3
          A8FF36CF9FFF33CD9BFF3ED0A3FF319470FF0F32237F00000007184D37B63DB3
          8CFF39CD9FFF4BD5A9FF43A382FF699782FFF8F1EEFFF9F3EEFF357F5DFF56C4
          A1FF43D5A8FF3ED3A4FF3CD1A4FF41BC95FF1B5C43CD0000000B1C6446DF4BCA
          A4FF44D2A8FF4FB392FF4E826AFFF0E9E6FFC0C3B5FFEFE3DDFFCEDDD4FF1B75
          4FFF60DCB8FF48D8ACFF47D6AAFF51D4ACFF247A58F80000000E217050F266D9
          B8FF46D3A8FF0B6741FFD2D2CBFF6A8F77FF116B43FF73967EFFF1E8E3FF72A2
          8BFF46A685FF5EDFBAFF4CD9AFFF6BE2C2FF278460FF020604191E684ADC78D9
          BEFF52DAB1FF3DBA92FF096941FF2F9C76FF57DEB8FF2D9973FF73967EFFF0EA
          E7FF4F886CFF5ABB9AFF5BDEB9FF7FE2C7FF27835FF80000000C19523BAB77C8
          B0FF62E0BCFF56DDB7FF59DFBAFF5CE1BDFF5EE2BEFF5FE4C1FF288C67FF698E
          76FFE6E1DCFF176B47FF5FD8B4FF83D5BDFF1E674CC60000000909201747439C
          7BFF95ECD6FF5ADFBAFF5EE2BDFF61E4BFFF64E6C1FF67E6C5FF67E8C7FF39A1
          7EFF1F6D4AFF288B64FF98EFD9FF4DAC8CFF1036286D00000004000000041C5F
          46B578C6ADFF9AEED9FF65E5C0FF64E7C3FF69E7C6FF6BE8C8FF6CE9C9FF6BEA
          C9FF5ED6B6FF97EDD7FF86D3BBFF237759D20102010C0000000100000001030A
          0718247B5BDA70C1A8FFB5F2E3FF98F0DAFF85EDD4FF75EBCEFF88EFD6FF9CF2
          DDFFBAF4E7FF78CDB3FF2A906DEA0615102E0000000200000000000000000000
          0001030A07171E694FB844AB87FF85D2BBFFA8E6D6FFC5F4EBFFABE9D8FF89D8
          C1FF4BB692FF237F60CB05130E27000000030000000000000000000000000000
          000000000001000000030A241B411B60489D258464CF2C9D77EE258867CF1F71
          56B00E3226560000000600000002000000000000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0002000000080000000C0000000D0000000D0000000D0000000E0000000E0000
          000E0000000E0000000D00000009000000020000000000000000000000000000
          0007295B8FBF357DC6FF337AC5FF3079C3FF2E76C2FF2C74C0FF2971BEFF276F
          BEFF256EBCFF236CBAFF174C86C1000000090000000000000000000000000000
          000A3C83CAFF87CAF0FF66B9EBFF62B6EAFF5FB4E9FF5CB2E8FF59B1E8FF56AF
          E8FF53AEE7FF50ACE7FF246CBBFF0000000D0000000000000000000000000000
          000A3F85CCFF8FCFF1FF4690D2FF438DD0FF408BCEFF3D88CEFF3B86CCFF3984
          CBFF3683C9FF55AFE7FF256EBDFF0000000D0000000000000000000000000000
          000A4389CEFF98D3F3FF4A92D4FFFAF7F3FFF5EFE9FFF5EFEAFFF5EFEAFFF5EF
          EAFF3A85CBFF59B1E8FF2971BEFF0000000D0000000000000000000000000000
          0009468CD0FFA0D7F3FF4F97D5FFFBF8F6FFF6F0EBFFF6EFEAFFF5F0EBFFF5F0
          EAFF3D87CCFF5EB5E9FF2B73C0FF0000000C0000000000000000000000000000
          00094A8FD2FFA9DBF5FF5399D6FFFCFAF8FFF6F1ECFFF6F0ECFFF7F0ECFFF6F1
          EBFF408BCFFF65B6E8FF3275BEFF0000000B0000000000000000000000000000
          00084D92D3FFB0DFF7FF569CD8FFFDFBFAFFF7F2EDFFF7F1ECFFF7F1EDFFF7F1
          EDFF478CCBFF6BB2DEFF3876BBFF0000000B0000000000000000000000000000
          00075094D6FFB8E2F7FF5B9FDAFFFDFCFBFFF8F2EEFFF8F2EEFFF7F2EDFFF0E8
          E3FF7FA9D2FF98C6E2FF729BC9FF0000000A0000000000000000000000000000
          00075398D7FFBDE6F8FF5EA3DCFFFEFDFDFFF8F3F0FFF8F3F0FFEDE5E0FFE2D6
          D0FFABC4DCFF6296CCFFA2BCD8FF09315F8B0000000000000000000000000000
          0006559AD9FFC3E9F9FF62A6DEFFFEFEFEFFF9F4F1FFEAE1DCFFE5D9D5FFEEE6
          E3FFAFC4D9FF1664B9FF3878BFFF1462B7FF0000000000000000000000000000
          0006589DDAFFC9EBFAFF66A9E0FFFFFFFEFFD0BEB7FFBBA298FFCCB8B2FF7E9E
          C6FF1C6ABEFF80D2F8FF4EAEE9FF5CC0F8FF1967BCFF0C335C7E000000000000
          00055BA0DDFFCDEDFBFF69ACE1FFAC8E83FF946C5DFF926A5CFFAD8E84FFCDBB
          B4FF4985C4FF6DBDEDFF6ECBF8FF58B1EAFF185290BD00000000000000000000
          00034678A5C05C9FDCFF599EDCFFB38D7EFFF1E9E2FFE2D2C6FFE9DCD2FF7A98
          BEFF2879CAFFABE4FCFF89C9EFFFA7E2FBFF2576C7FF133B6481000000000000
          0001000000030000000400000005866A5FBEB58F80FFB58F80FFB48E7FFF8569
          5EBF000000062C7FCEFF215F9AC12B7ECDFF0000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000001842687E000000001841687E0000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000002000000090000000B000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000000000020000000E542F18A08A491EDD000000000000
          0000000000000000000000000000000000010000000500000008000000030000
          000000000000000000030000000E553019A0C79A6AFF975125EC000000000000
          0000000000000000000000000002000000092D180B63A06136EC0000000A0000
          0004000000091009052C683C1FBAC69561FFDFC295FF7D4824C8000000000000
          000000000001000000040F09042762371BB7B07A4AFFB27243FF140C07362D1B
          0F5B5C371EA69D643BF9D2A66FFFEECD94FFCCA37AFF472B177B000000000000
          0002000000073922126B996239F3D7AE77FFF3D597FFB57A4CFFA66D42FFB47F
          51FFCEA26EFFEECC90FFEFCD92FFEEDEB6FFA36C43E90805031A00000002150E
          0829764B2CBFC19262FFF2D49AFFF5DA9FFFF6DA9DFFF4D79DFFF4D79CFFF4D6
          9AFFF3D59AFFF3D89FFFF5ECC5FFCAA177FF3825175E0000000500000003B185
          5DEBEADDBCFFFBF7D4FFFCF3CCFFFCF3CEFFFCF2CAFFFAECC0FFF8E6B6FFF6E3
          B2FFF6ECC3FFF1EBCAFFCDA981FF5139257D0000000800000001000000011C16
          0F2A947251C4DABE99FFFAF7D8FFFDFAD9FFFDF7D4FFFDFDE1FFF5EFD0FFEADB
          BAFFD6B892FFAE825CE83D2D1F5E000000070000000100000000000000000000
          0001000000054D3D2C6AC19D78F2ECDFBDFFFEFDDFFFB48359FFAB845EE18A6A
          4CBB4E3C2A700705041300000004000000010000000000000000000000000000
          0000000000000000000315110D21896E51B3DABF9AFFBC8D64FF0000000B0000
          0005000000030000000100000000000000000000000000000000000000000000
          0000000000000000000000000001000000044336285BB08865E80706040E0000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000200000003000000010000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          00000000000000000001000000050000000A0000000D0000000E0000000D0000
          000A000000050000000100000000000000000000000000000000000000000000
          0000000000030000000A0B0C3065191A67BF1F2285EE222492FF1F2085EE1719
          66BF0B0B30650000000B00000003000000010000000000000000000000000000
          000403040D241A1C67B83439A5FF515CC7FF606DD6FF6574DEFF5E6CD6FF4E5A
          C6FF3034A2FF181A65BA03030C25000000040000000100000000000000030606
          1632292B8FEA525BC1FF6574DBFF535DCBFF484FC0FF444CBEFF4750C1FF535E
          CBFF6371DAFF4851BDFF23278BEB04040E2700000003000000000101020D2629
          84D5636ECBFF6472D7FF464DBEFF999CD9FFDBDCF1FFFEFDFDFFDBDBF1FF999C
          DAFF484EC0FF626ED7FF515CC4FF1C1E69B80000000A0000000112143D694A52
          B8FF7887E0FF4A50BEFFD2D3EBFFFCFBFAFFFDFBF9FFFCFBFAFF73584EFFE2E1
          E0FFD5D5ECFF4B52C0FF6B7ADDFF3E46AFFF0F10366500000004262A79BB7681
          D5FF5964CBFFA1A2D8FFFBF8F5FFFCF8F6FFFBF8F6FFFBF8F6FF755950FFE1DF
          DDFFFBF8F6FFA2A3D9FF5761CBFF606CCEFF222574BD00000008363AA2EC92A0
          E7FF454CBCFFE1DEEAFFFAF5F2FFF9F5F2FFFAF5F2FFFAF5F2FF775C51FFE0DC
          D9FFFAF5F2FFE2DFEAFF464DBEFF7485DFFF2E3398EA000000093E45B4FFA1AF
          EEFF3F43B7FFF4EEEDFFF8F2EDFFE0DCD8FFE3E2E1FFE5E5E5FF785D53FFE0DC
          D8FFF7F2EDFFF4EEECFF4146B8FF8193E7FF363BA9FA000000093B42AAECA1AE
          EBFF454BBAFFDED8E2FF4945B2FF433CACFF3B33A5FF352B9EFF5B4770FFF8F7
          F6FFF6EFE9FFDED9E2FF464BBCFF8595E5FF343BA2EA00000008313685B9919C
          E1FF6770CCFF9492CBFFF5ECE6FFFEFDFCFFFFFFFFFFFFFFFFFFFFFFFFFFFEFD
          FCFFF4ECE5FF9593CDFF5D68CCFF7D8ADCFF2C3280BA000000061A1C44607982
          D6FFA5B3EBFF4246B7FFC5C0D7FFF5EEE7FFFBF8F5FFFEFDFCFFFBF8F6FFF5EE
          E7FFC6C0D7FF4247B7FF8E9FE7FF6B73CFFF17193F6000000003010102085A61
          AACF9FAAE8FF96A1E3FF4044B5FF8C89C6FFCDC6D7FFF0E6DFFFCDC6D7FF8C8A
          C7FF3F44B6FF8593DFFF949EE3FF464A8CB40000000600000001000000011112
          1E296F75C4E99AA5E6FFADBAEDFF6A74CCFF444AB8FF383CB0FF454AB8FF636C
          CBFFA1B0EBFF959EE3FF6269BDE90A0B141F0000000100000000000000000000
          00010A0B121A535894B18690DEFFACB7EDFFBDC9F5FFC5D2F8FFBDC8F5FFAAB6
          ECFF828BDCFF4E5491B2090A111B000000020000000000000000000000000000
          000000000001000000032527404F494D7F9A6168ABCC7780D4F96066ABCC474C
          7F9B23263F500000000400000001000000000000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00020000000900000012000000180000001A0000001A00000018000000100000
          0005000000010000000000000000000000000000000000000000000000020000
          000D3524146A936338E5A56B3AFFA36938FFA16736FF9D6233FB633E20B70805
          022800000006000000010000000000000000000000000000000000000008442F
          1D78C18B59FEE1AC76FFE4C296FFB5793BFFB5793CFFB5793CFFAD7239FF7E50
          2AD80302042A00000006000000010000000000000000000000000000000DB07D
          4EF3E6B17AFFE9B47DFFE9B47DFFE7C79DFFB67A3DFFB57A3DFFB57A3DFF6953
          7BFF090E5ED50001052800000006000000010000000000000000000000086A4E
          329DEFD7B3FFE9B47DFFE9B47DFFE9B47DFFEACDA4FFB57B3EFF735C86FF313F
          CCFF2935B8FF0B1161D501010627000000050000000100000000000000010000
          000C745538A5F2DDBBFFE9B47DFFE9B47DFFE9B47DFFD1CEE1FF3443CEFF3443
          CDFF3443CEFF2C39BAFF0D1463D4010106260000000500000001000000000000
          00020000000B76583BA4F5E2C1FFE9B47DFFB5A9B8FF829FF1FFB1C9F5FF3949
          D1FF3A4AD1FF3A49D1FF303FBDFF111767D30101062500000005000000000000
          0000000000010000000B785B3DA3E9E1D2FF87A3F2FF87A4F1FF87A3F2FFB9D0
          F7FF3E50D5FF3E50D5FF3F50D5FF3545C2FF141B6AD201010622000000000000
          000000000000000000010000000A2C386FA2C9E2F9FF8CA8F3FF8DA8F3FF8CA8
          F3FFC0D8F9FF4457D9FF4356D9FF4456D9FF3949C2FF141A61C2000000000000
          000000000000000000000000000100000009303D74A1CFE7FBFF92ADF3FF91AD
          F4FF92ADF4FFC6DEFAFF495EDBFF495DDCFF475AD7FF232F8BF0000000000000
          00000000000000000000000000000000000100000008334177A0D4ECFCFF97B2
          F5FF97B2F4FF97B3F5FFCCE4FBFF4A5FDAFF3141A4F6090C214A000000000000
          000000000000000000000000000000000000000000010000000736457A9FD8F0
          FDFF9DB7F5FF9CB7F5FFD9F1FEFF6B81CAF50B0E234700000006000000000000
          0000000000000000000000000000000000000000000000000001000000063947
          7D9EDBF3FEFFDBF3FFFF677FCFF513192C440000000500000001000000000000
          0000000000000000000000000000000000000000000000000000000000010000
          00053543728E4F63AACD151A2D40000000040000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0001000000030000000400000002000000000000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000020000000A170D0738542D1894814626D193502AEA924F2AE87F45
          25D0522C17931209053000000009000000010000000000000000000000000000
          00030201011159311B97A96239FAC58957FFD6A36DFFDDAF75FFDDAF74FFD6A4
          6BFFC58956FFA46137F53C2112730000000F0000000300000000000000020201
          0110744226B9BC7C4DFFDDAE77FFDEB076FFDEAF75FFDEAF75FFDEB074FFDDAF
          75FFDEAF74FFDBAB72FFBD7E4EFF6F3E24B50000001000000002000000085C36
          2095BE8053FFE0B37CFFDFB076FFDEB177FFDFB279FFE0B379FFE0B27AFFE0B2
          79FFDFB279FFDFB277FFDEB077FFC08253FF55321D920000000A190F0932B070
          47FADFB27DFFDFB27AFFE0B37BFFE0B57DFFE1B67EFFE2B67FFFE2B77FFFE2B7
          7FFFE2B67EFFE0B47CFFE0B47BFFDEB079FFB3734AFB130B072F613C2795CD9B
          6FFFE2B780FFE5BD89FFE7C291FFE8C393FFE8C494FFE8C594FFE8C495FFE8C4
          95FFE8C494FFE8C393FFE5BF8CFFE1B77FFFD09C6EFF5434218B935E3DD2DCB3
          83FFE3B781FFBA8659FFA97043FFAB7245FFAC7346FFB0794AFFAB7245FFAD75
          47FFB0784AFFB17A4BFFC29162FFE4B983FFDEB17EFF8E5B3BD0B0744CF2E3BF
          8FFFE4BB84FFA56B3FFFF5EEE9FFFAF6F3FFFAF7F3FFFBF7F4FFFBF7F5FFFAF7
          F4FFFAF7F3FFFAF6F2FFAB7245FFE5BD87FFE5BE8BFFAB714CEEAE764FECE9C9
          A0FFE5BE89FFA56B3FFFE6D9D2FFE7DBD4FFE9DED7FFEAE0D9FFEAE0DAFFEBE1
          DBFFEBE0DBFFEEE5E1FFAA7144FFE7C08CFFEACA9DFFAE764FEE9A6A49D0E9CD
          ACFFEAC796FFB78456FFA56B3FFFA56B3FFFA56B3FFFA56B3FFFA56B3FFFA56B
          3FFFA56B3FFFA56B3FFFB78457FFEACA99FFEBD1ADFF996A49D46E4E3697DDBB
          9DFFEED3A9FFEECFA2FFEED2A5FFF0D6A9FFF1D7ABFFF1D8ADFFF1D8ADFFF1D8
          ADFFF1D6AAFFF0D5A8FFEED2A5FFEFD4A7FFE0C2A2FF6246318F1C140E2BC794
          6CFCF5E8CCFFEFD6ABFFF1D8AEFFF2DAB0FFF3DCB3FFF3DEB4FFF3DEB4FFF3DE
          B4FFF3DCB2FFF1DBB0FFF1D8ADFFF7EACDFFC69470FA1A120D2E000000036F52
          3C92D7B08CFFF8EFD3FFF3E0B9FFF3DFB7FFF4E1B9FFF5E3BBFFF5E2BBFFF5E2
          BBFFF4E1B9FFF4E2BDFFFAF1D5FFD9B390FF664B368C00000006000000010202
          0107906C4EB8D9B38FFFF7EDD3FFF8EED0FFF7EBC9FFF6E8C4FFF6E8C5FFF7EC
          CAFFF8EED0FFF4E8CDFFD7AF8BFF88664AB30202010B00000001000000000000
          00010202010770543F8FCFA078FCE2C4A2FFEBD7B8FFF4E9CDFFF4EACEFFECD8
          B9FFE3C5A3FFC59973F24C392A67000000060000000100000000000000000000
          000000000001000000022019122C6C543E89A47E5FCCC59770F1C19570EEA47E
          60CD6C543F8B16110D2200000003000000010000000000000000}
      end>
  end
  object cximgCombo: TcxImageList
    FormatVersion = 1
    DesignInfo = 9961848
    ImageInfo = <
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000010000
          000300000006000000080000000B0000000D0000000E0000000F0000000F0000
          00100000000F0000000E0000000C00000009000000050000000100000003081A
          1341143F2E921B543CBD1F5E43D3257553FF247353FF247251FF247151FF2370
          51FF23704FFF1C593FD6184E37C1123A29990717104900000006000000072A82
          60FF36A780FF34B186FF33B589FF34C192FF33BF90FF31BD8DFF30BC8CFF2FBA
          8AFF2EB989FF2CAA7DFF2AA378FF2A976FFF257353FF0000000D000000062A7E
          5FE973D9BDFF5AD4B0FF43CCA1FF40C99DFF3EC79AFF3BC598FF38C295FF36C0
          92FF34BE90FF32BD8EFF3ABF93FF42BE95FF247153EE0000000C000000031845
          357B48A787FF74CEB4FF83E5CBFF57CEABFF2F9773FF207D5DFF1B7859FF2189
          66FF42BF98FF5AD1ADFF4BB996FF349674FF1540308800000006000000000000
          00030C221A3C256B53B345A987FF53AC90FF63A9B1FF71A7CCFF5D8CB7FF3674
          89FF2C8E73FF3A9D7BFF22664DBA0B2019460000000700000001000000000000
          00000000000100000003091B15322D7A61CC5588ABFF325994FF2C538FFF3460
          8EFF28735CCD091A143600000007000000030000000100000000000000000000
          00000000000000000000000000071221316B4572ACFF659FD7FF629CD6FF3968
          A5FF0C1B2C6F0000000800000000000000000000000000000000000000000000
          000000000000000000010000000E2B466EC379AFDAFF90CCF5FF77B4E8FF5991
          CBFF1D3A66CA0000000E00000001000000000000000000000000000000000000
          0000000000000101010315253F8A36598DF8BAE1F6FFBDE6FCFF8CC9F2FF69A5
          DBFF21447AF70D19349501010103000000000000000000000000000000000000
          000000000000010101052A4B7DE2385F95FFD5F0FBFFD1EDFBFF94CFF3FF6DA7
          DDFF24467DFF172C59E201010105000000000000000000000000000000000000
          0000000000000101010538649EFA3C6CA8FFBAD7E9FF698EB7FF325A91FF2B50
          86FF28518FFE1D3869FA01010105000000000000000000000000000000000000
          00000000000001010104386498DC539CE0FF497BB7FF5390CDFF4E8FD3FF3C76
          C1FF396CB1FF223F72FF01010105000000000000000000000000000000000000
          000000000000010101021525374D4B84C2F17FB9E7FF86BDE9FF8DC4EEFF75A8
          DAFF5683B8FF1C3359BF01010104000000000000000000000000000000000000
          00000000000000000000010101020E16202B37618DB03D6B9FD24276B3FF3054
          85D51C3251930305071200000001000000000000000000000000000000000000
          0000000000000000000000000000000000010101010301010105010101050101
          0104010101030000000100000000000000000000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          00000000000100000004000000070000000A0000000B0000000D0000000D0000
          000D0000000C0000000A00000007000000020000000000000000000000000000
          0000000000040E38297E134F37B414593DD212583BDB126441FF105D3DFF0D49
          2FDC0A432AD4063521B804251689000000080000000000000000000000000000
          000000000006278A63FF4EBB98FF3CBB90FF2FB586FF2FBF8DFF2DBD8BFF27AE
          7DFF24A878FF209A6CFF0D5535FF0000000C0000000000000000000000000000
          000000000004267B5BD557C09FFF6CDBBCFF65DAB7FF45CBA0FF39C697FF37C4
          95FF3AC396FF36A47EFF0F5136DA000000090000000000000000010101090101
          010E0101011106120E301D5E45A7349875F868CEB1FF6AD1B4FF4BC59EFF42BD
          95FF237F5EFA114732AA040E0A3801010115010101100101010A194737942266
          4ED1286A54D9548F7AF471A392FF679988F93A8870FC40768EFF386E87FF327B
          63FB679988F770A391FF4D8A74F7205E47DB14523BD30D3627983A9275FD6EC6
          ABFF52B795FF46B08DFF47AF8CFF79B5A1FF487591FF7DA4CDFF588BC1FF315F
          81FF7EBAA6FF47AF8CFF3AA17EFF359875FF379878FF1A6B4EFD2F725CBE5EB5
          9AFF84D4BDFF65CBABFF80CFB5FF86BFADFF567EB0FFBDE0F5FF8BC2EBFF345E
          97FF84BAADFF79CBB1FF55BC9BFF5DB89CFF399173FF195540C1010101091838
          2D6339836BD687D0BBFFAEDDD0FF416992FF5C80B0FFCAE8F6FF94C6E9FF375D
          95FF4A6C8BFFACDACDFF78C7AFFF2D745ED7112E25670101010D000000010101
          01060A0C0C2B44708BF67D96B8FF255696FF4F7CB1FF517CAFFF2C5088FF325D
          98FF1E3F79FF7C95B5FF33617CF60A0B0B2E0101010700000001000000000303
          0308273242707097C3FFA7D2F4FF326BAEFF6C9ED1FF5C8CC1FF76A5D3FF5385
          BEFF214784FFA5CEF0FF4676B0FF1822336E0303030A00000000000000000404
          040C354F75C994B8D8FFB4DAF7FF79A2CEFF427BB8F63878BAF13174BBFC3C6D
          A6F885A2C3FFB6DAF6FF5C8FC5FF223A60C20505051100000000000000000404
          040D28528BFA85A7CAFF5D88B5FF7291B8FF6A83A8FC6262636F606060697890
          B5FCA7BFD7FF618BB7FF3A679EFF1F467CFB0606071600000000000000000202
          02072B5996F54778B2FE6197D0FF4E87C6FF275490FF06080B1A020202072B59
          96F54778B2FE6197D0FF4E87C6FF275490FF06080B1A00000000000000000101
          010213263E612C5A93D6326AABF729578DD81529447301010103010101021326
          3E612C5A93D6326AABF729578DD8152944730101010300000000000000000000
          0000000000000000000100000001000000010000000100000000000000000000
          0000000000010000000100000001000000010000000000000000}
      end>
  end
  object dsCategorias: TDataSource
    DataSet = cdCategorias
    Left = 112
    Top = 216
  end
  object zPernoctas: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select sIdPernocta from pernoctan;')
    Params = <>
    Left = 440
    Top = 120
  end
  object zPlataformas: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select sIdPlataforma from plataformas where lStatus = "Activa";')
    Params = <>
    Left = 472
    Top = 120
  end
  object dsPernoctas: TDataSource
    DataSet = zPernoctas
    Left = 440
    Top = 152
  end
  object dsPlataformas: TDataSource
    DataSet = zPlataformas
    Left = 472
    Top = 152
  end
  object popPrincipal: TPopupMenu
    Images = cxIconosBotones16
    Left = 408
    Top = 120
    object btnpopCrearHorario: TMenuItem
      Caption = 'Crear Corte de Horario'
      ImageIndex = 3
      OnClick = PrepararNuevoHorario
    end
    object btnpopEditarHorario: TMenuItem
      Caption = 'EditarHorario'
      ImageIndex = 4
      OnClick = EditarHorario
    end
    object btnpopEliminarHorario: TMenuItem
      Caption = 'Eliminar Horario'
      ImageIndex = 5
      OnClick = EliminarHorario
    end
    object ImprimirEstructura1: TMenuItem
      Caption = 'Imprimir Estructura'
      OnClick = ImprimirEstructura1Click
    end
    object Ajustes1: TMenuItem
      Caption = 'Ajustes'
      OnClick = Ajustes1Click
    end
  end
  object dsActividades: TDataSource
    DataSet = cdActividades
    Left = 48
    Top = 216
  end
  object dsFolios: TDataSource
    DataSet = cdFolios
    Left = 16
    Top = 216
  end
  object dsHorarios: TDataSource
    DataSet = cdHorarios
    Left = 80
    Top = 216
  end
  object cdFolios: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 408
    Top = 248
  end
  object cdActividades: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 416
    Top = 256
  end
  object cdCategorias: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 424
    Top = 264
  end
  object cdHorarios: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 432
    Top = 272
  end
end
