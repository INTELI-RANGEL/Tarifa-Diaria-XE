object frmCalidad_Rir: TfrmCalidad_Rir
  Left = 0
  Top = 0
  Caption = 'Calidad RIR'
  ClientHeight = 562
  ClientWidth = 904
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  DesignSize = (
    904
    562)
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 8
    Top = 8
    Width = 79
    Height = 174
    BevelOuter = bvNone
    TabOrder = 0
    object btnInsertar: TcxButton
      Left = 0
      Top = 0
      Width = 79
      Height = 25
      Align = alTop
      Caption = 'Insertar'
      TabOrder = 0
      TabStop = False
      OnClick = btnInsertarClick
      OptionsImage.ImageIndex = 0
      OptionsImage.Images = cxImgAcciones
    end
    object btnEditar: TcxButton
      Left = 0
      Top = 25
      Width = 79
      Height = 25
      Align = alTop
      Caption = 'Editar'
      TabOrder = 1
      TabStop = False
      OnClick = btnEditarClick
      OptionsImage.ImageIndex = 1
      OptionsImage.Images = cxImgAcciones
    end
    object btnEliminar: TcxButton
      Left = 0
      Top = 50
      Width = 79
      Height = 25
      Align = alTop
      Caption = 'Eliminar'
      TabOrder = 2
      TabStop = False
      OnClick = btnEliminarClick
      OptionsImage.ImageIndex = 2
      OptionsImage.Images = cxImgAcciones
    end
    object btnGuardar: TcxButton
      Left = 0
      Top = 75
      Width = 79
      Height = 25
      Align = alTop
      Caption = 'Guardar'
      Enabled = False
      TabOrder = 3
      TabStop = False
      OnClick = btnGuardarClick
      OptionsImage.ImageIndex = 3
      OptionsImage.Images = cxImgAcciones
    end
    object btnCancelar: TcxButton
      Left = 0
      Top = 100
      Width = 79
      Height = 25
      Align = alTop
      Caption = 'Cancelar'
      Enabled = False
      TabOrder = 4
      TabStop = False
      OnClick = btnCancelarClick
      OptionsImage.ImageIndex = 4
      OptionsImage.Images = cxImgAcciones
    end
    object btnImprimir: TcxButton
      Left = 0
      Top = 125
      Width = 79
      Height = 25
      Align = alTop
      Caption = 'Imprimir'
      TabOrder = 5
      TabStop = False
      OnClick = btnImprimirClick
      OptionsImage.ImageIndex = 5
      OptionsImage.Images = cxImgAcciones
    end
    object btnConceptos: TcxButton
      Left = 0
      Top = 150
      Width = 79
      Height = 25
      Align = alTop
      Caption = 'Conceptos'
      TabOrder = 6
      OnClick = btnConceptosClick
      OptionsImage.ImageIndex = 6
      OptionsImage.Images = cxImgAcciones
    end
  end
  object gridDatos: TcxGrid
    Left = 93
    Top = 8
    Width = 803
    Height = 241
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 1
    object cxGridDatosVista: TcxGridDBTableView
      Navigator.Buttons.CustomButtons = <>
      DataController.DataSource = dsCalidad_rir
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      OptionsData.CancelOnExit = False
      OptionsData.Deleting = False
      OptionsData.DeletingConfirmation = False
      OptionsData.Editing = False
      OptionsData.Inserting = False
      OptionsView.CellAutoHeight = True
      object cxGridDatosVistafecha: TcxGridDBColumn
        Caption = 'Fecha'
        DataBinding.FieldName = 'fecha'
        HeaderAlignmentHorz = taCenter
      end
      object cxGridDatosVistanumero_oc: TcxGridDBColumn
        Caption = 'Orden de Compra'
        DataBinding.FieldName = 'numero_oc'
        HeaderAlignmentHorz = taCenter
        Width = 215
      end
      object cxGridDatosVistanumero_aviso_emb: TcxGridDBColumn
        Caption = 'Aviso de Embarque'
        DataBinding.FieldName = 'numero_aviso_emb'
        HeaderAlignmentHorz = taCenter
        Width = 297
      end
      object cxGridDatosVistasNumeroReporte: TcxGridDBColumn
        Caption = 'Numero Reporte'
        DataBinding.FieldName = 'sNumeroReporte'
        HeaderAlignmentHorz = taCenter
        Width = 141
      end
      object cxGridDatosVistasIdPlataforma: TcxGridDBColumn
        Caption = 'Proyecto'
        DataBinding.FieldName = 'sIdPlataforma'
        HeaderAlignmentHorz = taCenter
        Width = 186
      end
      object cxGridDatosVistasIdEmbarcacion: TcxGridDBColumn
        Caption = 'Lug. Inspecci'#243'n'
        DataBinding.FieldName = 'sIdEmbarcacion'
        HeaderAlignmentHorz = taCenter
        Width = 174
      end
      object cxGridDatosVistadCantidad: TcxGridDBColumn
        Caption = 'Cantidad'
        DataBinding.FieldName = 'dCantidad'
        HeaderAlignmentHorz = taCenter
      end
      object cxGridDatosVistasIdInsumo: TcxGridDBColumn
        Caption = 'Insumo'
        DataBinding.FieldName = 'sIdInsumo'
        HeaderAlignmentHorz = taCenter
      end
      object cxGridDatosVistasInstrumento: TcxGridDBColumn
        Caption = 'Instrumento'
        DataBinding.FieldName = 'sInstrumento'
        HeaderAlignmentHorz = taCenter
        Width = 188
      end
    end
    object gridDatosLevel1: TcxGridLevel
      GridView = cxGridDatosVista
    end
  end
  object grpCaptura: TcxGroupBox
    Left = 93
    Top = 264
    Anchors = [akLeft, akTop, akRight, akBottom]
    Caption = 'DATOS DE CAPTURA'
    Enabled = False
    Style.LookAndFeel.NativeStyle = False
    Style.LookAndFeel.SkinName = 'DevExpressStyle'
    StyleDisabled.LookAndFeel.NativeStyle = False
    StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
    StyleFocused.LookAndFeel.NativeStyle = False
    StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
    StyleHot.LookAndFeel.NativeStyle = False
    StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
    TabOrder = 3
    Height = 287
    Width = 803
    object Label1: TLabel
      Left = 12
      Top = 24
      Width = 29
      Height = 13
      Caption = 'Fecha'
    end
    object Label2: TLabel
      Left = 12
      Top = 56
      Width = 70
      Height = 13
      Caption = 'Orden Compra'
    end
    object Label3: TLabel
      Left = 12
      Top = 88
      Width = 69
      Height = 13
      Caption = 'Aviso Embarq.'
    end
    object Label4: TLabel
      Left = 13
      Top = 120
      Width = 39
      Height = 13
      Caption = 'Reporte'
    end
    object Label5: TLabel
      Left = 13
      Top = 150
      Width = 43
      Height = 13
      Caption = 'Proyecto'
    end
    object Label6: TLabel
      Left = 13
      Top = 183
      Width = 69
      Height = 13
      Caption = 'Lug. Ispecci'#243'n'
    end
    object Label7: TLabel
      Left = 13
      Top = 214
      Width = 50
      Height = 13
      Caption = 'Proveedor'
    end
    object Label8: TLabel
      Left = 13
      Top = 249
      Width = 43
      Height = 13
      Caption = 'Cantidad'
    end
    object Label9: TLabel
      Left = 316
      Top = 25
      Width = 35
      Height = 13
      Caption = 'Insumo'
    end
    object Label10: TLabel
      Left = 316
      Top = 57
      Width = 52
      Height = 13
      Caption = 'Referencia'
    end
    object Label11: TLabel
      Left = 316
      Top = 89
      Width = 71
      Height = 13
      Caption = 'Observaciones'
    end
    object Label12: TLabel
      Left = 316
      Top = 121
      Width = 59
      Height = 13
      Caption = 'Instrumento'
    end
    object dbFecha: TcxDBDateEdit
      Left = 105
      Top = 22
      DataBinding.DataField = 'fecha'
      DataBinding.DataSource = dsCalidad_rir
      Style.Color = clWindow
      TabOrder = 0
      OnKeyUp = GlobalKeyUp
      Width = 121
    end
    object dbOrdenCompra: TcxDBTextEdit
      Left = 105
      Top = 56
      DataBinding.DataField = 'numero_oc'
      DataBinding.DataSource = dsCalidad_rir
      TabOrder = 3
      OnKeyUp = GlobalKeyUp
      Width = 184
    end
    object dbAvEmbarque: TcxDBTextEdit
      Left = 105
      Top = 86
      DataBinding.DataField = 'numero_aviso_emb'
      DataBinding.DataSource = dsCalidad_rir
      TabOrder = 4
      OnKeyUp = GlobalKeyUp
      Width = 184
    end
    object dbReporte: TcxDBTextEdit
      Left = 105
      Top = 118
      DataBinding.DataField = 'sNumeroReporte'
      DataBinding.DataSource = dsCalidad_rir
      TabOrder = 6
      OnKeyUp = GlobalKeyUp
      Width = 184
    end
    object cbbPlataforma: TcxDBExtLookupComboBox
      Left = 105
      Top = 147
      DataBinding.DataField = 'sIdPlataforma'
      DataBinding.DataSource = dsCalidad_rir
      Properties.DropDownWidth = 400
      Properties.View = cxGridPlataformas
      Properties.KeyFieldNames = 'sIdPlataforma'
      Properties.ListFieldItem = cxGridPlataformassDescripcion
      TabOrder = 8
      OnKeyUp = GlobalKeyUp
      Width = 184
    end
    object cbbEmbarcacion: TcxDBExtLookupComboBox
      Left = 105
      Top = 180
      DataBinding.DataField = 'sIdEmbarcacion'
      DataBinding.DataSource = dsCalidad_rir
      Properties.DropDownWidth = 400
      Properties.View = cxGridEmbarcaciones
      Properties.KeyFieldNames = 'sIdEmbarcacion'
      Properties.ListFieldItem = cxGridEmbarcacionessDescripcion
      TabOrder = 9
      OnKeyUp = GlobalKeyUp
      Width = 184
    end
    object cbbProveedor: TcxDBExtLookupComboBox
      Left = 105
      Top = 211
      DataBinding.DataField = 'sIdProveedor'
      DataBinding.DataSource = dsCalidad_rir
      Properties.DropDownWidth = 400
      Properties.View = cxGridProveedores
      Properties.KeyFieldNames = 'sIdProveedor'
      Properties.ListFieldItem = cxGridProveedoressRazon
      TabOrder = 10
      OnKeyUp = GlobalKeyUp
      Width = 184
    end
    object dbCantidad: TcxDBCalcEdit
      Left = 105
      Top = 246
      DataBinding.DataField = 'dCantidad'
      DataBinding.DataSource = dsCalidad_rir
      TabOrder = 11
      OnKeyUp = GlobalKeyUp
      Width = 121
    end
    object cbbInsumo: TcxDBExtLookupComboBox
      Left = 400
      Top = 22
      DataBinding.DataField = 'sIdInsumo'
      DataBinding.DataSource = dsCalidad_rir
      Properties.DropDownWidth = 400
      Properties.View = cxGridInsumos
      Properties.KeyFieldNames = 'sIdInsumo'
      Properties.ListFieldItem = cxGridInsumosmDescripcion
      TabOrder = 1
      OnKeyUp = GlobalKeyUp
      Width = 300
    end
    object dbReferencia: TcxDBTextEdit
      Left = 400
      Top = 54
      DataBinding.DataField = 'sReferencia'
      DataBinding.DataSource = dsCalidad_rir
      TabOrder = 2
      OnKeyUp = GlobalKeyUp
      Width = 300
    end
    object dbObservaciones: TcxDBTextEdit
      Left = 400
      Top = 86
      DataBinding.DataField = 'sObservaciones'
      DataBinding.DataSource = dsCalidad_rir
      TabOrder = 5
      OnKeyUp = GlobalKeyUp
      Width = 300
    end
    object dbInstrumento: TcxDBTextEdit
      Left = 400
      Top = 118
      DataBinding.DataField = 'sInstrumento'
      DataBinding.DataSource = dsCalidad_rir
      TabOrder = 7
      OnKeyUp = GlobalKeyUp
      Width = 300
    end
  end
  object grpConceptos: TcxGroupBox
    Left = 53
    Top = 64
    TabOrder = 2
    Visible = False
    DesignSize = (
      707
      341)
    Height = 341
    Width = 707
    object Label13: TLabel
      Left = 93
      Top = 209
      Width = 54
      Height = 13
      Caption = 'Descripcion'
    end
    object Label14: TLabel
      Left = 93
      Top = 236
      Width = 48
      Height = 13
      Caption = 'Resultado'
    end
    object lblPaquete: TLabel
      Left = 199
      Top = 271
      Width = 146
      Height = 13
      Caption = 'Seleccionar Concepto a Anidar'
      Visible = False
    end
    object Panel2: TPanel
      Left = 8
      Top = 11
      Width = 79
      Height = 124
      BevelOuter = bvNone
      TabOrder = 0
      object btnInsertarConcepto: TcxButton
        Left = 0
        Top = 0
        Width = 79
        Height = 25
        Align = alTop
        Caption = 'Insertar'
        TabOrder = 0
        TabStop = False
        OnClick = btnInsertarConceptoClick
        OptionsImage.ImageIndex = 0
        OptionsImage.Images = cxImgAcciones
      end
      object btnEditarConcepto: TcxButton
        Left = 0
        Top = 25
        Width = 79
        Height = 25
        Align = alTop
        Caption = 'Editar'
        TabOrder = 1
        TabStop = False
        OnClick = btnEditarConceptoClick
        OptionsImage.ImageIndex = 1
        OptionsImage.Images = cxImgAcciones
      end
      object btnEliminarConcepto: TcxButton
        Left = 0
        Top = 50
        Width = 79
        Height = 25
        Align = alTop
        Caption = 'Eliminar'
        TabOrder = 2
        TabStop = False
        OnClick = btnEliminarConceptoClick
        OptionsImage.ImageIndex = 2
        OptionsImage.Images = cxImgAcciones
      end
      object btnGuardarConcepto: TcxButton
        Left = 0
        Top = 75
        Width = 79
        Height = 25
        Align = alTop
        Caption = 'Guardar'
        Enabled = False
        TabOrder = 3
        TabStop = False
        OnClick = btnGuardarConceptoClick
        OptionsImage.ImageIndex = 3
        OptionsImage.Images = cxImgAcciones
      end
      object btnCancelarConcepto: TcxButton
        Left = 0
        Top = 100
        Width = 79
        Height = 25
        Align = alTop
        Caption = 'Cancelar'
        Enabled = False
        TabOrder = 4
        TabStop = False
        OnClick = btnCancelarConceptoClick
        OptionsImage.ImageIndex = 4
        OptionsImage.Images = cxImgAcciones
      end
    end
    object cxGrid1: TcxGrid
      Left = 93
      Top = 11
      Width = 604
      Height = 183
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 3
      object cxGrid1DBTableView1: TcxGridDBTableView
        Navigator.Buttons.CustomButtons = <>
        DataController.DataSource = dsCalidad_Conceptos
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsData.CancelOnExit = False
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsView.ColumnAutoWidth = True
        OptionsView.GroupByBox = False
        object cxGrid1DBTableView1sDescripcion: TcxGridDBColumn
          Caption = 'Descripcion'
          DataBinding.FieldName = 'sDescripcion'
          Width = 509
        end
        object cxGrid1DBTableView1sresultado: TcxGridDBColumn
          Caption = 'Resultado'
          DataBinding.FieldName = 'sresultado'
          Width = 93
        end
      end
      object cxGrid1Level1: TcxGridLevel
        GridView = cxGrid1DBTableView1
      end
    end
    object dbDescripcionConcepto: TcxDBTextEdit
      Left = 153
      Top = 206
      DataBinding.DataField = 'sDescripcion'
      DataBinding.DataSource = dsCalidad_Conceptos
      TabOrder = 1
      OnKeyUp = GlobalKeyUp
      Width = 512
    end
    object chkSubConcepto: TcxCheckBox
      Left = 93
      Top = 268
      Caption = 'Sub Concepto'
      Enabled = False
      TabOrder = 4
      OnClick = chkSubConceptoClick
      Width = 100
    end
    object cbbPadre: TcxDBLookupComboBox
      Left = 199
      Top = 290
      DataBinding.DataField = 'Padre'
      DataBinding.DataSource = dsCalidad_Conceptos
      Properties.KeyFieldNames = 'idregistro'
      Properties.ListColumns = <
        item
          FieldName = 'sdescripcion'
        end>
      Properties.ListSource = dsSrc_Calidad_Conceptos
      TabOrder = 5
      Visible = False
      Width = 337
    end
    object dbResultado: TcxDBTextEdit
      Left = 153
      Top = 233
      DataBinding.DataField = 'sresultado'
      DataBinding.DataSource = dsCalidad_Conceptos
      TabOrder = 2
      OnKeyUp = GlobalKeyUp
      Width = 98
    end
  end
  object cxImgAcciones: TcxImageList
    FormatVersion = 1
    DesignInfo = 13631504
    ImageInfo = <
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000020000000A170D0738542D1894814626D193502AEA924F2AE87F45
          25D0522C17931209053000000009000000010000000000000000000000000000
          00030201011159311B97A96239FAC58957FFD6A36DFFDDAF75FFDDAF74FFD6A4
          6BFFC58956FFA46137F53C2112730000000F0000000300000000000000020201
          0110744226B9BC7C4DFFDDAE77FFDEB076FFE2B782FFE3BB87FFE3BC86FFE1B7
          82FFDEAF74FFDBAB72FFBD7E4EFF6F3E24B50000001000000002000000085C36
          2095BE8053FFE0B37CFFDFB076FFDEB177FFB78254FFAA7144FFAB7245FFBC88
          59FFDFB279FFDFB277FFDEB077FFC08253FF55321D920000000A190F0932B070
          47FADFB27DFFDFB27AFFE0B37BFFE0B57DFFA56B3FFFF5EFEAFFF8F3EEFFAB72
          45FFE2B67EFFE0B47CFFE0B47BFFDEB079FFB3734AFB130B072F613C2795CD9B
          6FFFE2B780FFE5BD89FFE7C291FFE8C393FFA56B3FFFF1E6DEFFF9F5F1FFAA71
          44FFE8C494FFE8C393FFE5BF8CFFE1B77FFFD09C6EFF5434218B935E3DD2DCB3
          83FFE3B781FFBA8659FFA97043FFAB7245FFAC7346FFF5EDE6FFFAF6F3FFAD75
          47FFB0784AFFB17A4BFFC29162FFE4B983FFDEB17EFF8E5B3BD0B0744CF2E3BF
          8FFFE4BB84FFA56B3FFFF3EBE6FFFAF6F3FFF6EFE8FFF7F0EAFFFBF7F5FFFAF7
          F4FFFAF7F3FFFAF6F2FFAB7245FFE5BD87FFE5BE8BFFAB714CEEAE764FECE9C9
          A0FFE5BE89FFA56B3FFFE0D2CAFFE1D3CCFFE3D5CFFFF2EAE4FFF8F3EFFFEADF
          D9FFE6DAD4FFE9DED9FFAA7144FFE7C08CFFEACA9DFFAE764FEE9A6A49D0E9CD
          ACFFEAC796FFB78456FFA56B3FFFA56B3FFFA56B3FFFF1EAE5FFFAF6F3FFA56B
          3FFFA56B3FFFA56B3FFFB78457FFEACA99FFEBD1ADFF996A49D46E4E3697DDBB
          9DFFEED3A9FFEECFA2FFEED2A5FFF0D6A9FFA56B3FFFF0EAE7FFFDFCFBFFA56B
          3FFFF1D6AAFFF0D5A8FFEED2A5FFEFD4A7FFE0C2A2FF6246318F1C140E2BC794
          6CFCF5E8CCFFEFD6ABFFF1D8AEFFF2DAB0FFA56B3FFFDECFC9FFDFD1CBFFA56B
          3FFFF3DCB2FFF1DBB0FFF1D8ADFFF7EACDFFC69470FA1A120D2E000000036F52
          3C92D7B08CFFF8EFD3FFF3E0B9FFF3DFB7FFB98A5FFFA56B3FFFA56B3FFFBA8A
          5FFFF4E1B9FFF4E2BDFFFAF1D5FFD9B390FF664B368C00000006000000010202
          0107906C4EB8D9B38FFFF7EDD3FFF8EED0FFF7EBC9FFF6E8C4FFF6E8C5FFF7EC
          CAFFF8EED0FFF4E8CDFFD7AF8BFF88664AB30202010B00000001000000000000
          00010202010770543F8FCFA078FCE2C4A2FFEBD7B8FFF4E9CDFFF4EACEFFECD8
          B9FFE3C5A3FFC59973F24C392A67000000060000000100000000000000000000
          000000000001000000022019122C6C543E89A47E5FCCC59770F1C19570EEA47E
          60CD6C543F8B16110D2200000003000000010000000000000000}
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
      end
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
          0000000000020000000C05031A46110852AB190C76E31D0E89FF1C0E89FF190C
          76E4120852AD06031B4D0000000E000000030000000000000000000000000000
          000301010519130A55A9211593FF2225AEFF2430C2FF2535CBFF2535CCFF2430
          C3FF2225AFFF211594FF140B58B20101051E0000000400000000000000020101
          03151C1270CD2522A6FF2D3DCCFF394BD3FF3445D1FF2939CDFF2839CDFF3344
          D0FF394AD4FF2D3CCDFF2523A8FF1C1270D20101051D00000003000000091912
          5BA72A27AAFF2F41D0FF3541C7FF2726ABFF3137BCFF384AD3FF384BD3FF3137
          BCFF2726ABFF3540C7FF2E40D0FF2927ACFF1A115EB10000000D08061C3D3129
          A2FD2C3CCCFF3842C6FF5F5DBDFFEDEDF8FF8B89CEFF3337B9FF3437B9FF8B89
          CEFFEDEDF8FF5F5DBDFF3741C6FF2B3ACDFF3028A4FF0907204A1E185F9F373B
          BCFF3042D0FF2621A5FFECE7ECFFF5EBE4FFF8F2EEFF9491D1FF9491D1FFF8F1
          EDFFF3E9E2FFECE6EBFF2621A5FF2E3FCFFF343ABEFF201A66B0312A92E03542
          CBFF3446D1FF2C2FB5FF8070ADFFEBDBD3FFF4EAE4FFF7F2EDFFF8F1EDFFF4E9
          E2FFEADAD1FF7F6FACFF2B2EB5FF3144D0FF3040CBFF312A95E53E37AEFA3648
          D0FF374AD3FF3A4ED5FF3234B4FF8A7FB9FFF6ECE7FFF5ECE6FFF4EBE5FFF6EB
          E5FF897DB8FF3233B4FF384BD3FF3547D2FF3446D1FF3E37AEFA453FB4FA4557
          D7FF3B50D5FF4C5FDAFF4343B7FF9189C7FFF7EFE9FFF6EEE9FFF6EFE8FFF7ED
          E8FF9087C5FF4242B7FF495DD8FF394CD4FF3F52D4FF443FB3FA403DA1DC5967
          DAFF5B6EDDFF4F4DBAFF8F89CAFFFBF6F4FFF7F1ECFFEDE1D9FFEDE0D9FFF7F0
          EAFFFAF5F2FF8F89CAFF4E4DB9FF576ADCFF5765D9FF403EA4E12E2D70987C85
          DDFF8798E8FF291D9BFFE5DADEFFF6EEEBFFEDDFDAFF816EA9FF816EA9FFEDDF
          D8FFF4ECE7FFE5D9DCFF291D9BFF8494E7FF7A81DDFF33317BAC111125356768
          D0FC9EACEDFF686FCEFF5646A1FFCCB6BCFF7A68A8FF4C4AB6FF4D4BB7FF7A68
          A8FFCBB5BCFF5646A1FF666DCCFF9BAAEEFF696CD0FD1212273F000000043B3B
          79977D84DFFFA5B6F1FF6D74D0FF2D219BFF5151B9FF8EA2ECFF8EA1ECFF5252
          BBFF2D219BFF6B72D0FFA2B3F0FF8086E0FF404183A700000008000000010303
          050C4E509DBC8087E2FFAEBDF3FFA3B6F1FF9DAFF0FF95A9EEFF95A8EEFF9BAD
          EFFFA2B3F0FFACBCF3FF838AE3FF4F52A0C10303051100000002000000000000
          000100000005323464797378D9F8929CEAFFA1AEEFFFB0BFF3FFB0BFF4FFA2AE
          EFFF939DE9FF7479DAF83234647D000000080000000200000000000000000000
          000000000000000000031213232D40437D935D61B5D07378DFFC7378DFFC5D61
          B5D040437D951212223000000004000000010000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          000000000000000000030000000C000000120000001400000014000000150000
          0015000000140000000D00000003000000000000000000000000000000000000
          00000000000000000009896256C2BD8A78FFBE8A78FFBD8A78FFBD8978FFBD89
          78FFBE8978FF876356C30000000B000000000000000000000000000000030000
          000E0000001500000021C08D7CFFF6EEE9FFF5EDE9FFF5EDE9FFF5ECE8FFF4EC
          E8FFF5ECE7FFBF8D7BFF00000026000000180000000F000000040000000C7B50
          42C5A76E5BFF9F6755FFC2917FFFF7F0ECFFE2B47DFFE2B37AFFE1B077FFE0AE
          72FFF6EEEAFFC2907FFF845545FF895847FF613E32C70000000E00000011BB7E
          6BFFECD9CCFFE3CEBEFFC59483FFF9F2F0FFE4B984FFE3B781FFE3B47CFFE0B1
          76FFF7F0EDFFC59483FFE0CBBCFFEBD8CBFFB67763FF0000001400000010BE85
          71FFF1E5DAFFECDBD0FF7A4835FF7A4835FF7A4835FF7A4835FF7A4835FF7A48
          35FF7A4835FF7A4835FFEBDBCFFFF1E2D8FFB97C69FF000000130000000EC28B
          78FFF5EEE7FFF2E7DDFFF2E7DEFFF3E7DEFFF2E5DEFFF3E5DEFFF2E7DDFFF2E7
          DDFFF2E7DEFFF2E7DDFFF2E5DEFFF5EDE6FFBC826EFF000000120000000CC793
          7FFFFAF4F1FFCDBEB8FF6F5448FF614337FF614035FF5F3F34FF5E3F33FF5D3D
          32FF5D3D34FF6A4C44FFCABCB6FFF9F5F1FFC18875FF000000100000000ACC99
          86FFFDFAFAFF7D6054FF745043FF744F43FF744E43FF734E43FF734E42FF724D
          42FF724C41FF724C40FF73584DFFFDFAFAFFC58F7CFF0000000E00000008CF9F
          8DFFFFFFFFFF7A5A4CFF8E695AFFF9F4F1FFF0E6E0FFF0E5DFFFEFE5DEFFEFE5
          DEFFF6EFEBFF866253FF704F43FFFFFFFFFFCA9683FF0000000B00000005BE95
          84E9F5ECE8FF866656FF977262FFFAF6F4FFF2E8E3FFF1E8E1FFF1E7E2FFF1E7
          E1FFF8F2EEFF8E6A5BFF7A5B4CFFF5EAE6FFBA8E7DEA00000008000000023429
          2545A78375CC947262FFA07B6AFFFCF9F8FFF3EBE6FFF4EAE5FFF2EAE5FFF3EA
          E3FFF9F5F3FF977263FF876658FFA68072CE3428234800000003000000000000
          0001000000030000000AC89B89FFFDFBFAFFF5EDE8FFF4EDE8FFF5EDE7FFF5EC
          E7FFFBF7F6FFC59685FF00000011000000040000000200000000000000000000
          00000000000000000005CA9E8DFFFEFCFCFFF7F0ECFFF6EFEBFFF7EFEBFFF5EF
          EAFFFCFAF8FFC89A89FF00000009000000000000000000000000000000000000
          00000000000000000003CDA291FFFEFEFDFFFEFDFDFFFEFDFCFFFEFCFCFFFEFC
          FBFFFDFBFAFFCB9F8DFF00000007000000000000000000000000000000000000
          000000000000000000019A796DBFCFA493FFCEA493FFCEA493FFCEA492FFCDA3
          91FFCDA391FF98786BC100000004000000000000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000001000000060000
          000A0000000C0000000A00000006000000020000000000000000000000000000
          000000000000000000000000000000000000000000020201010D4E32217F8A58
          3AD3AB6C47FF885638D44C2F1F800201010E0000000200000000000000000000
          0000000000000000000000000000000000010403021086593CC7BF885CFFD4A3
          6EFFDEB076FFD2A16BFFBC8257FF825236C70201010E00000001000000000000
          0000000000000000000000000000000000045B3E2C88C59167FFE1B984FFE1B8
          82FFE1B781FFE0B67FFFDFB57EFFC08A60FF52352484000000050000000B0000
          0011000000130000001300000013000000199B6E4FE9CDAA7FFFD3B385FFD4B2
          84FFD9B686FFE0BC89FFE3BC8AFFDAB07FFF9B6846E200000007735246C3A072
          61FF9F7261FF9F7060FF9E7060FF9E705FFF9E705FFF9D6F5EFF9D6E5DFF9D6E
          5DFFA67D66FFDCBE91FFE7C696FFE7C79AFFB27B55FA00000008A37666FFFFFF
          FFFFF6EDE6FFF6ECE6FFF7EDE6FFF7EDE6FFF7ECE6FFF7ECE5FFF7ECE5FFF7EC
          E5FF9F7060FFDBC199FFEACEA3FFE8CEACFFA47352E100000005A77B6BFFFFFF
          FFFFF7EEE9FFF7EFE9FFF7EEE8FFF7EEE8FFF7EDE7FFF7EDE7FFF7EDE8FFF6ED
          E7FFA17463FFDFCBABFFF3E4C8FFD3AB89FF5A412F7E00000003AB7F70FFFFFF
          FFFFF8F0EBFFC3907BFFC38F7AFFC28F79FFC18D78FFC18D77FFF7EFE9FFF7EE
          E9FFA57868FFDECCB6FFD5AD8CFF956F52C30202010700000001AE8575FFFFFF
          FFFFF9F2EDFFF9F2EDFFF9F2EDFFF9F2EDFFF8F2EDFFF8F1ECFFF8F1ECFFF8F0
          EBFFA87D6DFF9F795BD66048367C020201060000000100000000B2897AFFFFFF
          FFFFFAF4F0FFC89883FFC89782FFC79781FFC69680FFC6947FFFF9F3EEFFF9F3
          EEFFAC8172FF0000000F00000001000000000000000000000000B78E80FFFFFF
          FFFFFAF6F3FFFAF6F2FFFAF5F2FFFAF5F1FFF9F5F1FFF9F5F1FFF9F5F1FFFAF5
          F0FFB08777FF0000000C00000000000000000000000000000000BA9384FFFFFF
          FFFFFBF8F5FFCCA08BFFCC9F8AFFCB9E89FFCB9E89FFCB9D88FFFAF6F4FFFAF7
          F3FFB48C7DFF0000000A00000000000000000000000000000000BD9789FFFFFF
          FFFFFCFAF7FFFCF9F6FFFBF9F7FFFBF9F7FFFBF8F6FFFBF8F6FFFBF8F6FFFBF8
          F6FFB89081FF0000000900000000000000000000000000000000C19B8DFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFBB9587FF000000070000000000000000000000000000000091766CC0C29E
          90FFC29E8FFFC29D8FFFC19D8FFFC19C8EFFC19B8DFFC09B8DFFBF9A8CFFC09A
          8CFF8E7267C10000000400000000000000000000000000000000}
      end>
  end
  object zCalidad_Rir: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select * from calidad_rir where sContrato = :contrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end>
    Left = 16
    Top = 240
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end>
    object zCalidad_Riridregistro: TIntegerField
      FieldName = 'idregistro'
      Required = True
    end
    object zCalidad_Rirfecha: TDateField
      FieldName = 'fecha'
    end
    object zCalidad_Rirnumero_oc: TStringField
      FieldName = 'numero_oc'
      Size = 100
    end
    object zCalidad_Rirnumero_aviso_emb: TStringField
      FieldName = 'numero_aviso_emb'
      Size = 100
    end
    object zCalidad_RirsContrato: TStringField
      FieldName = 'sContrato'
      Size = 15
    end
    object zCalidad_RirsNumeroReporte: TStringField
      FieldName = 'sNumeroReporte'
      Size = 100
    end
    object zCalidad_RirsIdPlataforma: TStringField
      FieldName = 'sIdPlataforma'
      Size = 50
    end
    object zCalidad_RirsIdEmbarcacion: TStringField
      FieldName = 'sIdEmbarcacion'
      Size = 10
    end
    object zCalidad_RirsIdProveedor: TStringField
      FieldName = 'sIdProveedor'
      Size = 5
    end
    object zCalidad_RirdCantidad: TFloatField
      FieldName = 'dCantidad'
    end
    object zCalidad_RirsIdInsumo: TStringField
      FieldName = 'sIdInsumo'
      Size = 25
    end
    object zCalidad_RirsReferencia: TStringField
      FieldName = 'sReferencia'
      Size = 200
    end
    object zCalidad_RirsObservaciones: TStringField
      FieldName = 'sObservaciones'
      Size = 300
    end
    object zCalidad_RirsInstrumento: TStringField
      FieldName = 'sInstrumento'
      Size = 250
    end
  end
  object dsCalidad_rir: TDataSource
    DataSet = zCalidad_Rir
    Left = 48
    Top = 240
  end
  object zPlataformas: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select sIdPlataforma, sDescripcion from plataformas where lStatu' +
        's = "Activa"')
    Params = <>
    Left = 16
    Top = 272
    object zPlataformassIdPlataforma: TStringField
      FieldName = 'sIdPlataforma'
      Required = True
      Size = 50
    end
    object zPlataformassDescripcion: TStringField
      FieldName = 'sDescripcion'
      Required = True
      Size = 50
    end
  end
  object zEmbarcaciones: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select sIdEmbarcacion, sDescripcion from embarcaciones where sCo' +
        'ntrato = :contrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end>
    Left = 16
    Top = 304
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end>
    object zEmbarcacionessIdEmbarcacion: TStringField
      FieldName = 'sIdEmbarcacion'
      Required = True
      Size = 10
    end
    object zEmbarcacionessDescripcion: TStringField
      FieldName = 'sDescripcion'
      Required = True
      Size = 50
    end
  end
  object zProveedores: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select sIdProveedor, sRazon from proveedores;')
    Params = <>
    Left = 16
    Top = 336
    object zProveedoressIdProveedor: TStringField
      FieldName = 'sIdProveedor'
      Required = True
      Size = 5
    end
    object zProveedoressRazon: TStringField
      FieldName = 'sRazon'
      Size = 50
    end
  end
  object zInsumos: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select sIdInsumo, mDescripcion from insumos where sContrato = :c' +
        'ontrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end>
    Left = 16
    Top = 368
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end>
    object zInsumossIdInsumo: TStringField
      FieldName = 'sIdInsumo'
      Required = True
      Size = 25
    end
    object zInsumosmDescripcion: TMemoField
      FieldName = 'mDescripcion'
      Required = True
      BlobType = ftMemo
    end
  end
  object dsPlataformas: TDataSource
    DataSet = zPlataformas
    Left = 48
    Top = 272
  end
  object dsEmbarcaciones: TDataSource
    DataSet = zEmbarcaciones
    Left = 48
    Top = 304
  end
  object dsProveedores: TDataSource
    DataSet = zProveedores
    Left = 48
    Top = 336
  end
  object dsInsumos: TDataSource
    DataSet = zInsumos
    Left = 48
    Top = 368
  end
  object dxGrids: TcxGridViewRepository
    Left = 48
    Top = 208
    object cxGridPlataformas: TcxGridDBTableView
      Navigator.Buttons.CustomButtons = <>
      DataController.DataSource = dsPlataformas
      DataController.KeyFieldNames = 'sIdPlataforma'
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      FilterRow.Visible = True
      OptionsData.CancelOnExit = False
      OptionsData.Deleting = False
      OptionsData.DeletingConfirmation = False
      OptionsData.Editing = False
      OptionsData.Inserting = False
      OptionsView.ColumnAutoWidth = True
      OptionsView.GroupByBox = False
      object cxGridPlataformassIdPlataforma: TcxGridDBColumn
        Caption = 'Id Proyecto'
        DataBinding.FieldName = 'sIdPlataforma'
        HeaderAlignmentHorz = taCenter
        Width = 70
      end
      object cxGridPlataformassDescripcion: TcxGridDBColumn
        Caption = 'Descripcion'
        DataBinding.FieldName = 'sDescripcion'
        HeaderAlignmentHorz = taCenter
      end
    end
    object cxGridEmbarcaciones: TcxGridDBTableView
      Navigator.Buttons.CustomButtons = <>
      DataController.DataSource = dsEmbarcaciones
      DataController.KeyFieldNames = 'sIdEmbarcacion'
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      FilterRow.Visible = True
      OptionsData.CancelOnExit = False
      OptionsData.Deleting = False
      OptionsData.DeletingConfirmation = False
      OptionsData.Editing = False
      OptionsData.Inserting = False
      OptionsView.ColumnAutoWidth = True
      OptionsView.GroupByBox = False
      object cxGridEmbarcacionessIdEmbarcacion: TcxGridDBColumn
        Caption = 'Id Embarcaci'#243'n'
        DataBinding.FieldName = 'sIdEmbarcacion'
        HeaderAlignmentHorz = taCenter
        Width = 70
      end
      object cxGridEmbarcacionessDescripcion: TcxGridDBColumn
        Caption = 'Descripcion'
        DataBinding.FieldName = 'sDescripcion'
        HeaderAlignmentHorz = taCenter
      end
    end
    object cxGridProveedores: TcxGridDBTableView
      Navigator.Buttons.CustomButtons = <>
      DataController.DataSource = dsProveedores
      DataController.KeyFieldNames = 'sIdProveedor'
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      FilterRow.Visible = True
      OptionsData.CancelOnExit = False
      OptionsData.Deleting = False
      OptionsData.DeletingConfirmation = False
      OptionsData.Editing = False
      OptionsData.Inserting = False
      OptionsView.ColumnAutoWidth = True
      OptionsView.GroupByBox = False
      object cxGridProveedoressIdProveedor: TcxGridDBColumn
        Caption = 'Id Proveedor'
        DataBinding.FieldName = 'sIdProveedor'
        HeaderAlignmentHorz = taCenter
        Width = 80
      end
      object cxGridProveedoressRazon: TcxGridDBColumn
        Caption = 'Descripci'#243'n'
        DataBinding.FieldName = 'sRazon'
        HeaderAlignmentHorz = taCenter
      end
    end
    object cxGridInsumos: TcxGridDBTableView
      Navigator.Buttons.CustomButtons = <>
      DataController.DataSource = dsInsumos
      DataController.KeyFieldNames = 'sIdInsumo'
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      FilterRow.Visible = True
      OptionsData.CancelOnExit = False
      OptionsData.Deleting = False
      OptionsData.DeletingConfirmation = False
      OptionsData.Editing = False
      OptionsData.Inserting = False
      OptionsView.ColumnAutoWidth = True
      OptionsView.GroupByBox = False
      object cxGridInsumossIdInsumo: TcxGridDBColumn
        Caption = 'Insumo'
        DataBinding.FieldName = 'sIdInsumo'
        Width = 50
      end
      object cxGridInsumosmDescripcion: TcxGridDBColumn
        Caption = 'Descripcion'
        DataBinding.FieldName = 'mDescripcion'
      end
    end
  end
  object zCalidad_Conceptos: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select * from calidad_conceptos_rir where id_registro_rir = :id_' +
        'rir order by padre')
    Params = <
      item
        DataType = ftUnknown
        Name = 'id_rir'
        ParamType = ptUnknown
      end>
    Left = 16
    Top = 416
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'id_rir'
        ParamType = ptUnknown
      end>
  end
  object zSrcCalidad_Conceptos: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select idregistro, sdescripcion from calidad_conceptos_rir where' +
        ' id_registro_rir = :id_rir and idregistro = padre')
    Params = <
      item
        DataType = ftUnknown
        Name = 'id_rir'
        ParamType = ptUnknown
      end>
    Left = 16
    Top = 448
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'id_rir'
        ParamType = ptUnknown
      end>
  end
  object dsCalidad_Conceptos: TDataSource
    DataSet = zCalidad_Conceptos
    Left = 48
    Top = 416
  end
  object dsSrc_Calidad_Conceptos: TDataSource
    DataSet = zSrcCalidad_Conceptos
    Left = 48
    Top = 448
  end
  object frxTitulo: TfrxDBDataset
    UserName = 'frxTitulo'
    CloseDataSource = False
    FieldAliases.Strings = (
      'imagen=imagen'
      'cia=cia'
      'direccion=direccion')
    DataSet = zqryTitulo
    BCDToCurrency = False
    Left = 576
    Top = 432
  end
  object frxCuerpo: TfrxDBDataset
    UserName = 'frxCuerpo'
    CloseDataSource = False
    FieldAliases.Strings = (
      'NoReporte=NoReporte'
      'contrato=contrato'
      'fecha=fecha'
      'proyecto=proyecto'
      'lugarInspeccion=lugarInspeccion'
      'proveedor=proveedor'
      'noOrdenCompra=noOrdenCompra'
      'noAvisoEmbarque=noAvisoEmbarque'
      'descripcionMaterial=descripcionMaterial'
      'cantidad=cantidad'
      'unidad=unidad')
    DataSet = zqryCuerpo
    BCDToCurrency = False
    Left = 624
    Top = 432
  end
  object frxDetalle: TfrxDBDataset
    UserName = 'frxDetalle'
    CloseDataSource = False
    FieldAliases.Strings = (
      'PadreHijo=PadreHijo'
      'valor=valor'
      'nomenclatura=nomenclatura'
      'referencia=referencia'
      'observaciones=observaciones'
      'instrumento=instrumento')
    DataSet = zqryDetalle
    BCDToCurrency = False
    Left = 672
    Top = 432
  end
  object frxFirmantes: TfrxDBDataset
    UserName = 'frxFirmantes'
    CloseDataSource = False
    FieldAliases.Strings = (
      'nombreFirmante=nombreFirmante'
      'puestoFirmante=puestoFirmante')
    DataSet = zqryFirmante
    BCDToCurrency = False
    Left = 720
    Top = 432
  end
  object zqryTitulo: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select'
      '   conf.bImagen as imagen,'
      '   conf.sNombre as cia,'
      
        '   concat(conf.sDireccion1, '#39' '#39', conf.sDireccion2, '#39' '#39', conf.sDi' +
        'reccion3) as direccion'
      ''
      'from calidad_rir cr '
      '   inner join configuracion conf'
      '      on (conf.sContrato = cr.sContrato)'
      ''
      'where (cr.idregistro = :idregistro)')
    Params = <
      item
        DataType = ftUnknown
        Name = 'idregistro'
        ParamType = ptUnknown
      end>
    Left = 576
    Top = 472
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'idregistro'
        ParamType = ptUnknown
      end>
  end
  object zqryCuerpo: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select'
      '   cr.sNumeroReporte as NoReporte,'
      '   cont.sContrato as contrato,'
      '   cr.fecha as fecha,'
      '   plat.sDescripcion as proyecto,'
      '   emb.sDescripcion as lugarInspeccion,'
      '   prov.sRazon as proveedor,'
      '   cr.numero_oc as noOrdenCompra,'
      '   cr.numero_aviso_emb as noAvisoEmbarque,'
      '   ins.mDescripcion as descripcionMaterial,'
      '   cr.dCantidad as cantidad,'
      '   ins.sMedida as unidad'
      '   '
      'from calidad_rir cr '
      ''
      'inner join configuracion conf'
      'on (conf.sContrato = cr.sContrato)'
      ''
      'inner join contratos cont '
      'on (cont.sContrato = conf.sContrato)'
      ''
      'inner join plataformas plat '
      'on (plat.sIdPlataforma = cr.sIdPlataforma)'
      ''
      'inner join embarcaciones emb'
      'on (emb.sIdEmbarcacion = cr.sIdEmbarcacion)'
      ''
      'inner join proveedores prov'
      'on (prov.sIdProveedor = cr.sIdProveedor)'
      '   '
      'inner join insumos ins'
      'on (ins.sIdInsumo = cr.sIdInsumo)'
      ''
      'where (cr.idregistro = :idregistro)')
    Params = <
      item
        DataType = ftUnknown
        Name = 'idregistro'
        ParamType = ptUnknown
      end>
    Left = 624
    Top = 472
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'idregistro'
        ParamType = ptUnknown
      end>
  end
  object zqryDetalle: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select'
      '   ccr.idregistro as idregistro_concepto,'
      '   ccr.padre,'
      '   IF(ccr.padre = ccr.idregistro, '#39'PADRE'#39', '#39'HIJO'#39') as PadreHijo,'
      '   ccr.sResultado as valor,'
      '   '
      '   ccr.sDescripcion as nomenclatura,'
      '   '
      '   cr.sReferencia as referencia,'
      '   cr.sObservaciones as observaciones,'
      '   cr.sInstrumento as instrumento'
      ''
      'from calidad_rir cr'
      ''
      'inner join calidad_conceptos_rir ccr'
      'on (ccr.id_registro_rir = cr.idregistro)'
      ''
      'where (cr.idregistro = :idregistro)'
      'order by ccr.padre, ccr.idregistro')
    Params = <
      item
        DataType = ftUnknown
        Name = 'idregistro'
        ParamType = ptUnknown
      end>
    Left = 672
    Top = 472
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'idregistro'
        ParamType = ptUnknown
      end>
  end
  object zqryFirmante: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select'
      '   fir.sFirmante30 as nombreFirmante,'
      '   fir.sPuesto30 as puestoFirmante'
      ''
      'from calidad_rir cr'
      ''
      'inner join firmas fir'
      'on ((fir.sContrato = :contrato) and (fir.sIdTurno = :turno)'
      '   and fir.dIdFecha = :fecha and fir.sNumeroOrden = :orden)'
      '   '
      'where (cr.idregistro = :idregistro)')
    Params = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'turno'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'orden'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'idregistro'
        ParamType = ptUnknown
      end>
    Left = 720
    Top = 472
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'turno'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'orden'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'idregistro'
        ParamType = ptUnknown
      end>
  end
  object reporteRIR: TfrxReport
    Version = '4.7.109'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Por defecto'
    PrintOptions.PrintOnSheet = 0
    ReportOptions.CreateDate = 41989.986073657400000000
    ReportOptions.LastChange = 41991.350373668980000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin            '
      '    '
      'end.')
    OnGetValue = reporteRIRGetValue
    Left = 504
    Top = 448
    Datasets = <
      item
        DataSet = frxCuerpo
        DataSetName = 'frxCuerpo'
      end
      item
        DataSet = frxDetalle
        DataSetName = 'frxDetalle'
      end
      item
        DataSet = frxFirmantes
        DataSetName = 'frxFirmantes'
      end
      item
        DataSet = frxTitulo
        DataSetName = 'frxTitulo'
      end>
    Variables = <>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      PaperWidth = 215.900000000000000000
      PaperHeight = 279.400000000000000000
      PaperSize = 1
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object ReportTitle1: TfrxReportTitle
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'tahoma'
        Font.Style = []
        Height = 60.944960000000000000
        ParentFont = False
        Top = 18.897650000000000000
        Width = 740.409927000000000000
        object Memo1: TfrxMemoView
          Left = 121.119357140000000000
          Top = 3.689548570000000000
          Width = 510.236550000000000000
          Height = 26.456710000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxTitulo."cia"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo2: TfrxMemoView
          Left = 123.535560000000000000
          Top = 27.897650000000000000
          Width = 506.457020000000000000
          Height = 26.456710000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxTitulo."direccion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo3: TfrxMemoView
          Top = 0.472480000000000000
          Width = 740.787880000000000000
          Height = 60.472480000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          VAlign = vaCenter
        end
        object Picture1: TfrxPictureView
          Left = 8.779530000000000000
          Top = 3.779530000000000000
          Width = 109.606370000000000000
          Height = 56.692950000000000000
          ShowHint = False
          DataField = 'imagen'
          DataSet = frxTitulo
          DataSetName = 'frxTitulo'
          HightQuality = False
        end
      end
      object Header1: TfrxHeader
        Height = 230.551330000000000000
        Top = 139.842610000000000000
        Width = 740.409927000000000000
        object Memo4: TfrxMemoView
          Left = 3.811070000000000000
          Top = 1.440940000000000000
          Width = 729.449290000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -16
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8 = (
            'REGISTRO DE INSPECCI'#195#8220'N RECIBO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo5: TfrxMemoView
          Left = 34.015770000000000000
          Top = 49.118120000000000000
          Width = 64.252010000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'No. Reporte:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo6: TfrxMemoView
          Left = 309.921460000000000000
          Top = 49.118120000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'CONTRATO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo7: TfrxMemoView
          Left = 606.504330000000000000
          Top = 49.118120000000000000
          Width = 22.677180000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'D'#195#173'a')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo8: TfrxMemoView
          Left = 642.520100000000000000
          Top = 49.118120000000000000
          Width = 22.677180000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'Mes')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo9: TfrxMemoView
          Left = 676.535870000000000000
          Top = 49.118120000000000000
          Width = 22.677180000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'A'#195#177'o')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo11: TfrxMemoView
          Left = 109.606370000000000000
          Top = 49.118120000000000000
          Width = 98.267780000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Memo.UTF8 = (
            '[frxCuerpo."NoReporte"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo12: TfrxMemoView
          Left = 98.267780000000000000
          Top = 52.897650000000000000
          Width = 117.165430000000000000
          Height = 11.338590000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo13: TfrxMemoView
          Left = 374.173470000000000000
          Top = 49.118120000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Memo.UTF8 = (
            '[frxCuerpo."contrato"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo14: TfrxMemoView
          Left = 597.165740000000000000
          Top = 34.000000000000000000
          Width = 37.795300000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[Dia]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo15: TfrxMemoView
          Left = 668.976810000000000000
          Top = 34.000000000000000000
          Width = 37.795300000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[Ano]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo16: TfrxMemoView
          Left = 634.961040000000000000
          Top = 34.000000000000000000
          Width = 34.015770000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[Mes]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo17: TfrxMemoView
          Left = 34.015770000000000000
          Top = 79.354360000000000000
          Width = 49.133890000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'Proyecto: ')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo18: TfrxMemoView
          Left = 86.929190000000000000
          Top = 75.574830000000000000
          Width = 257.008040000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."proyecto"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo19: TfrxMemoView
          Left = 355.275820000000000000
          Top = 79.354360000000000000
          Width = 105.826840000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8 = (
            'Lugar de Inspecci'#195#179'n:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo20: TfrxMemoView
          Left = 464.882190000000000000
          Top = 75.574830000000000000
          Width = 245.669450000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."lugarInspeccion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo21: TfrxMemoView
          Left = 34.015770000000000000
          Top = 98.252010000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'Proveedor/Cliente:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo22: TfrxMemoView
          Left = 139.842610000000000000
          Top = 94.472480000000000000
          Width = 570.709030000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."proveedor"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo23: TfrxMemoView
          Left = 34.015770000000000000
          Top = 117.149660000000000000
          Width = 102.047310000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'No. Orden Compra:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo24: TfrxMemoView
          Left = 136.063080000000000000
          Top = 113.370130000000000000
          Width = 207.874150000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."noOrdenCompra"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo25: TfrxMemoView
          Left = 355.275820000000000000
          Top = 117.149660000000000000
          Width = 105.826840000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8 = (
            'No. Aviso Embarque:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo26: TfrxMemoView
          Left = 464.882190000000000000
          Top = 113.370130000000000000
          Width = 245.669450000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."noAvisoEmbarque"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo27: TfrxMemoView
          Left = 34.015770000000000000
          Top = 136.047310000000000000
          Width = 124.724490000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'Descripci'#195#179'n del material:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo28: TfrxMemoView
          Left = 158.740260000000000000
          Top = 132.267780000000000000
          Width = 551.811380000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."descripcionMaterial"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo29: TfrxMemoView
          Left = 34.015770000000000000
          Top = 154.944960000000000000
          Width = 49.133890000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'Cantidad:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo30: TfrxMemoView
          Left = 86.929190000000000000
          Top = 151.165430000000000000
          Width = 257.008040000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."cantidad"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo31: TfrxMemoView
          Left = 355.275820000000000000
          Top = 151.165430000000000000
          Width = 49.133890000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8 = (
            'Unidad:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo32: TfrxMemoView
          Left = 464.882190000000000000
          Top = 151.165430000000000000
          Width = 245.669450000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."unidad"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo69: TfrxMemoView
          Left = 554.118430000000000000
          Top = 37.000000000000000000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Memo.UTF8 = (
            'Fecha:')
          ParentFont = False
        end
        object Memo33: TfrxMemoView
          Left = 196.535560000000000000
          Top = 178.015770000000000000
          Width = 370.393940000000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8 = (
            'Concepto a inspeccionar del Material')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo38: TfrxMemoView
          Left = 544.252320000000000000
          Top = 212.031540000000000000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8 = (
            'Resultado')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object MasterData1: TfrxMasterData
        Height = 0.377860000000000000
        Top = 430.866420000000000000
        Width = 740.409927000000000000
        DataSet = frxDetalle
        DataSetName = 'frxDetalle'
        RowCount = 0
      end
      object GroupHeader1: TfrxGroupHeader
        Height = 15.236240000000000000
        Top = 393.071120000000000000
        Width = 740.409927000000000000
        Condition = '<frxDetalle."nomenclatura">'
        object Memo34: TfrxMemoView
          Left = 33.677180000000000000
          Top = 0.559060000000000000
          Width = 336.378170000000000000
          Height = 15.118120000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Memo.UTF8 = (
            '[PadreHijo]')
          ParentFont = False
        end
        object Memo43: TfrxMemoView
          Left = 509.897960000000000000
          Top = 0.118120000000000000
          Width = 200.315090000000000000
          Height = 15.118120000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxDetalle."valor"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object Header2: TfrxHeader
        Height = 79.370130000000000000
        Top = 453.543600000000000000
        Width = 740.409927000000000000
        object Memo50: TfrxMemoView
          Left = 33.795300000000000000
          Top = 12.850340000000000000
          Width = 136.063080000000000000
          Height = 64.252010000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Memo.UTF8 = (
            'Resultado de la Inspecci'#195#179'n:'
            'A = Aceptado'
            'R = Rechazado'
            'NA = No Aplica'
            'NP = No Presenta')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object Header3: TfrxHeader
        Height = 177.637910000000000000
        Top = 578.268090000000000000
        Width = 740.409927000000000000
        object Memo51: TfrxMemoView
          Left = 302.362400000000000000
          Top = 3.779530000000000000
          Width = 151.181200000000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8 = (
            'No de Rastreabilidad Asignado:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo52: TfrxMemoView
          Left = 332.598640000000000000
          Top = 26.456710000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxCuerpo."NoReporte"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo53: TfrxMemoView
          Left = 302.362400000000000000
          Top = 26.456710000000000000
          Width = 151.181200000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo54: TfrxMemoView
          Left = 34.795300000000000000
          Top = 64.252010000000000000
          Width = 109.606370000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8 = (
            'Referencia al punto C:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo55: TfrxMemoView
          Left = 148.181200000000000000
          Top = 60.472480000000000000
          Width = 566.929500000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          Frame.Typ = [ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            '[frxDetalle."referencia"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo56: TfrxMemoView
          Left = 322.039580000000000000
          Top = 94.488250000000000000
          Width = 113.385900000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8 = (
            'OBSERVACIONES')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo57: TfrxMemoView
          Left = 34.795300000000000000
          Top = 117.165430000000000000
          Width = 680.315400000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Memo.UTF8 = (
            '[frxDetalle."observaciones"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo58: TfrxMemoView
          Left = 34.795300000000000000
          Top = 143.622140000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8 = (
            'X')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo60: TfrxMemoView
          Left = 208.653680000000000000
          Top = 143.622140000000000000
          Width = 506.457020000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Memo.UTF8 = (
            '[frxDetalle."instrumento"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo59: TfrxMemoView
          Left = 76.370130000000000000
          Top = 143.622140000000000000
          Width = 136.063080000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8 = (
            'Instrumento Utilizado No.')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object MasterData2: TfrxMasterData
        Top = 555.590910000000000000
        Width = 740.409927000000000000
        DataSet = frxDetalle
        DataSetName = 'frxDetalle'
        RowCount = 0
      end
      object MasterData3: TfrxMasterData
        Top = 778.583180000000000000
        Width = 740.409927000000000000
        DataSet = frxDetalle
        DataSetName = 'frxDetalle'
        RowCount = 0
      end
      object Memo61: TfrxMemoView
        Left = 34.015770000000000000
        Top = 785.717070000000000000
        Width = 60.472480000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        Memo.UTF8 = (
          'Distribuci'#195#179'n')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo62: TfrxMemoView
        Left = 34.015770000000000000
        Top = 805.055660000000000000
        Width = 56.692950000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        Memo.UTF8 = (
          'Almacen')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo63: TfrxMemoView
        Left = 34.015770000000000000
        Top = 823.953310000000000000
        Width = 192.756030000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        Memo.UTF8 = (
          'Responsable de '#195#161'rea (cuando se aplica)')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo64: TfrxMemoView
        Left = 34.015770000000000000
        Top = 842.850960000000000000
        Width = 192.756030000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        Memo.UTF8 = (
          'Responsable de cliente (S'#195#173' se aplica)')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo65: TfrxMemoView
        Left = 34.015770000000000000
        Top = 861.748610000000000000
        Width = 90.708720000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        Memo.UTF8 = (
          'Control de Calidad')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo66: TfrxMemoView
        Left = 377.953000000000000000
        Top = 793.717070000000000000
        Width = 64.252010000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        Memo.UTF8 = (
          'Inspeccion'#195#179)
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo67: TfrxMemoView
        Left = 291.023810000000000000
        Top = 823.953310000000000000
        Width = 241.889920000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = [fsBold]
        Frame.Typ = [ftBottom]
        HAlign = haCenter
        Memo.UTF8 = (
          '[Firmante]')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo68: TfrxMemoView
        Left = 313.700990000000000000
        Top = 842.850960000000000000
        Width = 211.653680000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        HAlign = haCenter
        Memo.UTF8 = (
          '[PuestoFirmante]')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo10: TfrxMemoView
        Top = 83.149660000000000000
        Width = 740.787880000000000000
        Height = 672.756340000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
        ParentFont = False
        VAlign = vaCenter
      end
    end
    object Page2: TfrxReportPage
      Visible = False
      PaperWidth = 215.900000000000000000
      PaperHeight = 279.400000000000000000
      PaperSize = 1
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
    end
    object Page3: TfrxReportPage
      Visible = False
      PaperWidth = 215.900000000000000000
      PaperHeight = 279.400000000000000000
      PaperSize = 1
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
    end
    object Page4: TfrxReportPage
      Visible = False
      PaperWidth = 215.900000000000000000
      PaperHeight = 279.400000000000000000
      PaperSize = 1
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object Memo35: TfrxMemoView
        Left = 15.118120000000000000
        Top = 30.236240000000000000
        Width = 151.181200000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        HAlign = haCenter
        Memo.UTF8 = (
          'b) Inspeccionar dimensiones.')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo36: TfrxMemoView
        Left = 15.118120000000000000
        Top = 45.354360000000000000
        Width = 275.905690000000000000
        Height = 15.118120000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        HAlign = haCenter
        Memo.UTF8 = (
          'c) Inspeccionar monograma, c'#195#179'digo, lote, serie, marcas.')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo37: TfrxMemoView
        Left = 15.118120000000000000
        Top = 56.692950000000000000
        Width = 207.874150000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        HAlign = haCenter
        Memo.UTF8 = (
          'd) Inspeccionar Certificados/Documentos')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo39: TfrxMemoView
        Left = 154.960730000000000000
        Top = 71.811070000000000000
        Width = 132.283550000000000000
        Height = 15.118120000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        HAlign = haCenter
        Memo.UTF8 = (
          'No. y Fecha del certificado.')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo40: TfrxMemoView
        Left = 154.960730000000000000
        Top = 83.149660000000000000
        Width = 143.622140000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        HAlign = haCenter
        Memo.UTF8 = (
          'Especificaciones del Material.')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo41: TfrxMemoView
        Left = 154.960730000000000000
        Top = 98.267780000000000000
        Width = 120.944960000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        HAlign = haCenter
        Memo.UTF8 = (
          'Propiedades Mec'#195#161'nicas.')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo42: TfrxMemoView
        Left = 154.960730000000000000
        Top = 113.385900000000000000
        Width = 83.149660000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = []
        HAlign = haCenter
        Memo.UTF8 = (
          'An'#195#161'lisis Qu'#195#173'mico')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo44: TfrxMemoView
        Left = 491.338900000000000000
        Top = 26.456710000000000000
        Width = 200.315090000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = [fsBold]
        Frame.Typ = [ftBottom]
        HAlign = haCenter
        Memo.UTF8 = (
          'A')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo45: TfrxMemoView
        Left = 491.338900000000000000
        Top = 41.574830000000000000
        Width = 200.315090000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = [fsBold]
        Frame.Typ = [ftBottom]
        HAlign = haCenter
        Memo.UTF8 = (
          'NP')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo46: TfrxMemoView
        Left = 491.338900000000000000
        Top = 68.031540000000000000
        Width = 200.315090000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = [fsBold]
        Frame.Typ = [ftBottom]
        HAlign = haCenter
        Memo.UTF8 = (
          'NP')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo47: TfrxMemoView
        Left = 491.338900000000000000
        Top = 83.149660000000000000
        Width = 200.315090000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = [fsBold]
        Frame.Typ = [ftBottom]
        HAlign = haCenter
        Memo.UTF8 = (
          'NP')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo48: TfrxMemoView
        Left = 491.338900000000000000
        Top = 98.267780000000000000
        Width = 200.315090000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = [fsBold]
        Frame.Typ = [ftBottom]
        HAlign = haCenter
        Memo.UTF8 = (
          'NP')
        ParentFont = False
        VAlign = vaCenter
      end
      object Memo49: TfrxMemoView
        Left = 491.338900000000000000
        Top = 113.385900000000000000
        Width = 200.315090000000000000
        Height = 18.897650000000000000
        ShowHint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'tahoma'
        Font.Style = [fsBold]
        Frame.Typ = [ftBottom]
        HAlign = haCenter
        Memo.UTF8 = (
          'NP')
        ParentFont = False
        VAlign = vaCenter
      end
    end
  end
end
