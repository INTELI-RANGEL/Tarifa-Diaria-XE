object FrmImportaExportaActiv: TFrmImportaExportaActiv
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Importaci'#243'n / Exportaci'#243'n de actividades'
  ClientHeight = 278
  ClientWidth = 318
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object cxGroupBox1: TcxGroupBox
    AlignWithMargins = True
    Left = 3
    Top = 3
    Align = alClient
    Caption = 'Datos de Importaci'#243'n y Plantilla'
    Style.LookAndFeel.NativeStyle = False
    Style.LookAndFeel.SkinName = 'Lilian'
    StyleDisabled.LookAndFeel.NativeStyle = False
    StyleDisabled.LookAndFeel.SkinName = 'Lilian'
    StyleFocused.LookAndFeel.NativeStyle = False
    StyleFocused.LookAndFeel.SkinName = 'Lilian'
    StyleHot.LookAndFeel.NativeStyle = False
    StyleHot.LookAndFeel.SkinName = 'Lilian'
    TabOrder = 0
    Height = 272
    Width = 312
    object Label2: TLabel
      Left = 14
      Top = 47
      Width = 29
      Height = 13
      Caption = 'Fecha'
    end
    object Label3: TLabel
      Left = 14
      Top = 20
      Width = 71
      Height = 13
      Caption = 'Contrato( OT )'
    end
    object DFecha: TcxDateEdit
      Left = 102
      Top = 44
      Properties.DisplayFormat = 'yyyy-MM-dd'
      Properties.EditFormat = 'yyyy-MM-dd'
      Properties.ReadOnly = False
      Properties.ShowTime = False
      Style.BorderStyle = ebsOffice11
      Style.LookAndFeel.NativeStyle = False
      Style.LookAndFeel.SkinName = 'DevExpressStyle'
      Style.ButtonStyle = btsDefault
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.SkinName = 'DevExpressStyle'
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.SkinName = 'DevExpressStyle'
      StyleHot.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.SkinName = 'DevExpressStyle'
      TabOrder = 0
      Width = 121
    end
    object cxGroupBox2: TcxGroupBox
      AlignWithMargins = True
      Left = 6
      Top = 144
      Align = alBottom
      Caption = 'Opciones'
      Style.LookAndFeel.NativeStyle = False
      Style.LookAndFeel.SkinName = 'Lilian'
      StyleDisabled.LookAndFeel.NativeStyle = False
      StyleDisabled.LookAndFeel.SkinName = 'Lilian'
      StyleFocused.LookAndFeel.NativeStyle = False
      StyleFocused.LookAndFeel.SkinName = 'Lilian'
      StyleHot.LookAndFeel.NativeStyle = False
      StyleHot.LookAndFeel.SkinName = 'Lilian'
      TabOrder = 1
      Height = 115
      Width = 300
      object CbxSustituye: TcxCheckBox
        AlignWithMargins = True
        Left = 8
        Top = 20
        Margins.Left = 5
        Margins.Top = 5
        Margins.Right = 5
        Margins.Bottom = 5
        Align = alTop
        Caption = 'Actualizar Actividades'
        State = cbsChecked
        Style.LookAndFeel.NativeStyle = False
        Style.LookAndFeel.SkinName = 'Lilian'
        StyleDisabled.LookAndFeel.NativeStyle = False
        StyleDisabled.LookAndFeel.SkinName = 'Lilian'
        StyleFocused.LookAndFeel.NativeStyle = False
        StyleFocused.LookAndFeel.SkinName = 'Lilian'
        StyleHot.LookAndFeel.NativeStyle = False
        StyleHot.LookAndFeel.SkinName = 'Lilian'
        TabOrder = 0
        Transparent = True
        Width = 284
      end
      object CbxAvance: TcxCheckBox
        AlignWithMargins = True
        Left = 8
        Top = 49
        Margins.Left = 5
        Margins.Top = 5
        Margins.Right = 5
        Margins.Bottom = 5
        Align = alTop
        Caption = 'Validar avance'
        Style.LookAndFeel.NativeStyle = False
        Style.LookAndFeel.SkinName = 'Lilian'
        StyleDisabled.LookAndFeel.NativeStyle = False
        StyleDisabled.LookAndFeel.SkinName = 'Lilian'
        StyleFocused.LookAndFeel.NativeStyle = False
        StyleFocused.LookAndFeel.SkinName = 'Lilian'
        StyleHot.LookAndFeel.NativeStyle = False
        StyleHot.LookAndFeel.SkinName = 'Lilian'
        TabOrder = 1
        Transparent = True
        Width = 284
      end
      object BtnImportar: TcxButton
        AlignWithMargins = True
        Left = 199
        Top = 73
        Width = 95
        Height = 32
        Margins.Top = 0
        Margins.Bottom = 0
        Align = alRight
        Caption = 'Importar'
        TabOrder = 2
        OnClick = BtnImportarClick
        LookAndFeel.NativeStyle = False
        LookAndFeel.SkinName = 'DevExpressStyle'
        OptionsImage.ImageIndex = 0
      end
      object BtnExportar: TcxButton
        AlignWithMargins = True
        Left = 96
        Top = 73
        Width = 97
        Height = 32
        Margins.Top = 0
        Margins.Bottom = 0
        Align = alRight
        Caption = 'Generar plantilla'
        TabOrder = 3
        OnClick = BtnExportarClick
        LookAndFeel.NativeStyle = False
        LookAndFeel.SkinName = 'DevExpressStyle'
        OptionsImage.ImageIndex = 0
      end
    end
    object Panel2: TPanel
      Left = 3
      Top = 114
      Width = 306
      Height = 27
      Align = alBottom
      BevelOuter = bvNone
      Caption = 'Panel2'
      Ctl3D = False
      ParentCtl3D = False
      TabOrder = 2
      object txtFileName: TCurvyEdit
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 254
        Height = 21
        Cursor = crArrow
        Align = alClient
        TabOrder = 0
        Version = '1.1.0.0'
        Controls = <>
        EmptyText = 'Especifique una ruta =>'
        ReadOnly = True
      end
      object btnLoad: TcxButton
        AlignWithMargins = True
        Left = 263
        Top = 0
        Width = 40
        Height = 27
        Margins.Top = 0
        Margins.Bottom = 0
        Align = alRight
        TabOrder = 1
        OnClick = btnLoadClick
        LookAndFeel.NativeStyle = False
        LookAndFeel.SkinName = 'DevExpressStyle'
        OptionsImage.ImageIndex = 0
        OptionsImage.Images = img1
      end
    end
    object CmbContratos: TcxComboBox
      Left = 101
      Top = 17
      TabOrder = 3
      Text = 'NA'
      Width = 121
    end
  end
  object Lineas: TAdvEdit
    Left = 276
    Top = 23
    Width = 25
    Height = 21
    EditType = etNumeric
    LabelFont.Charset = DEFAULT_CHARSET
    LabelFont.Color = clWindowText
    LabelFont.Height = -11
    LabelFont.Name = 'Tahoma'
    LabelFont.Style = []
    Lookup.Separator = ';'
    Color = clWindow
    TabOrder = 1
    Text = '50'
    Visible = False
    Version = '2.9.3.1'
  end
  object PAthGuardar: TSaveDialog
    DefaultExt = '.xlsx'
    FileName = 'C:\Users\ricky_000\Desktop\rebuilt.Windows 95.zip'
    Filter = 'Excel|*.xlsx'
    FilterIndex = 0
    Options = [ofHideReadOnly, ofEnableSizing, ofDontAddToRecent]
    OptionsEx = [ofExNoPlacesBar]
    Title = 'Seleccione ruta y nombre del archivo'
    Left = 240
    Top = 24
  end
  object img1: TcxImageList
    Height = 24
    Width = 24
    FormatVersion = 1
    DesignInfo = 5243120
    ImageInfo = <
      item
        Image.Data = {
          36090000424D3609000000000000360000002800000018000000180000000100
          2000000000000009000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000040C10161D60
          7CA71D698AB71D698AB71D698AB71D698AB71D698AB71D698AB71D698AB71D69
          8AB71D698AB71D698AB71D698AB71D698AB71D698AB71D698AB71D698AB71D69
          8AB71D698AB71D698AB71E698BB8184E668900010102000000000C202B3B2196
          CAFF059CE0FF079BDEFF079BDEFF079BDEFF079BDEFF079BDEFF079BDEFF079B
          DEFF079BDEFF079BDEFF079BDEFF079BDEFF079BDEFF079BDEFF079BDEFF079B
          DEFF079BDEFF079BDEFF059CE0FF1F94C6FA050F131C000000000B232C3F2995
          C4FF0EA0E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1
          E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1E1FF0EA1
          E1FF0EA1E1FF0EA1E1FF0CA2E3FF1C9BD2FF0C253043000000000D2935482F94
          BFFF1DA5DFFF1BA8E4FF1BA7E3FF1BA7E3FF1BA7E3FF1BA7E3FF1BA7E3FF1BA7
          E3FF1BA7E3FF1BA7E3FF1BA7E3FF1BA7E3FF1BA7E3FF1BA7E3FF1BA7E3FF1BA7
          E3FF1BA7E3FF1BA7E3FF1AA8E4FF21A3DAFF133D5370000000000F2F3D523D9D
          C4FF2BA8DBFF2AB0E6FF2AAEE5FF2AAEE5FF2AAEE5FF2AAEE5FF2AAEE5FF2AAE
          E5FF2AAEE5FF2AAEE5FF2AAEE5FF2AAEE5FF2AAEE5FF2AAEE5FF2AAEE5FF2AAE
          E5FF2AAEE5FF2AAEE5FF2AAFE6FF2AACE1FF1D5A749C000000001033435A4CAC
          CFFF37A9D6FF3AB9E9FF3AB7E7FF3AB7E7FF3AB7E7FF3AB7E7FF3AB7E7FF3AB7
          E7FF3AB7E7FF3AB7E7FF3AB7E7FF3AB7E7FF3AB7E7FF3AB7E7FF3AB7E7FF3AB7
          E7FF3AB7E7FF3AB7E7FF3AB7E7FF39B8E8FF28779BC900010101123849635BBA
          D9FF40AAD2FF4BC1EBFF4ABFE9FF4AC0E9FF4AC0E9FF4AC0E9FF4AC0E9FF4AC0
          E9FF4AC0E9FF4AC0E9FF4AC0E9FF4AC0E9FF4AC0E9FF4AC0E9FF4AC0E9FF4AC0
          E9FF4AC0E9FF4AC0E9FF4ABFE9FF4BC2EBFF3597BFED050E1219143E506D6BC7
          E2FF49ADD1FF5AC8EBFF5BC9ECFF5BC9EBFF5BC9EBFF5BC9EBFF5BC9EBFF5BC9
          EBFF5BC9EBFF5BC9EBFF5BC9EBFF5BC9EBFF5BC9EBFF5BC9EBFF5BC9EBFF5BC9
          EBFF5BC9EBFF5BC9EBFF5BC9EBFF5ECCEEFF47B1D7FF0C212D3E1544577778D1
          E8FF54B3D3FF64C9E8FF6DD3EFFF6CD2EEFF6CD2EEFF6CD2EEFF6CD2EEFF6CD2
          EEFF6CD2EEFF6CD2EEFF6CD2EEFF6CD2EEFF6CD2EEFF6CD2EEFF6CD2EEFF6CD2
          EEFF6CD2EEFF6CD2EEFF6CD2EEFF6FD5F1FF5CC1E2FF123C4E6A1749608182D9
          EDFF61BCD9FF69C8E4FF81DEF4FF7EDBF1FF7EDBF1FF7EDBF1FF7EDBF1FF7EDB
          F1FF7EDBF1FF7EDBF1FF7EDBF1FF7EDBF1FF7EDBF1FF7EDBF1FF7DDAF1FF7CD9
          F0FF7CDAF0FF7CDAF0FF7CD9F0FF7FDCF2FF76D3ECFF1D5870951F52648988DD
          F0FF6FC8E0FF5EBAD8FF86DDF0FF82D9EDFF82D9EDFF82D9EDFF82D9EDFF83DA
          EDFF83DAEDFF83DAEDFF84DAEDFF84DAEDFF84DAEDFF84DAEDFF8CE1F3FF8CE2
          F3FF8AE1F3FF8BE1F3FF8BE2F3FF8BE1F3FF85DCF0FF368AABDD21586D928CDF
          EFFF8DE0F0FF55B2D1FF55B2D2FF54B2D2FF52B1D1FF51AFD1FF4FAFD1FF4EAE
          D0FF4CADD0FF4AACD0FF48ABD0FF46AACFFF44A9CFFF3DA4CCFF5AB6D4FF9CEC
          F7FF9BEBF7FF9AEBF7FF9FEFFAFF60B9D5FB276F8BBB1C52698E225F7AA392E2
          F0FF9DECF7FF97E8F6FF90E4F4FF8CE3F3FF86DFF3FF81DDF2FF7BDAF1FF75D6
          F0FF6FD3EFFF68D0EEFF61CCEDFF5BCAECFF54C5ECFF4EC2ECFF39A7D2FF6BBF
          D8FF8BDAE9FF89D8E8FF86D7E7FF2B718EBB0000000000000000296D8AB698E6
          F2FFA1EDF7FF99E9F5FF95E7F4FF91E5F4FF8BE2F3FF87E0F2FF81DDF2FF7BDA
          F1FF76D6F0FF6FD3EFFF69D1EEFF63CEEEFF5CCAECFF57C7ECFF4FC2EAFF3DAB
          D5FF3AA7D2FF37A7D3FF329FCCFF1544597A000000000000000018475A7A3586
          A6D66DC2DBFEA2EFF8FF99E8F5FF95E6F4FF90E4F3FF8CE2F3FF87DFF2FF81DC
          F1FF7CDAF1FF76D7F0FF6CD0ECFF3C95BBEF2E7D9FCE2E81A2D02D80A2D02C81
          A3D02A80A3D02A80A4D1277CA1CE0D2732450000000000000000000000000105
          06092F7996C5A1ECF6FFA6F1F9FFA0EEF8FF9CECF7FF97EAF7FF93E7F6FF8CE4
          F5FF87E1F5FF85E1F6FF66C1E0FF0D4F5D710002020301020304010203040102
          0304010203040102030401020203000000000000000000000000000000000000
          0000113646605AB3D1FA6AC0DAFE65BED8FD63BDD8FD61BCD7FD5FBAD7FD5DB9
          D7FD5AB8D6FD5BB9D9FE3D97BAEC021315210000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000061218211C556D921F5972981F5871971E5871971E5871971E5871971E58
          71971E5871971F5A739A17495F7F000303060000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
      end>
  end
  object dlgOpenExcel: TJvOpenDialog
    Filter = 'Excel|*.xlsx'
    Title = 'Abrir archivo'
    Height = 458
    Width = 571
    Left = 240
    Top = 48
  end
end
