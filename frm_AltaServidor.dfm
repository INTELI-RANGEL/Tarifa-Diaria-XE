object frmAltaServidor: TfrmAltaServidor
  Left = 0
  Top = 0
  AutoSize = True
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  ClientHeight = 169
  ClientWidth = 537
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object AdvPanel1: TAdvPanel
    Left = 0
    Top = 0
    Width = 537
    Height = 169
    Align = alClient
    BevelOuter = bvNone
    Color = clWhite
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Locked = True
    ParentFont = False
    TabOrder = 0
    UseDockManager = True
    Version = '2.0.2.1'
    BorderColor = clGray
    BorderShadow = True
    Caption.Color = clHighlight
    Caption.ColorTo = clBlue
    Caption.Font.Charset = DEFAULT_CHARSET
    Caption.Font.Color = clHighlightText
    Caption.Font.Height = -11
    Caption.Font.Name = 'Tahoma'
    Caption.Font.Style = []
    Caption.Indent = 2
    CollapsColor = clBtnFace
    CollapsDelay = 0
    ColorTo = 14938354
    ShadowColor = clBlack
    ShadowOffset = 0
    StatusBar.BorderColor = clSilver
    StatusBar.BorderStyle = bsSingle
    StatusBar.Font.Charset = DEFAULT_CHARSET
    StatusBar.Font.Color = clBlack
    StatusBar.Font.Height = -11
    StatusBar.Font.Name = 'Tahoma'
    StatusBar.Font.Style = []
    StatusBar.Color = 14938354
    StatusBar.ColorTo = clWhite
    Styler = AdvPanelStyler1
    FullHeight = 0
    object AdvGroupBox1: TAdvGroupBox
      Left = 12
      Top = 14
      Width = 513
      Height = 83
      BorderStyle = bsDouble
      Caption = 'Datos del Nuevo Servidor'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      object Label1: TLabel
        Left = 37
        Top = 49
        Width = 48
        Height = 13
        Caption = 'Servidor'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label5: TLabel
        Left = 41
        Top = 23
        Width = 44
        Height = 13
        Caption = 'Nombre'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label10: TLabel
        Left = 375
        Top = 49
        Width = 38
        Height = 13
        Caption = 'Puerto'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object edtServidor: TEdit
        Left = 91
        Top = 47
        Width = 246
        Height = 20
        Hint = 'Nombre del servidor o direcci'#243'n IP del mismo'
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
      object edtNombre: TEdit
        Left = 91
        Top = 21
        Width = 390
        Height = 20
        Hint = 'Nombre corto o especificaci'#243'n sencilla del servidor'
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        SkinData.SkinSection = 'EDIT'
        BoundLabel.Indent = 0
        BoundLabel.Font.Charset = DEFAULT_CHARSET
        BoundLabel.Font.Color = clWindowText
        BoundLabel.Font.Height = -11
        BoundLabel.Font.Name = 'MS Sans Serif'
        BoundLabel.Font.Style = []
        BoundLabel.Layout = sclLeft
        BoundLabel.MaxWidth = 0
        BoundLabel.UseSkinColor = True
      end
      object edtPuerto: TEdit
        Left = 421
        Top = 47
        Width = 60
        Height = 19
        Hint = 'N'#250'mero de puerto de comunicaci'#243'n comunmente 3306'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        Text = '3306'
      end
    end
    object btnAceptar: TAdvGlowButton
      Left = 320
      Top = 111
      Width = 78
      Height = 32
      AutoSize = True
      Caption = '&Aceptar'
      ImageIndex = 1
      Images = ImageList1
      NotesFont.Charset = DEFAULT_CHARSET
      NotesFont.Color = clWindowText
      NotesFont.Height = -11
      NotesFont.Name = 'Tahoma'
      NotesFont.Style = []
      TabOrder = 1
      OnClick = btnAceptarClick
      Appearance.BorderColor = 12631218
      Appearance.BorderColorHot = 10079963
      Appearance.BorderColorDown = 4548219
      Appearance.Color = 14671574
      Appearance.ColorTo = 15000283
      Appearance.ColorChecked = 7915518
      Appearance.ColorCheckedTo = 11918331
      Appearance.ColorDisabled = 15921906
      Appearance.ColorDisabledTo = 15921906
      Appearance.ColorDown = 7778289
      Appearance.ColorDownTo = 4296947
      Appearance.ColorHot = 15465983
      Appearance.ColorHotTo = 11332863
      Appearance.ColorMirror = 14144974
      Appearance.ColorMirrorTo = 15197664
      Appearance.ColorMirrorHot = 5888767
      Appearance.ColorMirrorHotTo = 10807807
      Appearance.ColorMirrorDown = 946929
      Appearance.ColorMirrorDownTo = 5021693
      Appearance.ColorMirrorChecked = 10480637
      Appearance.ColorMirrorCheckedTo = 5682430
      Appearance.ColorMirrorDisabled = 11974326
      Appearance.ColorMirrorDisabledTo = 15921906
      Appearance.GradientHot = ggVertical
      Appearance.GradientMirrorHot = ggVertical
      Appearance.GradientDown = ggVertical
      Appearance.GradientMirrorDown = ggVertical
      Appearance.GradientChecked = ggVertical
    end
    object btnCancelar: TAdvGlowButton
      Left = 433
      Top = 111
      Width = 83
      Height = 32
      AutoSize = True
      Caption = '&Cancelar'
      ImageIndex = 0
      Images = ImageList1
      NotesFont.Charset = DEFAULT_CHARSET
      NotesFont.Color = clWindowText
      NotesFont.Height = -11
      NotesFont.Name = 'Tahoma'
      NotesFont.Style = []
      TabOrder = 2
      OnClick = btnCancelarClick
      Appearance.BorderColor = 12631218
      Appearance.BorderColorHot = 10079963
      Appearance.BorderColorDown = 4548219
      Appearance.Color = 14671574
      Appearance.ColorTo = 15000283
      Appearance.ColorChecked = 7915518
      Appearance.ColorCheckedTo = 11918331
      Appearance.ColorDisabled = 15921906
      Appearance.ColorDisabledTo = 15921906
      Appearance.ColorDown = 7778289
      Appearance.ColorDownTo = 4296947
      Appearance.ColorHot = 15465983
      Appearance.ColorHotTo = 11332863
      Appearance.ColorMirror = 14144974
      Appearance.ColorMirrorTo = 15197664
      Appearance.ColorMirrorHot = 5888767
      Appearance.ColorMirrorHotTo = 10807807
      Appearance.ColorMirrorDown = 946929
      Appearance.ColorMirrorDownTo = 5021693
      Appearance.ColorMirrorChecked = 10480637
      Appearance.ColorMirrorCheckedTo = 5682430
      Appearance.ColorMirrorDisabled = 11974326
      Appearance.ColorMirrorDisabledTo = 15921906
      Appearance.GradientHot = ggVertical
      Appearance.GradientMirrorHot = ggVertical
      Appearance.GradientDown = ggVertical
      Appearance.GradientMirrorDown = ggVertical
      Appearance.GradientChecked = ggVertical
    end
  end
  object ImageList1: TImageList
    Height = 24
    Width = 24
    Left = 368
    Bitmap = {
      494C010102000400040018001800FFFFFFFFFF00FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000600000001800000001002000000000000024
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000BAB4AAFF928979FFAEA28FFFC4B7A3FFD6CABAFFE0D8CEFFDBD6
      CEFFB3ACA0FF928979FFEEEEEEFFFEFEFEFFFEFEFEFFFEFEFEFFE7E7E8FF0000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000BAB4AAFF928979FFAEA28FFFC4B7A3FFD6CABAFFE0D8CEFFDBD6
      CEFFB3ACA0FF928979FFEEEEEEFFFEFEFEFFFEFEFEFFFEFEFEFFE7E7E8FF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000928979FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF0EC
      E6FFF7F5F4FFFEFEFEFF1559E1FF1C2ADBFF1C2CDBFF1C29DBFF1556E1FFFEFE
      FEFFF3F3F3FF0000000000000000000000000000000000000000928979FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF0EC
      E6FFF7F5F3FFFEFEFEFF74D3A7FF27BF7BFF26C07FFF28BE78FF73D3A5FFFEFE
      FEFFF3F3F3FF0000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF6F6
      F6FF294AE0FF1C31DBFF1A3ADDFF1A3FDDFF1A41DDFF1A3EDDFF1B39DDFF1C2D
      DBFF1E3FDEFFE6E6E7FF000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF6F6
      F6FF66CA8FFF25C284FF25C284FF25C284FF25C284FF25C284FF25C284FF25C1
      81FF60C687FFE6E6E7FF00000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFE7E9
      FCFF1C30DBFF1A3EDDFFDFE0F1FF184DDFFF174FE0FF184CDFFFBFC4E8FF1A3A
      DDFF1C2DDBFFB2B7F3FF000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFEBF7
      EFFF25C284FF25C284FF25C284FF91D49EFF75B148FF25C284FF25C284FF25C2
      84FF25C181FFBFE7CAFF00000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB6A58CFFF1EEE9FFFFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFF1C2B
      DBFF1A3ADDFFFFFFFFFFFFFFFFFFDFE0F1FF145BE2FFBDC3E8FFFFFFFFFFFFFF
      FFFF1B37DCFF1C28DBFFDFDFE0FF0000000000000000958363FF9B8969FFA190
      71FFA9987BFFB6A58CFFF1EEE9FFFFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFF38BD
      75FF25C284FF25C284FFA9DBACFFFFFFFFFFFFFFFFFF25C283FF25C284FF25C2
      84FF25C284FF31BA6CFFDFDFE0FF000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FFFFFFFFFFCDC5
      B9FFB5AA99FFB5AA99FFB5AA99FFB5AA99FFB5AA99FFB5AA99FFFEFEFEFF495A
      E3FF1A42DDFF1751E0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF184C
      DFFF1A3EDDFF4957E3FFFEFEFEFF0000000000000000958363FFFFFFFFFFCDC5
      B9FFB5AA99FFB5AA99FFB5AA99FFB5AA99FFB5AA99FFB5AA99FFFEFEFEFF50CF
      9DFF25C284FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF25C284FF25C2
      84FF25C284FF50CE9BFFFEFEFEFF000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFF938670FF9386
      70FF938670FF938670FF938670FF938670FF938670FF938670FFBCBFF2FF485C
      E3FF4769E5FF1A39DDFF1942DEFFFFFFFFFFFFFFFFFFFFFFFFFF1A40DDFF1B34
      DCFF4363E4FF4959E3FFFEFEFEFF0000000000000000FFFFFFFF938670FF9386
      70FF938670FF938670FF938670FF938670FF938670FF938670FFC2E8CAFF50CF
      9DFF50CF9DFFFFFFFFFFFFFFFFFF6ACC96FFFFFFFFFFFFFFFFFF3F9200FF58C0
      79FF4CCD9AFF50CF9DFFFEFEFEFF000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFF75684FFF7568
      4FFF75684FFF75684FFF75684FFF7C6F56FF978A73FFAA9E89FFF2F4FDFF485B
      E3FF4A50DCFF4954E3FFBABDE7FFFFFFFFFFFFFFFFFFFFFFFFFFDCDDF0FF4A50
      E2FF4A50DCFF4958E3FFFEFEFEFF0000000000000000FFFFFFFF75684FFF7568
      4FFF75684FFF75684FFF75684FFF7C6F56FF978A73FFAA9E89FFF4FAF5FF50CF
      9DFF64C279FFFFFFFFFF57C078FF5CC27FFF75CC94FFFFFFFFFFFFFFFFFF6EC7
      84FF63C278FF50CF9DFFFEFEFEFF000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000BEB5A5FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFFEFEFEFF2530
      CEFF4A51D4FFFFFFFFFFFFFFFFFFFFFFFFFF4A50E0FFFFFFFFFFFFFFFFFFDBDC
      EFFF4A51D3FF4750D0FFF0F0F0FF000000000000000000000000BEB5A5FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFFEFEFEFF37AF
      4CFF5DBA62FF5FC172FF64C279FF69C57FFF6CC681FFFFFFFFFFFFFFFFFFFFFF
      FEFF5DB85DFF5AB452FFF0F0F0FF000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFF0F0F0FF1D25
      B9FF4A51CCFFEEEEFBFFFFFFFFFF4A51D7FF4A51D7FF4A51D7FFFFFFFFFF787C
      D4FF4A51CCFF1D25B9FFECECECFF0000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFF0F0F0FF3793
      00FF5EAF42FF5DB658FF5CBA64FF5CBE69FF5CBE6BFF5CBE69FFFFFFFFFFFFFF
      FFFF5EAF3EFF379200FFECECECFF000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFFEFE
      FEFF1D25B7FF4A50CAFF4A51CCFF4A51CDFF4A51CEFF4A51CDFF4A51CCFF4A50
      C8FF232BB7FFFEFEFEFF000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFFEFE
      FEFF379000FF5FAC33FF5EAF3FFF5EB148FF5EB24AFF5EB045FF5EAF3CFF5FAA
      33FF3C9207FFFEFEFEFF00000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFDDD6CAFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFEFEFEFF1C25B0FF3E44BDFF4A50C3FF4A50C4FF4A50C3FF484FC1FF1C25
      AEFFF6F6F6FF00000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFDDD6CAFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFEFEFEFF378A00FF549E25FF5FA533FF5FA633FF5FA533FF5DA331FF3789
      00FFF6F6F6FF0000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FFFFFFFFFFF0ED
      EAFFBBB09FFFBBB09FFFBBB09FFFBBB09FFFBBB09FFFBBB09FFFBBB09FFFBBB0
      9FFFBBB09FFFF0F0F0FFFEFEFEFF252EB6FF1C25ABFF232CB4FFFEFEFEFFD7D7
      D8FF0000000000000000000000000000000000000000958363FFFFFFFFFFF0ED
      EAFFBBB09FFFBBB09FFFBBB09FFFBBB09FFFBBB09FFFBBB09FFFBBB09FFFBBB0
      9FFFBBB09FFFF0F0F0FFFEFEFEFF3F9319FF388600FF3D9014FFFEFEFEFFD7D7
      D8FF000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFF9A8D78FF9A8D
      78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FF9A8D
      78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FFAAA193FF9A8D78FFFAF9F8FF0000
      00000000000000000000000000000000000000000000FFFFFFFF9A8D78FF9A8D
      78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FF9A8D
      78FF9A8D78FF9A8D78FF9A8D78FF9A8D78FFAAA193FF9A8D78FFFAF9F8FF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFF75684FFF7568
      4FFF75684FFF75684FFF796C53FFA89A83FFC7B9A4FFD6CABAFFE5DDD3FFDFD9
      D0FFA9A191FF75684FFF75684FFF75684FFF75684FFF928979FFF0EFEDFF0000
      00000000000000000000000000000000000000000000FFFFFFFF75684FFF7568
      4FFF75684FFF75684FFF796C53FFA89A83FFC7B9A4FFD6CABAFFE5DDD3FFDFD9
      D0FFA9A191FF75684FFF75684FFF75684FFF75684FFF928979FFF0EFEDFF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000FCFCFAFFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF0EC
      E6FFFAF8F6FFFAF8F6FFF1EEE9FFEAE5DDFFE6E0D6FFBBB5ABFF000000000000
      0000000000000000000000000000000000000000000000000000FCFCFAFFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF0EC
      E6FFFAF8F6FFFAF8F6FFF1EEE9FFEAE5DDFFE6E0D6FFBBB5ABFF000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF0EC
      E6FFFAF8F6FFFAF8F6FFF1EEE9FFEAE5DDFFE1D9CDFFD1C5B3FFF9F7F5FF0000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF0EC
      E6FFFAF8F6FFFAF8F6FFF1EEE9FFEAE5DDFFE1D9CDFFD1C5B3FFF9F7F5FF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF0EC
      E6FFFAF8F6FFFAF8F6FFF1EEE9FFEAE5DDFFE1D9CDFFD1C5B3FFC5B7A1FF0000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFB3A288FFBAAA91FFC4B6A0FFCCBEAAFFD6CABAFFE5DDD3FFF0EC
      E6FFFAF8F6FFFAF8F6FFF1EEE9FFEAE5DDFFE1D9CDFFD1C5B3FFC5B7A1FF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFBBAC94FFF7F6F3FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFDFCFBFFF1EEE9FFEAE5DDFFE1D9CDFFD1C5B3FFC5B7A1FF0000
      00000000000000000000000000000000000000000000958363FF9B8969FFA190
      71FFA9987BFFBBAC94FFF7F6F3FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFDFCFBFFF1EEE9FFEAE5DDFFE1D9CDFFD1C5B3FFC5B7A1FF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000958363FFE5E0D7FFFFFF
      FFFFE9E4DBFFEDE8E2FFF0ECE6FFF0ECE6FFEDE8E2FFE9E4DBFFE4DCD2FFDFD7
      CBFFD9CFC0FFD2C6B4FFCEC1ADFFD2C6B4FFFFFFFFFFD1C5B3FFC5B7A1FF0000
      00000000000000000000000000000000000000000000958363FFE5E0D7FFFFFF
      FFFFE9E4DBFFEDE8E2FFF0ECE6FFF0ECE6FFEDE8E2FFE9E4DBFFE4DCD2FFDFD7
      CBFFD9CFC0FFD2C6B4FFCEC1ADFFD2C6B4FFFFFFFFFFD1C5B3FFC5B7A1FF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFFE1D9CDFFE6DF
      D5FFEDE8E2FFF3F0EBFFF6F4F1FFF6F4F1FFF3F0EBFFEDE8E2FFE6DFD5FFE1D9
      CDFFDBD2C4FFD4C9B8FFCEC1ADFFCABCA7FFC4B6A0FFC0B199FFFFFFFFFF0000
      00000000000000000000000000000000000000000000FFFFFFFFE1D9CDFFE6DF
      D5FFEDE8E2FFF3F0EBFFF6F4F1FFF6F4F1FFF3F0EBFFEDE8E2FFE6DFD5FFE1D9
      CDFFDBD2C4FFD4C9B8FFCEC1ADFFCABCA7FFC4B6A0FFC0B199FFFFFFFFFF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFFFFE1D9CEFFE6E0
      D6FFEEE9E3FFF4F1EDFFFAF9F7FFFAF9F7FFF4F1EDFFEEE9E3FFE6E0D6FFE1D9
      CEFFDCD2C5FFD4C9B8FFCEC1AEFFCABCA8FFC4B6A0FF978D7CFFFFFFFFFF0000
      00000000000000000000000000000000000000000000FFFFFFFFE1D9CEFFE6E0
      D6FFEEE9E3FFF4F1EDFFFAF9F7FFFAF9F7FFF4F1EDFFEEE9E3FFE6E0D6FFE1D9
      CEFFDCD2C5FFD4C9B8FFCEC1AEFFCABCA8FFC4B6A0FF978D7CFFFFFFFFFF0000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000FFFFFFFFC2BC
      B1FFAEA699FFF3F0EBFFF6F4F1FFF6F4F1FFF3F0EBFFEDE8E2FFE6DFD5FFE1D9
      CDFFDBD2C4FFD4C9B8FFC4B8A5FF928979FFFFFFFFFFB8B2A7FF000000000000
      0000000000000000000000000000000000000000000000000000FFFFFFFFC2BC
      B1FFAEA699FFF3F0EBFFF6F4F1FFF6F4F1FFF3F0EBFFEDE8E2FFE6DFD5FFE1D9
      CDFFDBD2C4FFD4C9B8FFC4B8A5FF928979FFFFFFFFFFB8B2A7FF000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000BAB4AAFFD8D5CFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFDFDFCFFB0AA9EFFE7E6E2FF000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000BAB4AAFFD8D5CFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFDFDFCFFB0AA9EFFE7E6E2FF000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000060000000180000000100010000000000200100000000000000000000
      000000000000000000000000FFFFFF00F8001FF8001F000000000000C00007C0
      0007000000000000800003800003000000000000800003800003000000000000
      8000018000010000000000008000018000010000000000008000018000010000
      00000000800001800001000000000000C00001C0000100000000000080000180
      0001000000000000800003800003000000000000800007800007000000000000
      80000F80000F00000000000080001F80001F00000000000080001F80001F0000
      00000000C0003FC0003F00000000000080001F80001F00000000000080001F80
      001F00000000000080001F80001F00000000000080001F80001F000000000000
      80001F80001F00000000000080001F80001F000000000000C0003FC0003F0000
      00000000F801FFF801FF000000000000}
  end
  object AdvPanelStyler1: TAdvPanelStyler
    Tag = 0
    Settings.AnchorHint = False
    Settings.BevelInner = bvNone
    Settings.BevelOuter = bvNone
    Settings.BevelWidth = 1
    Settings.BorderColor = clGray
    Settings.BorderShadow = True
    Settings.BorderStyle = bsNone
    Settings.BorderWidth = 0
    Settings.CanMove = False
    Settings.CanSize = False
    Settings.Caption.Color = clHighlight
    Settings.Caption.ColorTo = clBlue
    Settings.Caption.Font.Charset = DEFAULT_CHARSET
    Settings.Caption.Font.Color = clHighlightText
    Settings.Caption.Font.Height = -11
    Settings.Caption.Font.Name = 'Tahoma'
    Settings.Caption.Font.Style = []
    Settings.Caption.Indent = 2
    Settings.Collaps = False
    Settings.CollapsColor = clBtnFace
    Settings.CollapsDelay = 0
    Settings.CollapsSteps = 0
    Settings.Color = clWhite
    Settings.ColorTo = 14938354
    Settings.ColorMirror = clNone
    Settings.ColorMirrorTo = clNone
    Settings.Cursor = crDefault
    Settings.Font.Charset = DEFAULT_CHARSET
    Settings.Font.Color = clWindowText
    Settings.Font.Height = -11
    Settings.Font.Name = 'Tahoma'
    Settings.Font.Style = []
    Settings.FixedTop = False
    Settings.FixedLeft = False
    Settings.FixedHeight = False
    Settings.FixedWidth = False
    Settings.Height = 120
    Settings.Hover = False
    Settings.HoverColor = clNone
    Settings.HoverFontColor = clNone
    Settings.Indent = 0
    Settings.ShadowColor = clBlack
    Settings.ShadowOffset = 0
    Settings.ShowHint = False
    Settings.ShowMoveCursor = False
    Settings.StatusBar.BorderColor = clSilver
    Settings.StatusBar.BorderStyle = bsSingle
    Settings.StatusBar.Font.Charset = DEFAULT_CHARSET
    Settings.StatusBar.Font.Color = clBlack
    Settings.StatusBar.Font.Height = -11
    Settings.StatusBar.Font.Name = 'Tahoma'
    Settings.StatusBar.Font.Style = []
    Settings.StatusBar.Color = 14938354
    Settings.StatusBar.ColorTo = clWhite
    Settings.TextVAlign = tvaTop
    Settings.TopIndent = 0
    Settings.URLColor = clBlue
    Settings.Width = 0
    Left = 136
    Top = 120
  end
end