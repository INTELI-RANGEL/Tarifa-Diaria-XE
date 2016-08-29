object frmContenidoNotacampo: TfrmContenidoNotacampo
  Left = 0
  Top = 0
  Caption = 'Contenido Nota de Campo'
  ClientHeight = 435
  ClientWidth = 979
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Visible = True
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 213
    Width = 979
    Height = 3
    Cursor = crVSplit
    Align = alBottom
    ExplicitLeft = 8
    ExplicitTop = 160
    ExplicitWidth = 871
  end
  object LbContenido: TPageControl
    Left = 0
    Top = 216
    Width = 979
    Height = 219
    ActivePage = TabSheet11
    Align = alBottom
    TabOrder = 0
    object TabSheet11: TTabSheet
      Caption = 'PRINCIPAL'
      ImageIndex = 10
      object PnlPrincipal: TPanel
        Left = 0
        Top = 0
        Width = 454
        Height = 191
        Align = alLeft
        TabOrder = 0
        object CmbTipo: TAdvDBComboBox
          Left = 81
          Top = 13
          Width = 195
          Height = 21
          Color = 15138559
          Version = '1.0.1.1'
          Visible = True
          DataField = 'ltipo'
          DataSource = dsContenido
          DropWidth = 0
          Enabled = True
          ItemIndex = 0
          ItemHeight = 13
          Items.Strings = (
            'PORTADA'
            'INDICE'
            'PRESENTACION'
            'OFICIO-T1'
            'OFICIO-T2'
            'OFICIO-T3'
            'OFICIO-T4'
            'OFICIO-T5'
            'OFICIO-T6'
            'OFICIO-T7')
          Items.StoredStrings = (
            'PORTADA'
            'INDICE'
            'PRESENTACION'
            'OFICIO-T1'
            'OFICIO-T2'
            'OFICIO-T3'
            'OFICIO-T4'
            'OFICIO-T5'
            'OFICIO-T6'
            'OFICIO-T7')
          LabelCaption = 'TIPO'
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          TabOrder = 0
          OnChange = CmbTipoChange
        end
        object DBAdvEdit1: TDBAdvEdit
          Left = 77
          Top = 67
          Width = 275
          Height = 21
          LabelCaption = 'NOMBRE'
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          MaxLength = 30
          TabOrder = 1
          Text = 'DBAdvEdit1'
          Visible = True
          Version = '2.9.3.1'
          DataField = 'snombreportada'
          DataSource = dsContenido
        end
        object DBAdvEdit2: TDBAdvEdit
          Left = 81
          Top = 40
          Width = 47
          Height = 21
          LabelCaption = 'ORDEN'
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          TabOrder = 2
          Text = 'DBAdvEdit1'
          Visible = True
          Version = '2.9.3.1'
          DataField = 'iorden'
          DataSource = dsContenido
        end
        object cxLabel1: TcxLabel
          Left = 2
          Top = 94
          Caption = 'DESCRIPCION'
        end
        object CbxIndice: TcxCheckBox
          Left = 200
          Top = 40
          Caption = 'EN INDICE'
          TabOrder = 4
          Width = 84
        end
        object DBMemo1: TDBMemo
          Left = 77
          Top = 94
          Width = 275
          Height = 46
          Color = 15138559
          DataField = 'sDescripcion'
          DataSource = dsContenido
          TabOrder = 5
        end
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'PORTADA'
      object PnlPortada: TPanel
        Left = 0
        Top = 0
        Width = 461
        Height = 191
        Align = alLeft
        Caption = 'PnlPortada'
        TabOrder = 0
        object ImgPortada: TImage
          Left = 136
          Top = 16
          Width = 273
          Height = 134
          Picture.Data = {
            0A544A504547496D6167651F0F0000FFD8FFE000104A46494600010101012C01
            2C0000FFDB004300080606070605080707070909080A0C140D0C0B0B0C191213
            0F141D1A1F1E1D1A1C1C20242E2720222C231C1C2837292C30313434341F2739
            3D38323C2E333432FFDB0043010909090C0B0C180D0D1832211C213232323232
            3232323232323232323232323232323232323232323232323232323232323232
            32323232323232323232323232FFC0001108003000C503012200021101031101
            FFC4001F0000010501010101010100000000000000000102030405060708090A
            0BFFC400B5100002010303020403050504040000017D01020300041105122131
            410613516107227114328191A1082342B1C11552D1F02433627282090A161718
            191A25262728292A3435363738393A434445464748494A535455565758595A63
            6465666768696A737475767778797A838485868788898A92939495969798999A
            A2A3A4A5A6A7A8A9AAB2B3B4B5B6B7B8B9BAC2C3C4C5C6C7C8C9CAD2D3D4D5D6
            D7D8D9DAE1E2E3E4E5E6E7E8E9EAF1F2F3F4F5F6F7F8F9FAFFC4001F01000301
            01010101010101010000000000000102030405060708090A0BFFC400B5110002
            0102040403040705040400010277000102031104052131061241510761711322
            328108144291A1B1C109233352F0156272D10A162434E125F11718191A262728
            292A35363738393A434445464748494A535455565758595A636465666768696A
            737475767778797A82838485868788898A92939495969798999AA2A3A4A5A6A7
            A8A9AAB2B3B4B5B6B7B8B9BAC2C3C4C5C6C7C8C9CAD2D3D4D5D6D7D8D9DAE2E3
            E4E5E6E7E8E9EAF2F3F4F5F6F7F8F9FAFFDA000C03010002110311003F00F7FA
            323D6AAEA3671EA1A7DC59CC1FCAB88DA27D8C55B6B290704118383D7B57C9FE
            22F03C7A2FC5D8FC39656175A8DA1961963B6320F3278B687750DC01C0719E3A
            67DE803EBACD158FE1C061D02CD5F4E9B4C48E1D82D6E26123C4ABC005C139E0
            039249E79C1E2AFC5A8595C094C3796F279240976480EC3E87078FC6802CD191
            EB546CF58D33536956C351B3BB6858A482DE7590A30EA080783EC6BCFBE2ECAA
            3C3F713C3710ADC436CF863ACBDA9438246225E25638E01EA78A00F4FC8F5A09
            03A9AF08F811E28D26C742D49358D72D21BDB8BDDC16EEE423C83628046E3F36
            7047E15EDB7FA859E9968F757F750DB5B272D2CD20455FA934016B23D69AD222
            21767554032589C002B8C8BE2C781A7BE6B34F11DA09579DCE1D50FF00C0C80A
            7F3ADBF104F652F84F50B9966B5365F65791E69A1FB4C5B369258A03F38C7607
            9A00B235ED1DAF21B31AB589BA9C911422E137C840C9DAB9C9C0F4AF99BE25F8
            7B56F13FC72D6F4ED1ACA4BBBB290BEC42176A8B78F2C59B01474192464903A9
            1557E1725BCFF1C34F6B59525B76BAB9789E387CA56511CA5484CFC9918C2F38
            E99E2BB3D6751F11786FE3CF8875ED1FC317BAD5AB4515A4C2082465E6085B01
            D5480C0843820F07B641001E3FE27F08EBDE12B8821D774E7B479D4B45F3ABAB
            007070CA48C8E32339191EB5835E99F177C59AE78A65D28EADE17BAD0A1B7597
            C84B957DD2B315DE7732A8206138038CF24E463CFF004FD2752D5EE1ADF4DD3E
            EEF67543218EDA169182F037614138E473EE28029D18CF4AD0D4F42D5F45F2BF
            B574ABEB0F373E5FDAADDE2DF8C671B80CE323F3AAB6B6F35D5CC76F6F0C934D
            2B048E3452CCEC4E02803924F4E28021C1F4A307D2BBF83E0E78FAE6DE29A3F0
            F384950380F710A3004646559C107D4100D727ADE87AA787EFDAC356B09ECEE9
            79D92A6370C91B94F465C82011907140199456C59784FC47A9D9A5E58787F55B
            AB5933B2682CE4746C120E180C1E411F514965E14F11EA56697963E1FD56EAD6
            4CEC9A0B391D1B07070C060F208FC28032319A5C1F4AE93C1FA6EAABE39B0482
            C2EFED561791CB3A2D9493B5B88E55DCCF12618853D4707B641AFB03C3FA849A
            A68F0CF711BC776142DC836B3DBA79BB41728B32AB14C93838F6CE41A00F86CA
            91D411495B5E22F0B6B7E15BA5B5D6F4D9ACA46276170191F0013B5C7CAD8DCB
            9C138279AC5A0028A28A00FBFCD78AEB9094FDA87C3AC3FE5A5817FCA39C7F4A
            F6793EEFB7D715F37DE78C6C9FE36E95AFAEA175AD416692C1225B69C6392052
            2450304FCF832124F1C5007BAF8B7C33A6F8AF42934ED52169610DE6A2ACAD1E
            1C0214E54F38CF4391ED5E19F063C0961E2BD0751BAD66E6EE6D3D6E7CA1A7C7
            3BC7133850DBDF6B7CC7918E9DFAF18F69F10F884D9781EE35FB29ADEDC7D9D6
            589AFE29762EEC7DF4405FA1E98C822BCDBF678D774D1A05FE866E40D48DDBDD
            084AB64C45225DD9C60FCC0FBD0067DCF852D7C11F1F7C376FA0C535AD8DEC5B
            D904ACE3FE5A075CB1C9180A793DEBD1FC59F0DBC2DE204BDD46F74749352781
            D84C923C64BEDE090A402781D41AE6BC6DADE8D6BF1ABC1AF737D144D682E16E
            77647965D311EE3EE4E01E9EBDEBA9F14F8F34FF00096B90DA6BD88349BCB42D
            15D089E4FDE86C3232A83C6D208FF8167B50079D7C09F0C685AD7842FAEF53D1
            ACAF274BF68D1EE6DD652ABB2338E7EA7F3AE6BC7DAE4BE31F8C767E1DB8B966
            D220D4E1B1585095192EA9231C756C9233D8018EF5D6FECF7AFE930786AE7449
            6F624D4A7D41E54B7390CC9E5C6011C7AA9FF26B95F8B3E0BD5FC31E3493C5DA
            5C0CF6135C0BC1344848B79810C778EC0B0CE7A73401DA7C48B9D22E4C7E158F
            5591B4F85147F616836426B967523866E55157838201E0F5C015ADF067C3BE2A
            F0E6837563E208560B27712DAC0D2EE92227EF8C0C80A78207A939E49AA1E11F
            8B5E077B1799ACFF00B3757BA6325CDB5BD8BC8F3CC792CA514EEC9FEF115D27
            867C63E20D57C52F69A9785EEF4CD2AE2DBCEB19E68C96E393E690485241185E
            08C6280388F0EDB28FDA775C3800C70BCA3233CB2479C7D771FCEBA887C7FA2E
            85F13357F0BDF2DD25E6A3A940D0CAB1068816B5B7455273BB25971F778C8C90
            39AF369FC5DA6D87C77D6B5CB6F11DA5A69F242918BB36AD771CE3645B9008CE
            472ADF37FB38ABB3683A878CBF681B9D5B4654974EB0BEB0B8BA9E43B1515634
            61F29C3163E5B018079C6700E680347E3AC76FAEF8C3C1BE1A5BAF26E6699924
            250B79493491C6AF8C80DF71F8CE7E5E71919DAF18F882C7E0A78434DD23C3BA
            62497176EE6392E3E65629B7CC9252082EE77280381F40A14C1F14BC21ACB78B
            2C7C796296B7167A0410CF35ABCC6396410CAD2BE3E52A06D3EB9E0E01E335BC
            7FA47823E24DC69BAB278E748D3678EDF633C932334B19F99032348A50A92F9E
            33F3F3D3140167E167C5DD47C67E249F47D6AD74FB691A032DB340590B3291B9
            36BB36E2412DC1C808783D97E0EF82B47D3353F11EA4B13CF79A7EB171A65B4B
            390C6289368C80001B9B760B71C74C0273D4C5E0CD22D3C4D7BE3BD16D20BCBE
            9F4F2D6D6D0B22C73CAC09F3164390A641B1770C0C162776E38E0FE0FF00C42B
            78B5ED63C3FAFC3FD9FAAEA5AA4B7685814433BED0D0956E51815F9413CE76F0
            7018032F5AD37C69F13BE24DFC5A75CDDD8E87A4EA5F6613B4E4456AF164798A
            0152F29C16E32577A82CAB83573E395CE9969E0BD0B419B548751D7EC658D657
            6C34FB0438777EA537931B618E5B8233B6BADF8B3E19F1C7886C913C33A9E2C3
            CAD973A624821798924121F8DCA55FE64660BF28E0935E6FE23F85DA4F81FE1A
            6A175AEDFD9CFE2699D4DA224E5115448A1C44BF2990ED6CB12BC7180319201E
            A9F063C4B3F88FE1E5B79D6F1C4DA6B8D3C14638916344DAC73D0ED650793920
            9E3381E71E1CF8E3E20D63E2269B6D3DB5AAE977B70966B668394123E164F308
            2C5C0201E8A40FBA09C8EA7E11E85349F06EF8E96F756BA8EA8970229E5B92B1
            AC8034692461092806065B1B8953D542D793783BC15E21B7F8AFA7E952586DBC
            D2AE6DEF2F23F3A33E542AF1B16CE70701D4E064F3D2803B1F8C7E21D4FC1BF1
            4E1D4B409E3B1BDBAD2112791604632832BF50C082711A0C9E70A074AF42F82D
            E26D63C5BE10BCD435ABCFB5DD47A83C2AE6344C208E360308A075635E75F19F
            46D4357F8ADA0C57163E5595EF93A7DACDF6A51F68F9C1739018C78336DE54F4
            C80738AEEB40F02EABE19F03F8CBC2F15BC7710DDA5CC9A64C2E41697CD89A31
            13E5576B011A127EE9F3383C1A00F9F3C4DE3BF12F8BACEDE0D6F5592E6081CB
            A442348D43118C908A0138CE09E9938EA6B98ADAF12F85F58F09DFC761ADDA7D
            96EA48C4CA86547CA1240395247553C67B7D2B16800A28A2803EFE6191D334C0
            A463A9C549450047B79040FF003F9D054B020E7AF1CD494500645D6836D75E21
            B1D664DDF68B3865863181B71215249CF391B703FDE35A880E075E9D09A7D140
            0D6048FF00EBD359723057208C11ED52514019B69A1E97657925DDA699676F75
            281E64D0DBA233FD580C9ABEC338C0C9A7D1401CF58785AD6C3C5DA87882058E
            396F6D62B768D220B828589627B920A0FF00800EBDBE73F897AC5CF87BE3A5F6
            AB60FB2E2D25B6923DC480C04117C8D82095232A47704D7D5A6BE58F8BFE13F1
            26A7F14759BCB0F0FEAB776D2793B2682CA4911B10A038603079047E140193E3
            CD461F1FFC572BA3EA1F6AB6BD96DAD2CE59F7AAA655148C32E5543972703D4E
            39E7DAFC5DA9E93F0C3C131E9BA545A2C9E588BED5A6DE9459AFADD8794EFB17
            1BE46C0CB104615B838AF9E2DBE1F78C6794A2F85B5952119F32594883E5049E
            58019C0381D49C0192403DFEAFE22F889E23D321B1D77E1CAEA42250AB3CFA3D
            D2CB9F97730746529BB68CEDDA0F4C638A008FE05DD6BF71E3CF2F4D13C3A00F
            365BEB689D8DBC5B94ECE1C9F9B72A00412E429E480D5CEFC62B882E3E2A6B92
            413C72A2BC51EE460C372C28AC3EA08208CF5047AD7557BE2FF8AB2686BA4E97
            E0DB8D12D8647FC4AF459A22A181C85CEE0BCB13B970D9EE2BCE5FC0DE2F6031
            E15D73AF4FECE9B8FF00C7680196FE33F12DA5BC76F6FE23D5A18635091C697B
            205451C28001C003D3A0F4AC5326411BBBE735B5FF0008278C3FE854D73FF05D
            37FF001347FC209E30FF00A1535CFF00C174DFFC4D00626E07038C738F5FCEAD
            E93AA4DA46AD63A95B8469ECE78EE235933B4B230619E791902B43FE104F187F
            D0A9AE7FE0BA6FFE268FF8413C61FF0042A6B9FF0082E9BFF89A00B1E37F174B
            E33F15DC6B5246F02C8889140D3F9A21555030AC40E0B066C6072C7B926B9BDC
            31D71F9FF8D6E7FC209E30FF00A1535CFF00C174DFFC4D1FF08278C3FE854D73
            FF0005D37FF1340186ED9403DFF2FF003FD4D475D07FC209E30FFA1535CFFC17
            4DFF00C4D1FF0008278C3FE854D73FF05D37FF0013401CFD15D07FC209E30FFA
            1535CFFC174DFF00C4D1401FFFD9}
          Stretch = True
        end
        object btnAdd: TAdvGlowButton
          Left = 11
          Top = 3
          Width = 86
          Height = 25
          Hint = 'Nuevo registro (CTRL + Insert)'
          Caption = 'Cargar'
          ImageIndex = 0
          Images = frmBarra1.ImgBtns
          NotesFont.Charset = DEFAULT_CHARSET
          NotesFont.Color = clWindowText
          NotesFont.Height = -11
          NotesFont.Name = 'Tahoma'
          NotesFont.Style = []
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnClick = btnAddClick
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
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'INDICE'
      ImageIndex = 1
      object PnlIndice: TPanel
        Left = 0
        Top = 0
        Width = 185
        Height = 191
        Align = alLeft
        TabOrder = 0
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'PRESENTACION'
      ImageIndex = 2
      object PnlPresentacion: TPanel
        Left = 0
        Top = 0
        Width = 185
        Height = 191
        Align = alLeft
        TabOrder = 0
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'OFICIO-T1'
      ImageIndex = 3
      object PnlOficio1: TPanel
        Left = 0
        Top = 0
        Width = 968
        Height = 191
        Align = alLeft
        TabOrder = 0
        object DBMemo2: TDBMemo
          Left = 65
          Top = 41
          Width = 482
          Height = 89
          Color = 15138559
          DataField = 'mtexto2'
          DataSource = dsContenido
          TabOrder = 0
        end
        object cxLabel2: TcxLabel
          Left = 13
          Top = 41
          Caption = 'CUERPO'
        end
        object GroupBox1: TGroupBox
          Left = 608
          Top = 8
          Width = 353
          Height = 77
          Caption = 'DIRIGIDO A:'
          TabOrder = 2
          object DBAdvEdit3: TDBAdvEdit
            Left = 66
            Top = 20
            Width = 275
            Height = 21
            LabelCaption = 'NOMBRE'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 0
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'snombref1'
            DataSource = dsContenido
          end
          object DBAdvEdit4: TDBAdvEdit
            Left = 66
            Top = 47
            Width = 275
            Height = 21
            LabelCaption = 'CARGO'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 1
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'scargof1'
            DataSource = dsContenido
          end
        end
        object GroupBox2: TGroupBox
          Left = 608
          Top = 87
          Width = 353
          Height = 82
          Caption = 'ATENTAMENTE:'
          TabOrder = 3
          object DBAdvEdit5: TDBAdvEdit
            Left = 66
            Top = 19
            Width = 275
            Height = 21
            LabelCaption = 'NOMBRE'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 0
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'snombref2'
            DataSource = dsContenido
          end
          object DBAdvEdit6: TDBAdvEdit
            Left = 66
            Top = 46
            Width = 275
            Height = 21
            LabelCaption = 'CARGO'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 1
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'scargof2'
            DataSource = dsContenido
          end
        end
        object DBAdvEdit7: TDBAdvEdit
          Left = 65
          Top = 14
          Width = 482
          Height = 21
          LabelCaption = 'ASUNTO'
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          TabOrder = 4
          Text = 'DBAdvEdit1'
          Visible = True
          Version = '2.9.3.1'
          DataField = 'stexto1'
          DataSource = dsContenido
        end
        object DBMemo3: TDBMemo
          Left = 65
          Top = 136
          Width = 482
          Height = 39
          Color = 15138559
          DataField = 'mtexto3'
          DataSource = dsContenido
          TabOrder = 5
        end
        object cxLabel3: TcxLabel
          Left = 2
          Top = 136
          Caption = 'DESPEDIDA'
        end
      end
    end
    object TabSheet5: TTabSheet
      Caption = 'OFICIO-T2'
      ImageIndex = 4
      object PnlOficio2: TPanel
        Left = 0
        Top = 0
        Width = 625
        Height = 191
        Align = alLeft
        TabOrder = 0
        object DBAdvEdit8: TDBAdvEdit
          Left = 88
          Top = 14
          Width = 459
          Height = 21
          LabelCaption = 'EMBARCACI'#211'N'
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          TabOrder = 0
          Text = 'DBAdvEdit1'
          Visible = True
          Version = '2.9.3.1'
          DataField = 'stexto1'
          DataSource = dsContenido
        end
      end
    end
    object TabSheet6: TTabSheet
      Caption = 'OFICIO-T3'
      ImageIndex = 5
      object PnlOficio3: TPanel
        Left = 0
        Top = 0
        Width = 953
        Height = 191
        Align = alLeft
        TabOrder = 0
        object GroupBox3: TGroupBox
          Left = 592
          Top = 3
          Width = 353
          Height = 62
          Caption = 'ATENCION:'
          TabOrder = 0
          object DBAdvEdit9: TDBAdvEdit
            Left = 66
            Top = 14
            Width = 275
            Height = 21
            LabelCaption = 'NOMBRE'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 0
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'snombref1'
            DataSource = dsContenido
          end
          object DBAdvEdit10: TDBAdvEdit
            Left = 66
            Top = 38
            Width = 275
            Height = 21
            LabelCaption = 'CARGO'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 1
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'scargof1'
            DataSource = dsContenido
          end
        end
        object GroupBox4: TGroupBox
          Left = 592
          Top = 66
          Width = 353
          Height = 62
          Caption = 'ATENTAMENTE:'
          TabOrder = 1
          object DBAdvEdit11: TDBAdvEdit
            Left = 66
            Top = 15
            Width = 275
            Height = 21
            LabelCaption = 'NOMBRE'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 0
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'snombref2'
            DataSource = dsContenido
          end
          object DBAdvEdit12: TDBAdvEdit
            Left = 66
            Top = 38
            Width = 275
            Height = 21
            LabelCaption = 'CARGO'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 1
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'scargof2'
            DataSource = dsContenido
          end
        end
        object GroupBox5: TGroupBox
          Left = 592
          Top = 125
          Width = 353
          Height = 62
          Caption = 'RECIBI DE CONFORMIDAD:'
          TabOrder = 2
          object DBAdvEdit13: TDBAdvEdit
            Left = 66
            Top = 15
            Width = 275
            Height = 21
            LabelCaption = 'NOMBRE'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 0
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'snombref3'
            DataSource = dsContenido
          end
          object DBAdvEdit14: TDBAdvEdit
            Left = 66
            Top = 38
            Width = 275
            Height = 21
            LabelCaption = 'CARGO'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 1
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'scargof3'
            DataSource = dsContenido
          end
        end
        object DBAdvEdit15: TDBAdvEdit
          Left = 65
          Top = 14
          Width = 521
          Height = 21
          LabelCaption = 'ASUNTO'
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          TabOrder = 3
          Text = 'DBAdvEdit1'
          Visible = True
          Version = '2.9.3.1'
          DataField = 'stexto1'
          DataSource = dsContenido
        end
        object DBMemo4: TDBMemo
          Left = 65
          Top = 41
          Width = 521
          Height = 96
          Color = 15138559
          DataField = 'mtexto2'
          DataSource = dsContenido
          TabOrder = 4
        end
        object cxLabel4: TcxLabel
          Left = 14
          Top = 40
          Caption = 'CUERPO'
        end
        object DBMemo5: TDBMemo
          Left = 65
          Top = 143
          Width = 521
          Height = 42
          Color = 15138559
          DataField = 'mtexto3'
          DataSource = dsContenido
          TabOrder = 6
        end
        object cxLabel5: TcxLabel
          Left = 2
          Top = 144
          Caption = 'NOTA FIN:'
        end
      end
    end
    object TabSheet7: TTabSheet
      Caption = 'OFICIO-T4'
      ImageIndex = 6
      object PnlOficio4: TPanel
        Left = 0
        Top = 0
        Width = 761
        Height = 191
        Align = alLeft
        TabOrder = 0
        object GroupBox6: TGroupBox
          Left = 11
          Top = 77
          Width = 353
          Height = 62
          Caption = 'ATENTAMENTE:'
          TabOrder = 0
          object DBAdvEdit16: TDBAdvEdit
            Left = 66
            Top = 15
            Width = 275
            Height = 21
            LabelCaption = 'NOMBRE'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 0
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'snombref1'
            DataSource = dsContenido
          end
          object DBAdvEdit17: TDBAdvEdit
            Left = 66
            Top = 38
            Width = 275
            Height = 21
            LabelCaption = 'CARGO'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 1
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'scargof1'
            DataSource = dsContenido
          end
        end
        object GroupBox7: TGroupBox
          Left = 370
          Top = 77
          Width = 353
          Height = 62
          Caption = 'RECIBI DE CONFORMIDAD:'
          TabOrder = 1
          object DBAdvEdit18: TDBAdvEdit
            Left = 66
            Top = 15
            Width = 275
            Height = 21
            LabelCaption = 'NOMBRE'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 0
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'snombref2'
            DataSource = dsContenido
          end
          object DBAdvEdit19: TDBAdvEdit
            Left = 66
            Top = 38
            Width = 275
            Height = 21
            LabelCaption = 'CARGO'
            LabelFont.Charset = DEFAULT_CHARSET
            LabelFont.Color = clWindowText
            LabelFont.Height = -11
            LabelFont.Name = 'Tahoma'
            LabelFont.Style = []
            Lookup.Separator = ';'
            Color = 15138559
            TabOrder = 1
            Text = 'DBAdvEdit1'
            Visible = True
            Version = '2.9.3.1'
            DataField = 'scargof2'
            DataSource = dsContenido
          end
        end
        object DBAdvEdit20: TDBAdvEdit
          Left = 65
          Top = 14
          Width = 521
          Height = 21
          LabelCaption = 'ASUNTO'
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          TabOrder = 2
          Text = 'DBAdvEdit1'
          Visible = True
          Version = '2.9.3.1'
          DataField = 'stexto1'
          DataSource = dsContenido
        end
      end
    end
    object TabSheet8: TTabSheet
      Caption = 'OFICIO-T5'
      ImageIndex = 7
      object PnlOficio5: TPanel
        Left = 0
        Top = 0
        Width = 185
        Height = 191
        Align = alLeft
        TabOrder = 0
      end
    end
    object TabSheet9: TTabSheet
      Caption = 'OFICIO-T6'
      ImageIndex = 8
      object PnlOficio6: TPanel
        Left = 0
        Top = 0
        Width = 185
        Height = 191
        Align = alLeft
        TabOrder = 0
      end
    end
    object TabSheet10: TTabSheet
      Caption = 'OFICIO-T7'
      ImageIndex = 9
      object PnlOficio7: TPanel
        Left = 0
        Top = 0
        Width = 185
        Height = 191
        Align = alLeft
        TabOrder = 0
      end
    end
  end
  object PnlSuperior: TPanel
    Left = 0
    Top = 0
    Width = 979
    Height = 213
    Align = alClient
    TabOrder = 1
    inline frmBarra1: TfrmBarra
      Left = 1
      Top = 1
      Width = 72
      Height = 211
      VertScrollBar.Style = ssHotTrack
      Align = alLeft
      Color = 7847370
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      ExplicitLeft = 1
      ExplicitTop = 1
      ExplicitWidth = 72
      ExplicitHeight = 211
      inherited AdvPanel1: TAdvPanel
        Width = 72
        Height = 211
        ParentShowHint = False
        ShowHint = True
        ExplicitWidth = 72
        ExplicitHeight = 211
        FullHeight = 0
        inherited btnRefresh: TAdvGlowButton
          OnClick = frmBarra1btnRefreshClick
        end
        inherited btnEdit: TAdvGlowButton
          OnClick = frmBarra1btnEditClick
        end
        inherited btnPost: TAdvGlowButton
          OnClick = frmBarra1btnPostClick
        end
        inherited btnCancel: TAdvGlowButton
          Hint = 'Cancelar cambios (F11)'
          OnClick = frmBarra1btnCancelClick
        end
        inherited btnDelete: TAdvGlowButton
          OnClick = frmBarra1btnDeleteClick
        end
        inherited btnPrinter: TAdvGlowButton
          Enabled = False
        end
        inherited btnExit: TAdvGlowButton
          OnClick = frmBarra1btnExitClick
        end
        inherited btnAdd: TAdvGlowButton
          OnClick = frmBarra1btnAddClick
        end
      end
      inherited ImgBtns: TImageList
        Bitmap = {
          494C0101080009004C0010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
          0000000000003600000028000000400000003000000001002000000000000030
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000000000000000000000000000000000008F8F8F001C1C
          1C00262626002626260026262600262626002626260026262600262626002626
          26001C1C1C009F9F9F0000000000000000000000000000000000000000000000
          000000000000F4F7FB00A0B6D900000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000EFF0FC00303E
          D600CFD2F500000000000000000000000000000000000000000000000000CFCF
          F4003030CE00EFEFFB00000000000000000000000000303030001C1C1C005757
          5700000000000000000000000000F9F9F900F2F2F200E5E5E500DFDFDF00D8D8
          D800555555001C1C1C0030303000000000000000000000000000000000000000
          000000000000F4F4F6002F69BC001B7DF000165EC000D5DCEB00000000000000
          00000000000000000000000000000000000053575500B3B9B700BAC0BE00BDC3
          C100BEC4C200BFC5C300BFC5C300BFC5C300BFC5C300BFC5C300BDC2C0000000
          00000000000000000000000000000000000000000000EFF0FC003042D600144B
          EC00101ECD00CFD2F50000000000000000000000000000000000CFD0F4001010
          C7000000E9003030CE00EFEFFB0000000000BFBFBF0094949400B7B7B700CACA
          CA00F6F6F600F6F6F600F6F6F600F5F5F500EFEFEF00E3E3E300DDDDDD00D6D6
          D600CACACA00AAAAAA0094949400BFBFBF000000000000000000000000000000
          000085A0CF00306DC100166BD200228AFF00238CFF001761C500648AC500E2E8
          F4000000000000000000000000000000000053575500AEAEAE00B0B1B100AFB5
          B300CACFCE00D1D5D400D3D7D500D3D7D600D3D7D600D3D7D600CFD4D2000000
          000000000000000000000000000000000000000000003045DA001C59ED003371
          FE001246EB00101ECD00CFD2F5000000000000000000CFD0F4001013C8000000
          E9000000FE000000E9003030CE000000000050505000B8B8B8008A8A8A000000
          00008E8E8E0096969600969696009696960095959500909090008D8D8D008A8A
          8A000000000082828200B8B8B80060606000000000000000000000000000CDD6
          E9001A7DEB00228FFF00248FFF00278FFF001E8DFF0078BAFF00245DB200CFD7
          EB000000000000000000000000000000000053575500ABABAB00ABABAB00ABAB
          AB00999A9900AFB1B100DADDDC00E4E6E600E5E8E700E5E8E700E0E2E2000000
          00000000000000000000000000000000000000000000CFD5F700102BD4001E5E
          EE003371FE002B68FE001246EB00CFD2F500CFD1F500031CEA000219FE00000A
          FE000000E9001010C700CFCFF4000000000020202000CCCCCC00CCCCCC00CCCC
          CC00CCCCCC00CCCCCC00CCCCCC00CCCCCC00CCCCCC00CCCCCC00CCCCCC00CCCC
          CC00CCCCCC00CCCCCC00CCCCCC002020200000000000F4F5F900346DBD00218E
          FD0061B1FF00549DEF002668C000B4E2FF0096C7F700D6DAEC00000000000000
          00000000000000000000000000000000000053575500ABABAB00AAAAAA00A4A4
          A4009C9C9C009B9B9B0096969600C5C6C600F1F2F200F2F3F200DFE0DF00E7E7
          E7000A0AA900E8E8E800F7F7F700000000000000000000000000CFD5F700102B
          D4003B7AFE003371FE002B68FE00101ECD00101DCC000B34FE000628FE000219
          FE001013C800CFD0F40000000000000000001F1F1F00D6D6D600D6D6D600D6D6
          D600D9D9D900DBDBDB00DCDCDC00DCDCDC00DCDCDC00DCDCDC00DBDBDB00D9D9
          D900D6D6D600D6D6D600D6D6D6001F1F1F00000000008AA4D000197CE800369E
          FF00337ED7006488C400B9BEDC0094C7F800396EBA000000000000000000F4F7
          FB00D3DAEA00D1DAEA00EBEFF7000000000053575500ABABAB00AAAAAA00A1A1
          A100989898009090900096969600C5C5C500F7F7F700E7E7E700D0D1D1000A0A
          A8003535D300CCCCCC00DDDDDD00EFEFEF00000000000000000000000000CFD5
          F7001E5EEE003B7AFE003371FE00144AEE001042EE00103FFE000B34FE00031C
          EA00CFD0F4000000000000000000000000001D1D1D00E5E5E500E9E9E900ECEC
          EC00EDEDED00EDEDED00EDEDED00EDEDED00EDEDED00EDEDED00EDEDED00EDED
          ED00EDEDED00E9E9E900E5E5E5001D1D1D00000000002F6CBF002898FF0064B9
          FF006F8DC60000000000FBFBFC003B71BC00E9EBF4000000000000000000B5C5
          E1001565D1000D5BCD006188C6000000000053575500ABABAB00A8A8A8009F9F
          9F00959595008A8A8A0091919100EFF0F000FCFCFC00EBEBEB000909A8003030
          D6003A3AD8004242D9004646DA000303A5000000000000000000000000000000
          0000CFD5F700102BD4002464F1003371FE002B68FE001042EE00101DCC00CFD1
          F5000000000000000000000000000000000050505000D5EDD50032A13200B5D2
          B500FCFCFC00FCFCFC00FCFCFC00FCFCFC00FCFCFC00FCFCFC00FCFCFC00FCFC
          FC00FCFCFC00FCFCFC00FCFCFC005050500000000000177BE7004FB7FF0078BB
          F60000000000000000000000000000000000000000000000000000000000D4DA
          EA00258FFF001D7EFF004477C4000000000053575500ABABAB00A5A5A5009C9C
          9C0093939300BFBFBF0085858500EEEFEF00FCFCFC000C0CAB002D2DD5000000
          CC000000CC000000CC000000CC000303A6000000000000000000000000000000
          0000CFD5F700102ED600296BF1003B7AFE003371FE00144AEE00101ECD00CFD2
          F500000000000000000000000000000000009F9F9F00E5ECE500D7EBD700EDF4
          ED00000000000000000000000000000000000000000000000000000000000000
          00000000000000000000F5F5F5009F9F9F00000000001882EF0055C0FF007FB8
          ED0000000000000000000000000000000000000000000000000000000000AEBE
          DC00238BFF001A7AFC005782C5000000000053575500ABABAB00A3A3A3009A9A
          9A009191910094949400A9A9A900EEEFEF00E5E5F4004040CB003737D7003A3A
          D8004040D9003939D7003D3DD8000303A600000000000000000000000000CFD8
          F800347CF00065A0FE005E9AFE00296BF1002464F1003371FE002B68FE001246
          EB00CFD2F50000000000000000000000000000000000BFBFBF00505050003F3F
          3F00000000000000000000000000F9F9F900F1F1F100E3E3E300DBDBDB00D4D4
          D4003838380050505000BFBFBF000000000000000000376FBA005184C5003E75
          BE00000000000000000000000000497EC000177BD20000000000869FCD001B7E
          DA003F93FF001D61C500D3DAEA000000000053575500AAAAAA00A0A0A0009797
          97008E8E8E0076767600BBBBBB00EEEEEE00FCFCFC00E6E7F5005151CD005656
          DE005858DD005B5BDE005555DD000303A6000000000000000000CFD9F900103E
          E00077ADFE006FA7FE0065A0FE00102ED600102BD4003B7AFE003371FE002B68
          FE00101ECD00CFD2F50000000000000000000000000000000000EFEFEF001919
          1900636363007B7B7B0078787800727272006F6F6F0068686800656565004949
          490019191900EFEFEF00000000000000000000000000F8F8FA00F7F7F900F7F7
          FA000000000000000000E8E9F40030B0F4002EACF200618EC9001D83DB002497
          FF005197F6005A83C300000000000000000053575500A8A8A8009E9E9E009595
          95008C8C8C00AAAAAA008E8E8E00EDEEEE00FCFCFC00FCFCFC00E6E7F3005959
          CF00A5A5EC000E0EAB000E0EAB002929B40000000000CFDAFA001044E400448F
          F3007FB3FE0077ADFE00347CF000CFD5F700CFD5F7001E5EEE003B7AFE003371
          FE001246EB00101ECD00CFD2F500000000000000000000000000000000009F9F
          9F00C5C5C500FAFAFA00F4F4F400E8E8E800E2E2E200D5D5D500CFCFCF007979
          79009F9F9F000000000000000000000000000000000000000000000000000000
          000000000000DFE2F0003473BC003BC7FF0033B3F80029A4FB00279CFF004C9F
          FF003068BD00DEE2EF00000000000000000053575500A6A6A6009B9B9B009292
          9200898989007F7F7F0083838300EDEEEE00FCFCFC00FCFCFC00F8F9F900F1F1
          FA005858CC00000000000000000000000000000000003066EC004F9CF600A1CB
          FE00448FF300103EE000CFD8F8000000000000000000CFD5F700102BD4001E5E
          EE003371FE00144BEC00303ED600000000000000000000000000000000000000
          0000BCBCBC000000000000000000F4F4F400EEEEEE00E2E2E200DBDBDB007272
          7200000000000000000000000000000000000000000000000000000000000000
          00002A61B10031B3F40046DBFF0034C2FF005FC8FF008ECBFE00659DE3004374
          BD000000000000000000000000000000000053575500A3A3A300979797008989
          890080808000767676007B7B7B00EBECEC00FCFCFC00FCFCFC00F8F9F9000000
          0000F5F5FB0000000000000000000000000000000000EFF3FE003066EC004F9C
          F6001044E400CFD9F90000000000000000000000000000000000CFD5F700102B
          D4001C59ED003042D600EFF0FC00000000000000000000000000000000000000
          0000B7B7B7000000000000000000FAFAFA00F4F4F400E8E8E800E2E2E2006F6F
          6F00000000000000000000000000000000000000000000000000000000000000
          0000C0CFE7005283C300247DC70065D7FF0090E0FF004C79BD009EB0D700F6F6
          F90000000000000000000000000000000000535755009A9A9A008A8A8A007E7E
          7E00747474006C6C6C0072727200E8E9E900FBFCFC00FCFCFC00F8F8F8000000
          0000000000000000000000000000000000000000000000000000EFF3FE003066
          EC00CFDAFA00000000000000000000000000000000000000000000000000CFD5
          F7003045DA00EFF0FC0000000000000000000000000000000000000000000000
          0000B2B2B200000000000000000000000000FAFAFA00EEEEEE00E8E8E8006C6C
          6C00000000000000000000000000000000000000000000000000000000000000
          00000000000000000000C8CFE600367DC3008FCEF900C8CCE500000000000000
          0000000000000000000000000000000000005357550090909000818181007676
          76006B6B6B006363630067676700E1E3E300F7F8F800F8F9F900F5F6F6000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000181818001818180018181800181818001818180018181800181818001212
          1200000000000000000000000000000000000000000000000000000000000000
          00000000000000000000000000000000000000000000F3F6FA00000000000000
          000000000000000000000000000000000000A8AAA90053575500535755005357
          55005357550053575500565A5800676A680074777600757977006D706E000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000084777300847773008477
          7300847773008477730084777300847773008477730084777300847773008477
          7300FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000ECEAEAFFE2DDDCFFE3DFDFFF00000000000000000000
          000000000000000000000000000000000000DED3CF00FAF7F500FAF7F500FAF7
          F500EDE9E700ECE7E500ECE7E500E8E0DE00F1EAE800F1E9E700F0E8E600EEE6
          E300FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000534B4AFF262626FF1D1D1DFFE4E0DFFF000000000000
          000000000000000000000000000000000000DED3CF00FBF8F700FBF8F700FBF8
          F7005E70A800E2DDDB00E1DBD900E0D8D600E0D8D600EAE2E000F1E9E700F0E7
          E500FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000EFF5EF00005A
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000BAD0E200DAE9F60000000000000000000000000000000000000000000000
          000000000000000000007F7978FF565656FF303030FFE4DFDEFF000000000000
          000000000000000000000000000000000000DED3CF00FBF8F700FBF8F700FBF8
          F700818BA2006680C000E1DBD900E0D8D600E0D8D600DCD3D100DED5D300E6DD
          DB00FFFFFF00FFFFFF00FFFFFF00FFFFFF0000000000000000009FC19F00006C
          0000609860000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000EAF0F400186D
          B4002399F7002194F400B5D3ED00000000000000000000000000000000000000
          00000000000000000000595453FF383838FF404040FFE4DFDEFF000000000000
          000000000000000000000000000000000000DED3CF00FBF8F700FBF8F700FBF8
          F700C9E7FE00A5E1FE0088C9F800E0D8D600E0D8D600DCD3D100DBD2D000DAD0
          CF00E8E6E600E8E6E600EDECEC00FFFFFF0000000000000000001065100000BA
          000000C0000000760000307A3000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000000000000000000000000000E6ECF200356B99001052
          8B00218EE5002297F70053A1E100000000000000000000000000000000000000
          00000000000000000000837E7CFF676767FF3E3E3EFFE4DFDEFF000000000000
          000000000000000000000000000000000000DED3CF00FBF8F700FBF8F700FBF8
          F70076E2FE0059D4FE0060CBFE00AFB0CD00E0D9D700DBD2D000DAD1CF00D8CF
          CD00E8E6E600E8E6E600E8E6E600EAE9E90000000000AFCDAF000075000000D1
          000000D1000000D10000008A0000EFF5EF000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000B2C6D700104C
          800012558E00259BF800249AF800000000000000000000000000FCFCFCFFFCFB
          FBFFFCFBFBFFFCFBFBFF747171FF606060FF565656FFD1CBCAFFFAFAFAFFFAFA
          FAFFFAFAFAFFFDFDFDFF0000000000000000DED3CF00FAF8F600FAF8F600FAF8
          F60061EFFE0050D1FE004DCBFE00688BD200F4EDEB00E0D8D600DAD0CF00DAD0
          CF00E8E6E600E8E6E600E8E6E600E8E6E6000000000070AA700000A0000000DA
          000000DA000000DA000000DA000010711000CFE3CF0000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000FBFCFD00104C
          8000104D8300279DF800269BF8000000000000000000CAB5B2FF242424FF1B1B
          1BFF202020FF1F1F1FFF555555FF626262FF686868FF201F1FFF1F1F1FFF2020
          20FF262626FF3B3635FFFBFBFBFF00000000DED3CF00FFFFFF00FFFFFF00FFFF
          FF0038C0F90040E7FE0036C4FE0044A4FE005C77C200EDE4E100EDE4E100EBE2
          DF00F5F4F400EAE9E900E8E6E600FAFAFA0000000000007B000000EC000000EC
          000000D6000000E5000000EC000000EC000000BF00009FCB9F00000000000000
          000000000000000000000000000000000000000000003869930014578D003CB3
          FA003AB1F90035A1E700429FDE0000000000000000000000000000000000104C
          8000115289002AA0F800299FF8000000000000000000C2B3B0FF7F7F7FFF6B6B
          6BFF676767FF797979FF646464FF666666FF696969FF6F6F6FFF7F7F7FFF7979
          79FF7F7F7FFF525252FFF5F4F4FF00000000DED3CF00FFFFFF00FFFFFF00FFFF
          FF009FD2F9002EE2FD002ECFFE0035A4FE002971DC00ECE3E000EBE2DF00E9DF
          DC00FFFFFF00FFFFFF00F8F7F700FFFFFF000000000000A8000000F6000000F6
          0000108210000081000000B7000000F6000000F60000008200009FCD9F000000
          000000000000000000000000000000000000000000000F4A7C00155A91003DB4
          FA003CB3FA0039ADF5007FB5DD0000000000000000000000000000000000104C
          8000196AAA002CA2F8002BA1F8000000000000000000EBD9D6FFD2D2D2FFD2D2
          D2FFCACACAFFC4C4C4FFACACACFF9E9E9EFFB1B1B1FFCCCCCCFFC5C5C5FFC7C7
          C7FFD3D3D3FFC1C0C0FFFAF9F9FF00000000DED3CF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF007FC5F70019E7FE0017B4FE002392FE00E3D9D500E3D7D400E2D6
          D300FFFFFF00FFFFFF00FFFFFF00FFFFFF00EFF8EF002AF72A0017D31700108F
          100000000000000000000000000060B46000028E020027F827002AF72A000B9D
          0B0000000000000000000000000000000000000000000F4A7C001860960040B7
          FA003FB6FA003DB4FA0045A7E4000000000000000000000000008AA8C2002689
          D00031A8F9002FA5F900399AE3000000000000000000FCFBFBFFEAE8E7FFEAE8
          E8FFEAE8E8FFEAE8E8FFC2BDBCFFB0B0B0FF999999FFD1CBCAFFEBE9E9FFEBE9
          E9FFEBE9E9FFF4F3F3FF0000000000000000DED3CF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00EEF6FC0008B6F70011B0FE00149EFE008992BF00DCD0CC00DBCF
          CA00FFFFFF00FFFFFF00FFFFFF00FFFFFF00AFDCAF0035D5350010941000CFE9
          CF0000000000000000000000000000000000DFF0DF000594050028C728004FF1
          4F0040AA4000EFF8EF000000000000000000000000000F4A7C001A649B003CAE
          EF0040B7F9003EB6FA003DB4FA005EABDF009EC9E7009EC3DF002376B40034AB
          F90033AAF90031A7F80090C0E400000000000000000000000000000000000000
          00000000000000000000C3BDBCFF9D9D9DFFB3B3B3FFE4DFDEFF000000000000
          000000000000000000000000000000000000DED3CF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FEFEFD005EB0F2000BB1FE000AAAFE001F51C400DACDC900D9CC
          C800FFFFFF00FFFFFF00FFFFFF00FFFFFF0060BC6000109B1000CFEACF000000
          00000000000000000000000000000000000000000000DFF1DF0060BC6000089B
          08002DB52D0030A83000EFF8EF0000000000000000000F4A7C00276799001157
          9300125894003FB5F60040B7FA003EB5FA003DB4FA003BB2FA0039B0FA0037AE
          F90036ADF900DAEAF50000000000000000000000000000000000000000000000
          00000000000000000000D4CFCEFFCDCDCDFFB1B1B1FFE4DFDEFF000000000000
          000000000000000000000000000000000000DED3CF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FDFCFB00FBF9F80000DFFE0002A6FE000E73FE003F6CD200C8B6
          B100FFFFFF00FFFFFF00FFFFFF00FFFFFF0040B8410000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000DFF3DF0060C4600000A10300CFEED000000000006F92B000E6ECF2002B6B
          A100125A98001969A800339BDC003FB7FA003EB6FA003CB3FA003BB2FA0038AD
          F60056A5DC000000000000000000000000000000000000000000000000000000
          00000000000000000000E4E0DFFFDDDDDDFFBEBEBEFFE4DFDEFF000000000000
          000000000000000000000000000000000000DED3CF00FFFFFF00FFFFFF00FFFF
          FF00FFFFFF00FDFCFB00FBF8F70000C5F80000BEFE000587FE000959E400BFC7
          E600FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000DFF4E10010B326000000000000000000000000000000
          0000000000007AA5CA003879B1001461A3001964A50084ADD000DDEAF4000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000E2D3D0FFF6E2DDFFF3DEDAFFFCFCFBFF000000000000
          000000000000000000000000000000000000DED3CF00FFFFFF00FFFFFF00FFFF
          FF00FEFEFD00FBF9F800F9F5F30037ADE80000DBFE000AA0FE002E8CFE006584
          D100FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000F3F7FA00F7F9FC0000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000000DDD1CD00DED3CF00DED3CF00DED3
          CF00DED3CF00DACEC900D5C8C300C8B6B1002088DF004D77CA009FAAD900FFFF
          FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000000424D3E000000000000003E000000
          2800000040000000300000000100010000000000800100000000000000000000
          000000000000000000000000FFFFFF0000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000000000000000FFFFC003F9FFFFFFC7E38E01F83F001F
          83C10000F00F001F81810000E00F001F80010000803F0001C003000080610000
          E007000084610000F00F00008FE10000F00F0FFC8FE10000E0078E018E410000
          C003C0038C0300008001E007F80300078181F60FF00F001783C1F60FF00F001F
          C7E3F70FFC3F001FFFFFF00FFFBF001FFFFF8000FFFFFFFFFC7F0000FFFFFFFF
          FC3F0000CFFFFFF3FC3F0000C7FFFFC1FC3F0000C1FFFF81FC3F000080FFFFC1
          C0030000807FFFC180010000803F81E180010000801F81E1800100000E0F81C1
          800300000F038001FC3F00001F818003FC3F00007FF08007FC3F0000FFFCF81F
          FC3F0000FFFFFE7FFFFF0000FFFFFFFF00000000000000000000000000000000
          000000000000}
      end
    end
    object Panel1: TPanel
      Left = 73
      Top = 1
      Width = 905
      Height = 211
      Align = alClient
      Caption = 'Panel1'
      TabOrder = 1
      object DBGrid1: TDBGrid
        Left = 1
        Top = 1
        Width = 903
        Height = 209
        Align = alClient
        Color = 15138559
        DataSource = dsContenido
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        Columns = <
          item
            Expanded = False
            FieldName = 'iorden'
            Title.Caption = 'ORDEN'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'snombreportada'
            Title.Caption = 'NOMBRE'
            Width = 249
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ltipo'
            Title.Caption = 'TIPO'
            Width = 103
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'lincluirindice'
            Title.Caption = 'EN INDICE'
            Visible = True
          end>
      end
    end
  end
  object dsContenido: TDataSource
    DataSet = zqContenido
    Left = 728
    Top = 8
  end
  object zqContenido: TZQuery
    Connection = connection.zConnection
    AfterOpen = zqContenidoAfterOpen
    AfterClose = zqContenidoAfterClose
    AfterScroll = zqContenidoAfterScroll
    AfterInsert = zqContenidoAfterInsert
    AfterEdit = zqContenidoAfterEdit
    AfterPost = zqContenidoAfterPost
    AfterCancel = zqContenidoAfterCancel
    SQL.Strings = (
      'select * from contenidonotacampo order by iOrden')
    Params = <>
    Left = 728
    Top = 40
  end
  object OpenPicture: TOpenPictureDialog
    Filter = 'JPEG Image File (*.jpg)|*.jpg'
    Left = 728
    Top = 80
  end
end
