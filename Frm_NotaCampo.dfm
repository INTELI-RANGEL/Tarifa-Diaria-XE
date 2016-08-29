object FrmNotaCampo: TFrmNotaCampo
  Left = 0
  Top = 0
  Caption = 'Acta de Entrega'
  ClientHeight = 592
  ClientWidth = 1151
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Scaled = False
  Visible = True
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object CxPage1: TcxPageControl
    Left = 0
    Top = 0
    Width = 1151
    Height = 592
    Align = alClient
    TabOrder = 0
    Properties.ActivePage = cTs1
    Properties.CustomButtons.Buttons = <>
    ClientRectBottom = 588
    ClientRectLeft = 4
    ClientRectRight = 1147
    ClientRectTop = 24
    object cTs1: TcxTabSheet
      Caption = 'cTs1'
      ImageIndex = 0
      object Spl1: TcxSplitter
        Left = 0
        Top = 213
        Width = 1143
        Height = 7
        AlignSplitter = salBottom
      end
      object GBx1: TcxGroupBox
        Left = 0
        Top = 0
        Align = alClient
        PanelStyle.Active = True
        TabOrder = 1
        Height = 213
        Width = 1143
        object cxGrid1: TcxGrid
          Left = 73
          Top = 2
          Width = 1068
          Height = 209
          Align = alClient
          PopupMenu = pmActa
          TabOrder = 0
          object cxGrid1DBTableView1: TcxGridDBTableView
            Navigator.Buttons.CustomButtons = <>
            DataController.DataSource = dsActa
            DataController.Summary.DefaultGroupSummaryItems = <>
            DataController.Summary.FooterSummaryItems = <>
            DataController.Summary.SummaryGroups = <>
            FilterRow.Visible = True
            OptionsData.Deleting = False
            OptionsData.DeletingConfirmation = False
            OptionsData.Editing = False
            OptionsData.Inserting = False
            OptionsView.ColumnAutoWidth = True
            OptionsView.GroupByBox = False
            Styles.Background = cxStyle1
            object cxGrid1DBTableView1Column1: TcxGridDBColumn
              Caption = 'FOLIO'
              DataBinding.FieldName = 'snumeroorden'
              Width = 155
            end
            object cxGrid1DBTableView1Column2: TcxGridDBColumn
              Caption = 'No. DE ACTA'
              DataBinding.FieldName = 'sNoActa'
              Width = 215
            end
            object cxGrid1DBTableView1Column3: TcxGridDBColumn
              Caption = 'FECHA DE ACTA'
              DataBinding.FieldName = 'dFecha'
              Width = 140
            end
          end
          object cxGrid1Level1: TcxGridLevel
            GridView = cxGrid1DBTableView1
          end
        end
        inline BrPrincipal: TfrmBarra
          Left = 2
          Top = 2
          Width = 71
          Height = 209
          VertScrollBar.Style = ssHotTrack
          Align = alLeft
          TabOrder = 1
          ExplicitLeft = 2
          ExplicitTop = 2
          ExplicitWidth = 71
          ExplicitHeight = 209
          inherited AdvPanel1: TAdvPanel
            Width = 71
            Height = 209
            ExplicitWidth = 71
            ExplicitHeight = 209
            FullHeight = 0
            inherited btnRefresh: TAdvGlowButton
              OnClick = BrPrincipalbtnRefreshClick
            end
            inherited btnEdit: TAdvGlowButton
              OnClick = BrPrincipalbtnEditClick
            end
            inherited btnPost: TAdvGlowButton
              OnClick = BrPrincipalbtnPostClick
            end
            inherited btnCancel: TAdvGlowButton
              OnClick = BrPrincipalbtnCancelClick
            end
            inherited btnDelete: TAdvGlowButton
              OnClick = BrPrincipalbtnDeleteClick
            end
            inherited btnPrinter: TAdvGlowButton
              OnClick = BrPrincipalbtnPrinterClick
            end
            inherited btnExit: TAdvGlowButton
              OnClick = BrPrincipalbtnExitClick
            end
            inherited btnAdd: TAdvGlowButton
              OnClick = BrPrincipalbtnAddClick
            end
          end
        end
      end
      object CxPageDetalle: TcxPageControl
        Left = 0
        Top = 220
        Width = 1143
        Height = 344
        Align = alBottom
        TabOrder = 2
        Properties.ActivePage = cTsCaratula
        Properties.CustomButtons.Buttons = <>
        OnChange = CxPageDetalleChange
        OnPageChanging = CxPageDetallePageChanging
        ClientRectBottom = 340
        ClientRectLeft = 4
        ClientRectRight = 1139
        ClientRectTop = 24
        object cTsCaratula: TcxTabSheet
          Caption = 'Informaci'#243'n General'
          ImageIndex = 0
          object GBx2: TcxGroupBox
            Left = 0
            Top = -3
            Align = alBottom
            TabOrder = 0
            Height = 319
            Width = 1135
            object dxLayoutControl1: TdxLayoutControl
              Left = 2
              Top = 5
              Width = 1131
              Height = 312
              Align = alClient
              TabOrder = 0
              object DbMmObservacion: TcxDBMemo
                Left = 108
                Top = 118
                DataBinding.DataField = 'mObservaciones'
                DataBinding.DataSource = dsActa
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                TabOrder = 7
                OnEnter = DbMmObservacionEnter
                OnExit = DbMmObservacionExit
                Height = 72
                Width = 849
              end
              object DbRdGrpTipo: TcxDBRadioGroup
                Left = 772
                Top = 37
                Caption = 'Tipo:'
                DataBinding.DataField = 'eTipo'
                DataBinding.DataSource = dsActa
                Properties.Columns = 2
                Properties.Items = <
                  item
                    Caption = 'Acta Parcial'
                    Value = 'Parcial'
                  end
                  item
                    Caption = 'Acta Total'
                    Value = 'Total'
                  end>
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                TabOrder = 6
                Height = 43
                Width = 185
              end
              object DbTxtEdtActa: TcxDBTextEdit
                Left = 108
                Top = 37
                DataBinding.DataField = 'sNoActa'
                DataBinding.DataSource = dsActa
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                TabOrder = 1
                OnEnter = DbTxtEdtActaEnter
                OnExit = DbTxtEdtActaExit
                OnKeyPress = DbTxtEdtActaKeyPress
                Width = 658
              end
              object DbLkpCmbFolio: TcxDBLookupComboBox
                Left = 108
                Top = 10
                DataBinding.DataField = 'sNumeroOrden'
                DataBinding.DataSource = dsActa
                Properties.DropDownListStyle = lsFixedList
                Properties.KeyFieldNames = 'sNumeroOrden'
                Properties.ListColumns = <
                  item
                    Caption = 'Folio'
                    HeaderAlignment = taCenter
                    MinWidth = 150
                    FieldName = 'sIdFolio'
                  end
                  item
                    Caption = 'Id'
                    HeaderAlignment = taCenter
                    Width = 100
                    FieldName = 'sNumeroOrden'
                  end>
                Properties.ListOptions.SyncMode = True
                Properties.ListSource = dsFolios
                Properties.OnCloseUp = DbLkpCmbFolioPropertiesCloseUp
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                Style.ButtonStyle = bts3D
                Style.PopupBorderStyle = epbsFrame3D
                TabOrder = 0
                OnEnter = DbLkpCmbFolioEnter
                OnExit = DbLkpCmbFolioExit
                OnKeyPress = DbLkpCmbFolioKeyPress
                Width = 658
              end
              object DbTxtEdtEspecialidad: TcxDBTextEdit
                Left = 108
                Top = 64
                DataBinding.DataField = 'sEspecialidad'
                DataBinding.DataSource = dsActa
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                TabOrder = 2
                OnEnter = DbTxtEdtEspecialidadEnter
                OnExit = DbTxtEdtEspecialidadExit
                OnKeyPress = DbTxtEdtEspecialidadKeyPress
                Width = 658
              end
              object DbDtEdtFecha: TcxDBDateEdit
                Left = 852
                Top = 10
                DataBinding.DataField = 'dFecha'
                DataBinding.DataSource = dsActa
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                Style.ButtonStyle = bts3D
                Style.PopupBorderStyle = epbsFrame3D
                TabOrder = 5
                OnEnter = DbDtEdtFechaEnter
                OnExit = DbDtEdtFechaExit
                OnKeyPress = DbDtEdtFechaKeyPress
                Width = 105
              end
              object GBx5: TcxGroupBox
                Left = 963
                Top = 10
                Caption = 'Opciones de Impresi'#243'n'
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                TabOrder = 10
                Height = 292
                Width = 158
                object dxLayoutControl2: TdxLayoutControl
                  Left = 2
                  Top = 18
                  Width = 154
                  Height = 272
                  Align = alClient
                  TabOrder = 0
                  object DbChkBxPernocta: TcxDBCheckBox
                    Left = 10
                    Top = 10
                    Caption = 'Aplica Pernocta'
                    DataBinding.DataField = 'lPernocta'
                    DataBinding.DataSource = dsActa
                    Properties.ValueChecked = 'Si'
                    Properties.ValueUnchecked = 'No'
                    Style.BorderColor = clWindowFrame
                    Style.BorderStyle = ebs3D
                    Style.HotTrack = False
                    TabOrder = 0
                    Width = 121
                  end
                  object DbChkBxMaterial: TcxDBCheckBox
                    Left = 10
                    Top = 37
                    Caption = 'Aplica Material'
                    DataBinding.DataField = 'lMaterial'
                    DataBinding.DataSource = dsActa
                    Properties.ValueChecked = 'Si'
                    Properties.ValueUnchecked = 'No'
                    Style.BorderColor = clWindowFrame
                    Style.BorderStyle = ebs3D
                    Style.HotTrack = False
                    TabOrder = 1
                    Width = 121
                  end
                  object DbChkBxPaginas: TcxDBCheckBox
                    Left = 10
                    Top = 64
                    Caption = 'No Pags. Por Secci'#243'n'
                    DataBinding.DataField = 'lPaginado'
                    DataBinding.DataSource = dsActa
                    Properties.ValueChecked = 'Si'
                    Properties.ValueUnchecked = 'No'
                    Style.BorderColor = clWindowFrame
                    Style.BorderStyle = ebs3D
                    Style.HotTrack = False
                    TabOrder = 2
                    Width = 121
                  end
                  object DbChkBxPdas: TcxDBCheckBox
                    Left = 10
                    Top = 91
                    Caption = 'Folio en Partidas'
                    DataBinding.DataField = 'lPartidas'
                    DataBinding.DataSource = dsActa
                    Properties.ValueChecked = 'Si'
                    Properties.ValueUnchecked = 'No'
                    Style.BorderColor = clWindowFrame
                    Style.BorderStyle = ebs3D
                    Style.HotTrack = False
                    TabOrder = 3
                    Width = 133
                  end
                  object btnAjustar: TcxButton
                    Left = 10
                    Top = 118
                    Width = 134
                    Height = 43
                    Caption = 'Ajustar Acta'
                    OptionsImage.Glyph.Data = {
                      36100000424D3610000000000000360000002800000020000000200000000100
                      2000000000000010000000000000000000000000000000000000000000000000
                      0000000000000000000000000000000000000000000000000000000000000000
                      0001000000010000000100000001000000010000000100000001000000010000
                      0001000000010000000100000001000000010000000100000001000000010000
                      0001000000010000000100000001000000010000000000000000000000000000
                      0000000000000000000000000000000000000000000000000001000000020000
                      0003000000030000000300000003000000030000000300000003000000030000
                      0003000000040000000400000004000000040000000400000004000000040000
                      0004000000040000000400000004000000030000000300000001000000000000
                      0000000000000000000000000000000000000000000100000003000000080000
                      000B0000000D0000000D0000000E0000000E0000000E0000000E0000000E0000
                      000E0000000F0000000F0000000F0000000F0000000F0000000F0000000F0000
                      0010000000100000000F0000000F0000000E0000000900000004000000000000
                      00000000000000000000000000000000000000000002000000087F5E54BCB084
                      74FFB18473FFB08373FFB08272FFAF8172FFAE8170FFAE7F71FFAE7F70FFAE7E
                      6FFFAD7E6EFFAD7D6DFFAD7C6EFFAB7C6DFFAC7B6CFFAC7B6CFFA97B6BFFA97A
                      6BFFA9796AFFA97969FFA87969FFA87869FF79554BBD0000000A000000000000
                      000000000000000000000000000000000000000000020000000AB38777FFFBF8
                      F5FFFBF7F5FFFBF7F5FFFBF7F4FFFBF7F4FFFBF7F4FFFAF7F4FFFAF6F4FFFAF6
                      F3FFFAF6F3FFFAF6F2FFFAF5F2FFFAF4F2FFF9F4F1FFFAF4F1FFF9F4F1FFFAF4
                      F1FFF9F4F0FFF9F3F0FFF9F3F0FFF9F3F0FFA97969FF0000000D000000000000
                      000000000000000000000000000000000000000000020000000AB5897AFFFCF8
                      F7FFF6EEE8FFF6EDE7FFF6EDE7FFF6EDE7FFF6EDE7FFF6EDE7FFF5ECE6FFF5ED
                      E6FFF6ECE6FFF6EDE6FFF6EDE6FFF6ECE5FFF6ECE6FFF5ECE6FFF5ECE5FFF5EC
                      E5FFF5EAE5FFF5EAE5FFF5EAE5FFFAF4F1FFAB7B6CFF0000000E000000000000
                      000000000000000000000000000000000000000000020000000AB68B7CFFFCFA
                      F8FFF7EEE9FFF7EFE9FFF6EFE9FFF6EEE8FFF6EEE8FFF6EEE8FFF6EEE8FFF6ED
                      E7FFF6EEE8FFF6EDE7FFF6EDE7FFF6EDE7FFF5EDE7FFF6EDE6FFF6EDE7FFF6EC
                      E6FFF5ECE6FFF5ECE6FFF5ECE6FFFBF6F2FFAC7D6EFF0000000E000000000000
                      0000000000000000000000000000000000000000000200000009B88D7EFFFDFA
                      F9FFF7EFEAFFF6EFEAFFF7EFEAFFFAF6F4FFFAF6F4FFFBF7F4FFFBF6F4FFFAF6
                      F4FFFAF6F3FFFAF6F3FFFAF6F3FFFAF6F3FFFAF6F3FFFAF6F3FFFAF6F3FFFAF6
                      F3FFF6EDE7FFF6EDE7FFF5ECE6FFFBF6F4FFAE7F70FF0000000D000000000000
                      0000000000000000000000000000000000000000000200000009BA9081FFFDFB
                      FAFFF7F0EDFFF7F0ECFFF8F0ECFFB38573FFB38573FFB38472FFB18472FFB184
                      72FFB18372FFB18472FFB18471FFB18371FFB08371FFB08371FFB18270FFA06D
                      5CFFF6EEE9FFF6EEE8FFF6EEE8FFFBF8F5FFAF8273FF0000000C000000000000
                      0001000000010000000100000001000000010000000300000009BC9284FFFDFC
                      FBFFF8F1EEFFF8F0EEFFF8F0EDFFF1E6E0FFF0E5E0FFF0E5E0FFF0E5E1FFF1E5
                      E0FFF1E5E0FFF0E5DEFFF0E4E0FFF0E4DEFFF0E5DEFFF0E4DEFFF0E4DEFFEDE2
                      DBFFF6EEE9FFF7EEE9FFF7EFE9FFFCF9F6FFB18474FF0000000C000000020000
                      000300000003000000030000000300000003000000050000000BBD9487FFFEFC
                      FCFFF9F2EFFFF8F1EEFFF8F2EEFFB58876FFB58875FFB48775FFB48775FFB486
                      75FFB48674FFB48674FFB48574FFB38674FFB38573FFB38573FFB18472FFA26F
                      5EFFF7EFECFFF7F0ECFFF7EFEAFFFCFAF8FFB28777FF0000000B000000080000
                      000B0000000D0000000D0000000E0000000E0000001000000015BE9688FFFEFD
                      FCFFF9F4F0FFF9F4EFFFF8F2EFFFF2E7E3FFF2E7E3FFF2E7E3FFF1E6E3FFF1E7
                      E3FFF1E7E3FFF1E6E2FFF1E6E2FFF1E6E2FFF1E6E2FFF1E6E1FFF1E6E1FFEFE3
                      DFFFF8F1EDFFF7F1EDFFF7F0ECFFFCFAF9FFB48979FF0000000A896C61BCC096
                      86FFBF9586FFBF9485FFBE9385FFBD9383FFBC9183FFD9C4BCFFC0998BFFFEFE
                      FDFFFAF5F1FFF9F4F0FFFAF4F0FFB78B7AFFB68A79FFB78A78FFB78A78FFB689
                      78FFB68A77FFB58977FFB68976FFB68877FFB58876FFB58775FFB48875FFA472
                      61FFF8F1EEFFF9F1EDFFF8F1EDFFFDFBFAFFB68C7CFF0000000AC19989FFFCF8
                      F6FFFCF8F6FFFBF8F6FFFBF8F5FFCEAFA3FFFAF7F4FFF7F5F4FFC29B8DFFFEFE
                      FDFFFAF6F2FFF9F5F2FFFAF6F1FFF3EAE6FFF2E9E5FFF2E9E5FFF2E9E5FFF2E9
                      E5FFF2E9E5FFF2EAE5FFF2E9E4FFF2E9E4FFF2E9E4FFF2E9E4FFF2E7E4FFF0E5
                      E1FFF9F4EFFFF8F2EFFFF8F2EFFFFDFCFBFFB88E7FFF00000009C39A8CFFFCF9
                      F7FFF7EFEAFFF7EFEAFFF7EFEAFFC8A698FFF6EEE9FFF6F2EFFFC39E8FFFFFFE
                      FEFFFAF7F4FFFAF7F4FFFAF6F4FFBA8F7EFFB98F7EFFB98F7DFFB98D7CFFB98D
                      7CFFB88C7CFFB88C7CFFB88C7AFFB78B7AFFB88B79FFB78B79FFB68B78FFA676
                      63FFF9F4F1FFF9F4F0FFF9F4F0FFFEFCFCFFBA9082FF00000008C59D8EFFFDFA
                      F9FFF8F1ECFFF7F1ECFFF7F0EBFFCAA89BFFF6EFEAFFF7F4F1FFC4A091FFFFFE
                      FEFFFCF7F6FFFAF7F5FFFCF7F5FFFCF7F5FFFCF6F5FFFAF6F5FFF3ECE8FFF5EB
                      E8FFF3EBE8FFF3EBE8FFF3EBE8FFF3EBE8FFF3EAE6FFF3EAE8FFF2EAE6FFF1E8
                      E4FFFAF5F2FFFAF5F1FFFAF5F1FFFEFDFCFFBB9285FF00000008C69F90FFD4B9
                      ADFFCCAC9FFFCCAC9FFFCBAB9EFFCBAB9DFFCAA99BFFE1D2CBFFC5A193FFFFFF
                      FFFFFCF8F6FFFCF8F7FFFCF7F6FFFCF8F6FFFCF7F6FFFCF7F6FFBB9180FFBB91
                      80FFBB917FFFBA907FFFBA907EFFBA907EFFBA907EFFB98D7DFFB98D7DFFA979
                      68FFFAF6F2FFFAF5F4FFFAF6F4FFFEFDFDFFBD9586FF00000007C7A193FFFDFB
                      FBFFF8F2EEFFF9F2EEFFF8F1EDFFCDADA0FFF8F0ECFFF9F6F3FFC8A395FFFFFF
                      FFFFFDF9F7FFFDF9F7FFFCF8F7FFFCF8F7FFFCF8F6FFFCF8F7FFFDF8F7FFFDF8
                      F6FFFCF8F7FFFCF8F6FFFCF7F6FFFCF7F6FFFCF7F6FFFAF7F5FFFCF7F5FFFAF7
                      F5FFFCF6F5FFFCF6F5FFFCF7F5FFFFFEFDFFBF9688FF00000006C9A395FFFEFC
                      FBFFF9F2F0FFF9F2EFFFF9F3EEFFCEAFA3FFF7F1EDFFF9F6F4FFC8A496FFFFFF
                      FFFFFDF9F9FFFDF9F9FFFDF8F8FFFDF9F7FFFCF9F7FFFDF9F8FFFDF9F7FFFCF8
                      F7FFFDF8F7FFFDF8F7FFFCF8F6FFFDF8F6FFFCF8F7FFFCF7F7FFFCF8F6FFFCF7
                      F6FFFCF7F5FFFCF7F6FFFCF7F6FFFFFEFEFFC1998BFF00000006CAA598FFFEFD
                      FCFFF9F3F0FFF9F4F0FFF9F3F0FFD0B2A6FFF8F2EFFFFAF7F5FFC9A597FFFFFF
                      FFFFFDF9F9FFFDFAF9FFFDFAF9FFFDFAF9FFFDF9F9FFFDF9F9FFFDFAF8FFFDF9
                      F9FFFDF9F8FFFDF9F8FFFDF9F8FFFCF8F7FFFCF8F7FFFCF9F7FFFCF8F7FFFDF8
                      F7FFFDF8F6FFFCF8F7FFFCF8F7FFFFFEFEFFC29B8DFF00000005CBA89AFFDCC3
                      BAFFD2B5A9FFD2B5A9FFD1B4A8FFD2B4A8FFD1B4A8FFE6D8D1FFCAA699FFFFFF
                      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFC39E8FFF00000004CCAA9CFFFEFE
                      FDFFFAF5F2FFFBF5F2FFFAF6F2FFD3B7ABFFFAF4F2FFFCF9F8FFD7BBB1FFCAA6
                      99FFCAA699FFCAA699FFCAA699FFCAA699FFC9A597FFCAA597FFC9A596FFC9A4
                      96FFC9A496FFC8A496FFC9A396FFC8A395FFC6A395FFC8A295FFC6A294FFC6A2
                      93FFC5A193FFC5A092FFC5A192FFC4A091FF91766BBF00000003CEAC9EFFFEFE
                      FEFFFAF6F4FFFBF7F3FFFBF6F4FFD4B9ADFFFAF5F3FFFDFAF9FFFCF9F8FFFCF9
                      F8FFE9DAD4FFFCF9F8FFFBF8F7FFFBF8F7FFFBF8F7FFE7D8D2FFFBF8F6FFFBF8
                      F6FFFAF8F6FFFCFCFBFFE1CEC7FF0000000B0000000500000003000000030000
                      0003000000030000000300000003000000030000000200000001D0AEA1FFFFFE
                      FEFFFBF8F5FFFBF7F5FFFBF6F5FFD6BBAFFFFBF7F5FFFBF6F4FFFAF6F4FFFBF6
                      F4FFD5B9AEFFFBF6F3FFFBF6F3FFFBF5F3FFFBF6F3FFD3B7ABFFFBF5F3FFFAF5
                      F3FFFAF5F2FFFEFDFCFFC9A295FF000000080000000200000001000000010000
                      0001000000010000000100000001000000010000000000000000D1B0A2FFE1CC
                      C4FFD8BEB3FFD8BEB3FFD7BDB2FFD7BDB2FFD7BCB1FFD6BCB1FFD6BCB0FFD6BB
                      B0FFD6BBB0FFD6BBAFFFD6BBAFFFD6BAAEFFD5BAAEFFD5B9ADFFD5B9ADFFD4B8
                      ADFFD4B8ACFFDCC4BBFFC9A597FF000000080000000200000000000000000000
                      0000000000000000000000000000000000000000000000000000D2B2A4FFFFFF
                      FFFFFCF9F8FFFCF8F7FFFCF9F7FFD8BFB4FFFCF8F7FFFCF8F6FFFCF8F7FFFBF8
                      F6FFD7BDB2FFFCF8F5FFFCF7F6FFFBF8F5FFFBF7F5FFD6BCB0FFFBF7F5FFFBF7
                      F4FFFBF6F5FFFEFEFDFFCAA799FF000000070000000200000000000000000000
                      0000000000000000000000000000000000000000000000000000D3B3A6FFFFFF
                      FFFFFDFAF8FFFCF9F8FFFCF9F8FFD9C0B5FFFCF9F8FFFDF9F8FFFDF9F7FFFCF9
                      F8FFD8BFB4FFFCF8F7FFFCF8F7FFFCF8F7FFFBF8F6FFD8BDB2FFFBF8F6FFFCF7
                      F6FFFCF7F6FFFFFEFEFFCCA99BFF000000060000000200000000000000000000
                      0000000000000000000000000000000000000000000000000000D4B4A8FFFFFF
                      FFFFFFFFFFFFFFFFFFFFFFFFFFFFE3CFC7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
                      FFFFE1CEC6FFFFFFFEFFFFFFFEFFFFFFFEFFFFFEFEFFE1CDC4FFFFFEFEFFFFFE
                      FEFFFFFEFEFFFFFEFEFFCDAB9DFF000000050000000100000000000000000000
                      00000000000000000000000000000000000000000000000000009D867DBED4B5
                      A8FFD4B5A8FFD4B4A8FFD4B5A7FFD4B4A7FFD4B3A7FFD4B3A6FFD3B3A5FFD3B2
                      A6FFD3B2A4FFD2B1A4FFD2B1A4FFD2B1A4FFD1B0A3FFD1B0A2FFD0AFA1FFD0AF
                      A1FFD0AEA1FFCFAEA0FF9A8076BF000000030000000100000000000000000000
                      0000000000000000000000000000000000000000000000000000000000010000
                      0002000000020000000200000003000000030000000300000003000000030000
                      0003000000030000000300000004000000040000000400000004000000040000
                      0004000000040000000400000003000000010000000000000000000000000000
                      0000000000000000000000000000000000000000000000000000000000000000
                      0000000000000000000000000001000000010000000100000001000000010000
                      0001000000010000000100000001000000010000000100000001000000010000
                      0001000000010000000100000001000000000000000000000000000000000000
                      0000000000000000000000000000000000000000000000000000000000000000
                      0000000000000000000000000000000000000000000000000000000000000000
                      0000000000000000000000000000000000000000000000000000000000000000
                      0000000000000000000000000000000000000000000000000000000000000000
                      0000000000000000000000000000000000000000000000000000}
                    TabOrder = 4
                    OnClick = btnAjustarClick
                  end
                  object dxLayoutControl2Group_Root: TdxLayoutGroup
                    AlignHorz = ahClient
                    AlignVert = avTop
                    ButtonOptions.Buttons = <>
                    Hidden = True
                    ShowBorder = False
                    Index = -1
                  end
                  object dxLayoutItem1: TdxLayoutItem
                    Parent = dxLayoutControl2Group_Root
                    AlignHorz = ahClient
                    CaptionOptions.Text = 'cxDBCheckBox1'
                    CaptionOptions.Visible = False
                    Control = DbChkBxPernocta
                    ControlOptions.ShowBorder = False
                    Index = 0
                  end
                  object dxLayoutControl2Item2: TdxLayoutItem
                    Parent = dxLayoutControl2Group_Root
                    AlignHorz = ahClient
                    CaptionOptions.Text = 'cxDBCheckBox1'
                    CaptionOptions.Visible = False
                    Control = DbChkBxMaterial
                    ControlOptions.ShowBorder = False
                    Index = 1
                  end
                  object dxLayoutControl2Item3: TdxLayoutItem
                    Parent = dxLayoutControl2Group_Root
                    CaptionOptions.Text = 'cxDBCheckBox1'
                    CaptionOptions.Visible = False
                    Control = DbChkBxPaginas
                    ControlOptions.ShowBorder = False
                    Index = 2
                  end
                  object dxLayoutControl2Item4: TdxLayoutItem
                    Parent = dxLayoutControl2Group_Root
                    CaptionOptions.Visible = False
                    Control = DbChkBxPdas
                    ControlOptions.ShowBorder = False
                    Index = 3
                  end
                  object dxLayoutControl2Item1: TdxLayoutItem
                    Parent = dxLayoutControl2Group_Root
                    CaptionOptions.Text = 'cxButton1'
                    CaptionOptions.Visible = False
                    Control = btnAjustar
                    ControlOptions.ShowBorder = False
                    Index = 4
                  end
                end
              end
              object cxDBMemo1: TcxDBMemo
                Left = 108
                Top = 196
                DataBinding.DataField = 'sFirma1'
                DataBinding.DataSource = dsActa
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                TabOrder = 8
                OnEnter = DbMmObservacionEnter
                OnExit = DbMmObservacionExit
                Height = 50
                Width = 849
              end
              object cxDBMemo2: TcxDBMemo
                Left = 108
                Top = 252
                DataBinding.DataField = 'sFirma2'
                DataBinding.DataSource = dsActa
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                TabOrder = 9
                OnEnter = DbMmObservacionEnter
                OnExit = DbMmObservacionExit
                Height = 50
                Width = 849
              end
              object DbTxtEdtActivo: TcxDBTextEdit
                Left = 438
                Top = 91
                DataBinding.DataField = 'sActivo'
                DataBinding.DataSource = dsActa
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                TabOrder = 4
                OnEnter = DbTxtEdtActivoEnter
                OnExit = DbTxtEdtActivoExit
                OnKeyPress = DbTxtEdtActivoKeyPress
                Width = 328
              end
              object DbTxtEdtCentro: TcxDBTextEdit
                Left = 108
                Top = 91
                DataBinding.DataField = 'sCentroProceso'
                DataBinding.DataSource = dsActa
                Style.BorderColor = clWindowFrame
                Style.BorderStyle = ebs3D
                Style.Color = 15138559
                Style.HotTrack = False
                TabOrder = 3
                OnEnter = DbTxtEdtCentroEnter
                OnExit = DbTxtEdtCentroExit
                OnKeyPress = DbTxtEdtCentroKeyPress
                Width = 285
              end
              object dxLayoutControl1Group_Root: TdxLayoutGroup
                AlignHorz = ahClient
                AlignVert = avClient
                ButtonOptions.Buttons = <>
                Hidden = True
                LayoutDirection = ldHorizontal
                ShowBorder = False
                Index = -1
              end
              object dxLayoutItem2: TdxLayoutItem
                Parent = dxLayoutAutoCreatedGroup1
                AlignHorz = ahClient
                AlignVert = avClient
                CaptionOptions.AlignVert = tavTop
                CaptionOptions.Text = 'Observaciones:'
                Control = DbMmObservacion
                ControlOptions.ShowBorder = False
                Index = 1
              end
              object dxLayoutControl1Item5: TdxLayoutItem
                Parent = dxLayoutControl1Group3
                AlignHorz = ahRight
                CaptionOptions.Text = 'cxDBRadioGroup1'
                CaptionOptions.Visible = False
                Control = DbRdGrpTipo
                ControlOptions.ShowBorder = False
                Index = 1
              end
              object dxLayoutControl1Group2: TdxLayoutAutoCreatedGroup
                Parent = dxLayoutAutoCreatedGroup1
                AlignHorz = ahClient
                LayoutDirection = ldHorizontal
                Index = 0
                AutoCreated = True
              end
              object dxLayoutItem3: TdxLayoutItem
                Parent = dxLayoutControl1Group4
                AlignHorz = ahClient
                CaptionOptions.Text = 'No. de Acta:'
                Control = DbTxtEdtActa
                ControlOptions.ShowBorder = False
                Index = 1
              end
              object dxLayoutControl1Item2: TdxLayoutItem
                Parent = dxLayoutControl1Group4
                AlignHorz = ahClient
                CaptionOptions.Text = 'Folio:'
                Control = DbLkpCmbFolio
                ControlOptions.ShowBorder = False
                Index = 0
              end
              object dxLayoutControl1Group4: TdxLayoutAutoCreatedGroup
                Parent = dxLayoutControl1Group2
                AlignHorz = ahClient
                Index = 0
                AutoCreated = True
              end
              object dxLayoutControl1Item3: TdxLayoutItem
                Parent = dxLayoutControl1Group4
                AlignHorz = ahClient
                CaptionOptions.Text = 'Especialidad:'
                Control = DbTxtEdtEspecialidad
                ControlOptions.ShowBorder = False
                Index = 2
              end
              object dxLayoutControl1Item4: TdxLayoutItem
                Parent = dxLayoutControl1Group3
                AlignHorz = ahRight
                CaptionOptions.Text = 'Fecha de Acta:'
                Control = DbDtEdtFecha
                ControlOptions.ShowBorder = False
                Index = 0
              end
              object dxLayoutControl1Group3: TdxLayoutAutoCreatedGroup
                Parent = dxLayoutControl1Group2
                AlignHorz = ahRight
                Index = 1
                AutoCreated = True
              end
              object dxLayoutControl1Item6: TdxLayoutItem
                Parent = dxLayoutControl1Group_Root
                AlignHorz = ahRight
                AlignVert = avClient
                CaptionOptions.Text = 'cxGroupBox1'
                CaptionOptions.Visible = False
                Control = GBx5
                ControlOptions.AutoColor = True
                ControlOptions.ShowBorder = False
                Index = 1
              end
              object dxLayoutAutoCreatedGroup1: TdxLayoutAutoCreatedGroup
                Parent = dxLayoutControl1Group_Root
                AlignHorz = ahClient
                Index = 0
                AutoCreated = True
              end
              object dxLayoutControl1Item7: TdxLayoutItem
                Parent = dxLayoutAutoCreatedGroup1
                AlignHorz = ahClient
                AlignVert = avBottom
                CaptionOptions.AlignVert = tavTop
                CaptionOptions.Text = 'Firmante Subsea'
                Control = cxDBMemo1
                ControlOptions.ShowBorder = False
                Index = 2
              end
              object dxLayoutControl1Item8: TdxLayoutItem
                Parent = dxLayoutAutoCreatedGroup1
                AlignHorz = ahClient
                AlignVert = avBottom
                CaptionOptions.AlignVert = tavTop
                CaptionOptions.Text = 'Firmante PEP'
                Control = cxDBMemo2
                ControlOptions.ShowBorder = False
                Index = 3
              end
              object dxLayoutControl1Item9: TdxLayoutItem
                Parent = dxLayoutControl1Group1
                AlignHorz = ahClient
                CaptionOptions.Text = 'Activo:'
                Control = DbTxtEdtActivo
                ControlOptions.ShowBorder = False
                Index = 1
              end
              object dxLayoutControl1Item1: TdxLayoutItem
                Parent = dxLayoutControl1Group1
                AlignHorz = ahLeft
                CaptionOptions.Text = 'Centro de Proceso:'
                Control = DbTxtEdtCentro
                ControlOptions.ShowBorder = False
                Index = 0
              end
              object dxLayoutControl1Group1: TdxLayoutAutoCreatedGroup
                Parent = dxLayoutControl1Group4
                LayoutDirection = ldHorizontal
                Index = 3
                AutoCreated = True
              end
            end
          end
        end
        object cTsMateriales: TcxTabSheet
          Caption = 'Materiales'
          ImageIndex = 1
          object CxGrd1: TcxGrid
            Left = 0
            Top = 0
            Width = 1135
            Height = 316
            Align = alClient
            TabOrder = 0
            object CxGrdDbTblVMateriales: TcxGridDBTableView
              Navigator.Buttons.CustomButtons = <>
              DataController.DataSource = dsMateriales
              DataController.Summary.DefaultGroupSummaryItems = <>
              DataController.Summary.FooterSummaryItems = <>
              DataController.Summary.SummaryGroups = <>
              OptionsData.Appending = True
              OptionsView.ColumnAutoWidth = True
              object CxGrdDbTblVMaterialesColumn1: TcxGridDBColumn
                Caption = 'Partida'
                DataBinding.FieldName = 'sNumeroActividad'
                PropertiesClassName = 'TcxLookupComboBoxProperties'
                Properties.DropDownListStyle = lsFixedList
                Properties.KeyFieldNames = 'sNumeroActividad'
                Properties.ListColumns = <
                  item
                    Caption = 'Partida'
                    Fixed = True
                    HeaderAlignment = taCenter
                    Width = 70
                    FieldName = 'sNumeroActividad'
                  end
                  item
                    Caption = 'Descripci'#243'n'
                    Fixed = True
                    HeaderAlignment = taCenter
                    Width = 250
                    FieldName = 'mDescripcion'
                  end>
                Properties.ListSource = dsActividades
                HeaderAlignmentHorz = taCenter
                Width = 70
              end
              object CxGrdDbTblVMaterialesColumn2: TcxGridDBColumn
                Caption = 'Id Material'
                DataBinding.FieldName = 'sIdInsumo'
                PropertiesClassName = 'TcxLookupComboBoxProperties'
                Properties.DropDownAutoSize = True
                Properties.DropDownListStyle = lsFixedList
                Properties.KeyFieldNames = 'sIdInsumo'
                Properties.ListColumns = <
                  item
                    Caption = 'C'#243'digo Insumo'
                    Fixed = True
                    HeaderAlignment = taCenter
                    Width = 70
                    FieldName = 'sIdInsumo'
                  end
                  item
                    Caption = 'Descripci'#243'n'
                    HeaderAlignment = taCenter
                    Width = 250
                    FieldName = 'mDescripcion'
                  end>
                Properties.ListOptions.SyncMode = True
                Properties.ListSource = dsInsumos
                Properties.OnCloseUp = CxGrdDbTblVMaterialesColumn2PropertiesCloseUp
                HeaderAlignmentHorz = taCenter
                Width = 100
              end
              object CxGrdDbTblVMaterialesColumn7: TcxGridDBColumn
                Caption = 'Material'
                DataBinding.FieldName = 'mDescripcion'
                PropertiesClassName = 'TcxMemoProperties'
                HeaderAlignmentHorz = taCenter
                Options.Editing = False
                Options.Focusing = False
                Width = 250
              end
              object CxGrdDbTblVMaterialesColumn3: TcxGridDBColumn
                Caption = 'Trazabilidad'
                DataBinding.FieldName = 'sTrazabilidad'
                HeaderAlignmentHorz = taCenter
                Width = 70
              end
              object CxGrdDbTblVMaterialesColumn4: TcxGridDBColumn
                Caption = 'U. Medida'
                DataBinding.FieldName = 'sMedida'
                HeaderAlignmentHorz = taCenter
                Width = 70
              end
              object CxGrdDbTblVMaterialesColumn5: TcxGridDBColumn
                Caption = 'Cantidad'
                DataBinding.FieldName = 'dCantidad'
                PropertiesClassName = 'TcxCalcEditProperties'
                HeaderAlignmentHorz = taCenter
                Width = 70
              end
              object CxGrdDbTblVMaterialesColumn6: TcxGridDBColumn
                Caption = '*'
                PropertiesClassName = 'TcxButtonEditProperties'
                Properties.Buttons = <
                  item
                    Default = True
                    Kind = bkEllipsis
                  end>
                Properties.ViewStyle = vsButtonsOnly
                Properties.OnButtonClick = CxGrdDbTblVMaterialesColumn6PropertiesButtonClick
                HeaderAlignmentHorz = taCenter
                Width = 20
              end
            end
            object CxGLvlGrid2Level1: TcxGridLevel
              GridView = CxGrdDbTblVMateriales
            end
          end
        end
      end
    end
    object cTs2: TcxTabSheet
      Caption = 'cTs2'
      ImageIndex = 1
      object GBx3: TcxGroupBox
        Left = 0
        Top = 0
        Align = alTop
        Caption = 'GBx3'
        TabOrder = 0
        Height = 105
        Width = 1143
        object dxLayoutControl3: TdxLayoutControl
          Left = 2
          Top = 18
          Width = 1139
          Height = 85
          Align = alClient
          TabOrder = 0
          object btnGuardar: TcxButton
            Left = 10
            Top = 10
            Width = 75
            Height = 55
            Caption = 'Guardar'
            OptionsImage.Glyph.Data = {
              36100000424D3610000000000000360000002800000020000000200000000100
              2000000000000010000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0001000000010000000200000004000000050000000600000007000000070000
              0006000000050000000400000002000000010000000100000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000001000000010000
              0003000000060000000B0000001000000015000000180000001A0000001A0000
              001800000016000000110000000C000000070000000400000001000000010000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000010000000100000003000000070000
              000E0000001707140E4D133324951D4E38D11D513BD9246448FF246448FF1E51
              3BD91C4E38D11233249607140F4F000000190000001000000008000000040000
              0001000000010000000000000000000000000000000000000000000000000000
              000000000000000000000000000100000001000000040000000B000000150C21
              176B1F563FE1257151FF278963FF299D72FF2AA176FF2BB07FFF2BAF80FF2AA2
              76FF2A9E72FF278964FF267151FF20573FE20C21186D000000180000000D0000
              0005000000010000000100000000000000000000000000000000000000000000
              0000000000000000000100000001000000050000000D040C09361C503ACF2678
              56FF2AA074FF2CB180FF2BB180FF2CB081FF2CB081FF2CB180FF2CB181FF2CB1
              80FF2CB080FF2CB081FF2CB180FF2AA074FF277957FF1D5039D1040C093A0000
              0010000000060000000100000001000000000000000000000000000000000000
              00000000000000000001000000050000000E09191255236248F029946BFF2CB1
              81FF2CB181FF2CB181FF2CB282FF2CB282FF2CB282FF2CB282FF2DB282FF2DB2
              82FF2CB182FF2CB281FF2CB181FF2CB181FF2CB181FF29956CFF246248F10919
              135A000000120000000600000001000000000000000000000000000000000000
              000000000000000000040000000C09191253256C4EFA2A9F74FF2CB181FF2DB1
              82FF2CB283FF2DB283FF2DB283FF2EB283FF2EB284FF2EB384FF2EB383FF2EB3
              84FF2EB384FF2EB283FF2EB383FF2EB383FF2DB282FF2DB182FF2A9F74FF256D
              4FFA091A12590000001000000005000000010000000000000000000000000000
              00000000000200000009050F0B38246549EE2EA47AFF2EB383FF2DB283FF2EB3
              83FF2EB384FF2EB385FF2EB484FF51C7A2FF60CFAEFF37B98DFF2EB485FF2FB5
              85FF2FB485FF2EB485FF2EB485FF2EB384FF2EB383FF2EB384FF2FB384FF2FA5
              7AFF23644AEF050F0B3E0000000C000000030000000100000000000000000000
              000100000005000000101D513BCB2E9770FF30B385FF2EB384FF2EB484FF2EB4
              85FF2FB586FF2FB586FF50C6A1FF32916EFF1E7652FF4AB995FF30B587FF31B6
              87FF30B686FF30B587FF30B587FF2FB586FF2EB486FF2EB485FF2EB485FF30B4
              86FF2E9871FF1D533CD000000016000000070000000100000000000000000000
              00020000000A0C231A652D7D5CFF34B689FF2EB384FF2EB485FF2FB586FF30B5
              86FF31B688FF4EC59FFF389774FF7CAE9AFFA2C4B6FF2F8C6AFF41BF95FF32B7
              89FF32B789FF31B688FF31B689FF31B688FF30B588FF30B686FF2FB486FF2EB4
              85FF34B68AFF2D7E5EFF0D241A6B0000000E0000000300000001000000010000
              00040000000F205B43DD32A67DFF30B587FF2FB586FF30B587FF31B688FF31B7
              88FF4BC49DFF3E9F7CFF6BA28BFFF9F5F1FFF5EFEAFF45896CFF4CB491FF35BB
              8DFF34B98BFF34B98AFF33B98AFF32B989FF32B789FF32B688FF31B688FF30B5
              87FF31B688FF33A77DFF215E44DF000000160000000700000001000000020000
              0007091B144E308061FF37B98BFF30B587FF30B688FF31B688FF33B78AFF48C4
              9BFF46A786FF5A967DFFF6F4F0FFF3E8DFFFF3E8DFFFC8D9D0FF247A58FF4FC5
              9EFF35BB8DFF35BB8DFF34BA8CFF34BA8BFF34BA8BFF33B98AFF32B989FF31B7
              88FF31B688FF37B98CFF308162FF0A1D15570000000B00000002000000020000
              0009153B2B923A9B78FF35B98BFF32B788FF32B989FF33B98BFF46C39AFF4CAF
              8EFF4A8A6FFFF4F4F1FFF4EAE2FFF4E9E0FFF3E8E0FFF7EDE7FF699D87FF409D
              7CFF3FC094FF37BC8FFF37BC8EFF36BB8EFF36BB8DFF35BA8DFF34BA8BFF34B9
              8BFF32B78AFF36BA8CFF3B9C79FF153C2C980000000E00000003000000020000
              000B1E553FC63DAC86FF35B88CFF33B78AFF33B98BFF45C398FF54B797FF4084
              67FFEAEFEAFFF5EBE3FFF2E6DEFFEDDFD6FFF4E9E1FFF4E9E0FFE3E8E1FF2570
              50FF56C19EFF39BE91FF39BE90FF38BD90FF37BC8FFF37BC8EFF35BB8DFF35BB
              8CFF33B98BFF35B98DFF3EAD87FF1E5640CA0000001100000004000000030000
              000C256A4EEC3EB88EFF34BA8CFF35B98CFF3FBF94FF57BD9DFF347B5CFFE5EC
              E8FFF6EDE6FFF1E6DDFFCAC3B6FF9DAA97FFEFE3DCFFF4EAE2FFF6EBE5FF9DBD
              AFFF318464FF4FC8A2FF3ABF94FF3ABF92FF39BE91FF38BD90FF37BD8FFF36BB
              8EFF35BB8CFF35BB8CFF3FB98FFF256D50ED0000001300000005000000030000
              000C287455FA43C096FF35BA8BFF36BB8DFF37B488FF1F704EFFC9D0C8FFF7ED
              E7FFF1E5DEFFBEBCB0FF2E7354FF246F4EFFB5B7A8FFF1E6DFFFF4EAE3FFF6F1
              ECFF548B72FF4EAD8DFF42C49AFF3CC195FF3BC094FF3BBF92FF39BE90FF39BD
              90FF37BC8EFF35BB8DFF44C197FF287657FA0000001300000005000000030000
              000B287656FA4CC49BFF35BC8DFF37BC8EFF37BC8EFF24805CFF608A72FFE4D5
              CCFFB2B5A7FF2C7757FF3CBD96FF3BBB93FF2B7151FFC8C3B6FFF4E9E2FFF5EB
              E4FFE1E8E2FF2D7355FF5FCAABFF40C59AFF3EC197FF3CC195FF3BC094FF3ABF
              92FF38BD90FF38BC8EFF4EC59DFF297859FA0000001200000005000000020000
              000A267052EC55C39FFF39BD91FF38BE90FF3ABE92FF3CC096FF257E5BFF4A7C
              61FF2B7F5EFF41CBA3FF45D3ACFF46D3ADFF39B28DFF3C7457FFDACFC5FFF5EB
              E4FFF7EDE8FFADC7BAFF2E8061FF5ED8B7FF43CDA4FF40C99FFF3DC399FF3CBF
              94FF3ABF92FF3BBF92FF56C5A0FF277256ED0000001100000004000000020000
              0008205D46C559BD9DFF3DC094FF3CC296FF43CDA6FF45D2ACFF44CEA8FF319D
              7AFF46D2ACFF48D5B0FF49D5B0FF49D5B0FF49D5B1FF34A481FF4C7B60FFE0D2
              CAFFF5ECE5FFF9F3EFFF699882FF459D7FFF57D6B3FF44CFA6FF43CDA4FF41C8
              A0FF3DC196FF3FC196FF5BBF9FFF205F47C80000000E00000003000000010000
              00061643328F58B194FF49CDA6FF47D2ACFF48D3AEFF49D5AFFF49D5B0FF4AD7
              B2FF4BD7B2FF4CD8B4FF4DD7B4FF4DD8B4FF4CD8B4FF4DD7B4FF329B78FF4B7B
              61FFDFD2C9FFF6EBE5FFEFF0ECFF468066FF58B699FF55D6B1FF45CFA7FF45CE
              A5FF43CAA3FF48C9A2FF58B295FF174433940000000B00000003000000010000
              00040B201848409E80FF63DEC0FF4BD4B1FF4CD7B2FF4DD7B3FF4ED7B4FF4FD8
              B4FF4FD9B7FF50DAB7FF50DAB7FF51DAB8FF51DAB7FF50DAB7FF51D9B7FF38A6
              84FF47795EFFDDD0C7FFF6ECE7FFF0F2EFFF478368FF61C1A5FF56D5B3FF47CF
              A8FF46CDA6FF62D8B8FF409C7DFF0B20184E0000000700000002000000000000
              0002000000072D7D62DA6CD4BBFF55D9B7FF51D8B5FF51D9B6FF53DAB8FF53DB
              BAFF54DCB9FF55DBBAFF55DCBBFF56DCBBFF55DCBAFF56DDBBFF56DDBAFF55DC
              BAFF40B090FF487A60FFD9CCC4FFF3E9E3FFEDF1EDFF4B876DFF67C7ACFF57D6
              B3FF4ED2ADFF6ECFB6FF29785DDC0000000E0000000400000001000000000000
              0001000000041231275B48A98CFF72E5CAFF56DBB9FF56DBBAFF58DCBCFF58DC
              BCFF59DDBDFF59DEBDFF5ADFBEFF5ADFBFFF5BDFBFFF5ADFBEFF5ADFBEFF59DE
              BDFF59DDBCFF47BB9AFF367559FFBAB9ADFFECDFD8FFDDDFDAFF237150FF45BA
              96FF70DFC3FF46A487FF10302561000000080000000200000000000000000000
              000000000002000000062B765EC66CCEB6FF6AE3C6FF5CDDBDFF5DDEBFFF5EDE
              C0FF5EDFC0FF5FE1C2FF60E1C2FF5FE1C2FF60E1C2FF5FE0C2FF5EE1C1FF5EE0
              C1FF5DDFBFFF5CDFBEFF55CFAFFF2E8464FF72937DFF8FA392FF2D8463FF63D8
              B9FF6DCBB1FF287259C80000000C000000040000000100000000000000000000
              00000000000100000003081410283B9679EC7EDFCBFF6FE4C9FF63E0C3FF63E0
              C3FF64E1C4FF65E2C4FF64E3C5FF64E3C5FF64E3C5FF65E3C5FF64E2C5FF63E2
              C4FF63E2C3FF61E0C1FF5FDFBFFF5EDEBDFF48B797FF2A8362FF61D2B5FF80DB
              C6FF379274ED07140F2D00000006000000020000000000000000000000000000
              00000000000000000001000000041028204644A689F982DFCBFF7EEAD3FF69E2
              C6FF69E3C7FF6AE4C7FF6AE4C8FF6AE5C9FF6AE4C8FF6AE4C8FF6AE4C8FF69E4
              C7FF68E3C6FF66E2C4FF64E0C2FF62DFC1FF61DDBEFF7AE5CDFF82DCC7FF40A2
              83FA0E271F4A0000000700000002000000000000000000000000000000000000
              000000000000000000000000000100000004102921453F9E81EE77D5BEFF93F1
              DFFF7BE9D1FF6FE5CAFF6FE6CBFF70E6CBFF70E6CBFF6FE7CCFF6EE6CBFF6DE6
              CAFF6CE4C8FF6BE3C6FF69E2C5FF75E5CBFF92EEDAFF75D2BAFF3B9B7CEE0F28
              204A000000070000000200000001000000000000000000000000000000000000
              000000000000000000000000000000000001000000030814102335856DC85ABF
              A3FF8BE4D2FF9DF4E5FF8DEFDCFF82EBD5FF7EEBD4FF75E8CFFF74E8CEFF7DEA
              D2FF7FEAD3FF8CEDDAFF9DF2E2FF8BE4D0FF58BEA1FF318469CA071410280000
              0006000000020000000100000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000100000002000000051738
              2E563D987CDA54BD9EFF75D4BCFF8EE6D3FF94EAD9FFA7F5E8FFA7F5E8FF95EA
              D9FF8DE6D3FF73D3BAFF52BB9CFF399679DA16392E5B00000007000000040000
              0002000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000001000000010000
              0003000000040D201A32255A4A82388D73C53B957ACE49B896FC49BA98FC3A95
              7ACF378D74C6235A4A840C201A34000000060000000400000002000000010000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0001000000010000000200000003000000040000000500000005000000050000
              0005000000040000000400000003000000020000000100000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000001000000010000000100000001000000010000
              0001000000010000000100000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000}
            OptionsImage.Layout = blGlyphTop
            TabOrder = 0
            OnClick = btnGuardarClick
          end
          object btnCancelar: TcxButton
            Left = 91
            Top = 10
            Width = 75
            Height = 55
            Caption = 'Cancelar'
            OptionsImage.Glyph.Data = {
              36100000424D3610000000000000360000002800000020000000200000000100
              2000000000000010000000000000000000000000000000000000000000000000
              0000000000000000000100000001000000010000000100000001000000010000
              0001000000010000000100000001000000010000000100000001000000010000
              0001000000010000000100000001000000010000000100000001000000010000
              0001000000010000000100000001000000000000000000000000000000000000
              0001000000010000000200000005000000060000000600000006000000060000
              0006000000070000000700000007000000070000000700000007000000070000
              0007000000070000000700000007000000070000000700000007000000070000
              0007000000060000000600000003000000010000000100000000000000000000
              0001000000040000000B00000013000000170000001800000019000000190000
              001900000019000000190000001A0000001A0000001A0000001A0000001A0000
              001B0000001B0000001B0000001B0000001B0000001B0000001C0000001C0000
              001C0000001B000000160000000E000000050000000100000000000000010000
              00020000000A0E0A2C5B2A208BEB2F2399FF2F2399FF2E2298FF2D2198FF2C20
              96FF2C2096FF2B1F95FF2B1F95FF2A1E94FF2B1D93FF291C93FF291B93FF291A
              93FF291A92FF271992FF271991FF271990FF26188FFF26178FFF26168FFF2416
              8EFF25168DFF201380EC0A06275E0000000D0000000300000001000000010000
              0004000000112D238DE8272DB1FF222FB9FF212DB7FF202DB7FF1F2BB6FF1F2B
              B6FF1F2BB6FF1E2AB6FF1E2AB6FF1F2BB4FF1E2BB6FF1E2AB6FF1E2AB6FF1E2A
              B4FF1D29B4FF1D29B4FF1C27B4FF1C28B4FF1B27B4FF1B27B3FF1B27B3FF1A26
              B3FF1B26B3FF1D20A8FF21147FE9000000150000000600000001000000010000
              000500000014342A9EFF2E3EC6FF2A3BC4FF2A3AC5FF2A3AC4FF2A3AC4FF2A3A
              C4FF2A3AC3FF2A38C4FF2A38C4FF2A38C4FF2937C3FF2837C3FF2A37C3FF2837
              C4FF2836C3FF2736C2FF2836C2FF2735C2FF2735C2FF2635C1FF2533C1FF2533
              C1FF2533C1FF1C27B3FF261890FF000000190000000600000001000000010000
              000500000015372DA1FF3142C7FF2D3DC6FF2E3CC5FF2D3EC5FF2D3DC6FF2C3C
              C5FF2D3CC5FF2C3DC5FF2C3CC5FF2D3BC5FF2C3CC4FF2C3AC4FF2B39C4FF2B3B
              C4FF2B39C4FF2B39C3FF2A38C3FF2939C3FF2937C3FF2837C3FF2736C3FF2836
              C2FF2735C1FF1D29B4FF281A92FF0000001A0000000700000001000000000000
              000500000014382FA4FF3747C9FF3041C7FF3041C7FF2F40C7FF3041C7FF2F40
              C7FF2F40C7FF2F40C6FF303EC6FF2F3FC6FF2E3FC6FF2F3DC6FF2E3EC7FF2E3C
              C6FF2D3DC6FF2D3BC6FF2C3CC5FF2D3AC5FF2B39C4FF2B39C4FF2B39C4FF2A38
              C4FF2937C4FF1F2CB7FF2A1D94FF000000190000000600000001000000000000
              0005000000133B32A5FF3A4CCAFF3243C9FF3343C9FF3243C8FF3243C9FF3242
              C8FF3646C9FF4A5BD2FF5363D5FF4757D1FF3242C9FF3142C8FF3142C8FF3241
              C8FF4555D0FF5060D4FF4857D1FF3340C7FF2E3EC7FF2E3EC6FF2E3BC5FF2D3B
              C5FF2B3CC5FF222DB8FF2D2095FF000000190000000600000001000000000000
              0005000000133D34A8FF3E4FCEFF3547CAFF3547CAFF3546CAFF3547CBFF3849
              CCFF4F5DD2FF444BBCFF3D41B2FF464DBEFF4C5CD1FF3545C9FF3545C9FF4B5B
              D1FF424ABBFF393BACFF3E44B6FF4B59CFFF3444C9FF313FC7FF303FC8FF2F40
              C7FF2F3EC6FF2430BBFF2F2198FF000000180000000600000001000000000000
              0004000000123E36AAFF4357CFFF384ACCFF3849CCFF384ACDFF394ACCFF4E5E
              D2FF4046B6FF9597D4FFE8E8F6FF8081CBFF4148B8FF4F5FD4FF4E5ED4FF3F45
              B7FF7E80C9FFE7E8F5FF9293D0FF3A3EB0FF4A58CEFF3444C9FF3243C8FF3242
              C8FF3142C8FF2734BCFF30249AFF000000170000000600000001000000000000
              0004000000114139ADFF495CD2FF3B4ECDFF3B4FCEFF3C4DCEFF3B4CCEFF444A
              BAFF9597D4FFF9F4F1FFF5ECE6FFF9F5F4FF8182CBFF4046B7FF3F45B6FF7F81
              C9FFF9F5F3FFF4EBE5FFF9F4F1FF9192D0FF3C40B3FF3647CBFF3646CAFF3545
              C9FF3444C9FF2A38BEFF32269CFF000000160000000600000001000000000000
              000400000010413BAFFF4D61D3FF3E52D0FF3E52CFFF3E51D0FF3E52CFFF3B3E
              AFFFE3DADCFFE7D9D1FFF4EAE4FFF4EBE4FFF9F5F4FF7F81C9FF7F80C8FFF9F5
              F4FFF4E9E2FFF4E9E2FFF3E8E1FFE9E2E5FF3233A6FF394ACCFF394ACCFF3849
              CCFF3747CBFF2B3CC1FF34299EFF000000160000000500000001000000000000
              000400000010443EB2FF5267D7FF4155D2FF4156D1FF4255D1FF4155D1FF3C44
              B9FF6E68AFFFD6C2BBFFE9DAD2FFF4EBE5FFF4EBE5FFF9F6F5FFF9F6F4FFF4EA
              E4FFF4EAE3FFF4EAE3FFE5D6D0FF6A64ADFF343BB1FF3D4DCEFF3B4CCEFF3B4C
              CDFF3A4BCDFF2F3EC3FF372DA1FF000000150000000500000001000000000000
              00040000000F4540B4FF586CD8FF4559D3FF465AD3FF4458D3FF4559D3FF4457
              D2FF3940B4FF6D66AEFFD7C2BCFFE9DBD3FFF6ECE6FFF6ECE5FFF4ECE5FFF4EB
              E5FFF6EBE5FFE7D8D1FF6A63ABFF3339AEFF4051CEFF3F51D0FF3F51CFFF3D50
              CFFF3D4FCEFF3141C5FF392FA3FF000000140000000500000001000000000000
              00040000000E4742B6FF5C72DAFF475CD5FF485DD4FF475BD4FF495CD5FF485C
              D5FF495CD3FF3E44B5FF6D67AFFFE9DAD4FFF7EDE9FFF6EDE9FFF6EDE8FFF6ED
              E6FFF4EBE5FF706AB4FF393EB0FF4456D0FF4356D2FF4354D2FF4153D1FF4153
              D1FF3F52D0FF3545C8FF3B31A5FF000000130000000500000000000000000000
              00030000000E4A45B9FF6178DDFF4B61D7FF4B60D6FF4B60D6FF4B5FD6FF4C60
              D6FF6074DBFF4247B5FF7A79C1FFF5EFE9FFF7F0E9FFF6EFE9FFF6EFE9FFF6EF
              E9FFF4ECE7FF7977BFFF3C3FB0FF5E70DAFF475AD4FF4758D3FF4558D2FF4457
              D2FF4356D1FF3849CAFF3B33A8FF000000130000000500000000000000000000
              00030000000D4C47BBFF667CDEFF4D64D8FF4E64D8FF4E64D7FF4D62D7FF6174
              DCFF484CB8FF8080C7FFFAF8F7FFF7F0EBFFF7F1EBFFF7F0EBFFF7F0EAFFF7F0
              EAFFF7F0EAFFFAF7F7FF7A7AC3FF3E42B0FF5C6FDAFF4A5CD5FF485CD4FF465A
              D4FF475AD3FF3A4CCCFF3E36AAFF000000120000000500000000000000000000
              00030000000C4D49BDFF6A82E0FF5067DAFF5066D9FF5066D9FF5166D9FF4C53
              BCFF7D7CC5FFFAF8F7FFF8F2EFFFF8F1EDFFF8F2ECFFE9DDD8FFDECEC8FFEADE
              D8FFF7F0EBFFF7F0EBFFFAF8F7FF7775C0FF4348B4FF4C5FD6FF4B5DD6FF4A5E
              D5FF4A5CD5FF3E50CEFF3E38ACFF000000110000000400000000000000000000
              00030000000C4D4BC0FF6E86E2FF536ADBFF5369DBFF5D73DCFF6F83E1FF3A3A
              A9FFE8E2E7FFEDE3DEFFF9F2EFFFF8F3EFFFE8DDD9FF6F68ADFF6D65A9FFD8C6
              BFFFEADFD8FFF8F2ECFFF7F1ECFFEDE9EEFF2F2D9EFF6073DCFF5367D9FF4D61
              D7FF4C5FD7FF4155D0FF413BAFFF000000110000000400000000000000000000
              00030000000B4F4DC2FF728AE4FF5E74DEFF798CE4FF8396E6FF8396E6FF4C51
              B7FF8179B0FFDCCBC4FFEADFDCFFE7DBD7FF6F68ADFF474AB2FF4649B2FF6C64
              A9FFD8C7C0FFE7DCD7FFEBDFD9FF7E75B0FF4043AFFF6E81E0FF6C7FE0FF6477
              DEFF5267D9FF4457D1FF423DB1FF000000100000000400000000000000000000
              00030000000A5050C4FF8A9FE9FF879AE8FF899BE8FF889CE8FF889AE8FF8191
              E2FF4547B0FF7E75ADFFC3B1B7FF6D66ABFF484DB3FF8091E3FF7E90E3FF4549
              B0FF6B64A9FFC2B0B6FF796FAAFF393BA7FF6F7FDCFF7386E2FF7284E1FF6E82
              E1FF6C7FDFFF5366D7FF4741B4FF0000000F0000000400000000000000000000
              00020000000A6164CCFFA2B6EFFF8DA1EAFF8EA1EAFF8EA1EAFF8D9FEAFF8D9F
              EAFF8594E2FF4D51B5FF3534A2FF5358BBFF8696E6FF8899E8FF8697E8FF8292
              E4FF4E53B8FF2E2C9CFF4347AFFF7585DEFF7B8DE4FF788BE3FF7589E3FF7487
              E2FF7084E1FF6478DDFF5755BEFF0000000E0000000400000000000000000000
              0002000000096C6FD2FFA6BCF1FF93A6ECFF93A6ECFF93A6ECFF93A5ECFF92A4
              EBFF92A4EBFF91A3EBFF90A2EBFF8FA1EAFF8E9FE9FF8D9DEAFF8B9CE9FF8A9B
              E9FF8799E8FF8697E8FF8395E7FF8293E6FF8091E6FF7E90E6FF7B8DE5FF788B
              E5FF7588E3FF697CDFFF6363C5FF0000000E0000000300000000000000000000
              0002000000087075D5FFABC1F2FF99ABEDFF99ABEDFF98ABEDFF98AAEDFF97A9
              EDFF97A8ECFF96A7ECFF96A6ECFF95A5EBFF93A4EBFF92A3EBFF90A1EAFF8F9F
              EAFF8D9EEAFF8B9CE9FF899AE8FF8698E8FF8595E7FF8294E7FF8092E6FF7D90
              E6FF7A8DE5FF6E82E0FF6768C8FF0000000D0000000300000000000000000000
              000200000008757AD8FFB1C5F3FF9DB1EFFF9DB0EFFF9DB0EFFF9DAFEFFF9DAE
              EEFF9BADEEFF9BACEEFF9AABEDFF99AAEDFF98A8EDFF96A7ECFF95A6ECFF93A4
              EBFF91A3EBFF8FA0EAFF8D9FEAFF8B9DE9FF889BE9FF8799E8FF8396E8FF8194
              E7FF7F91E6FF7386E3FF6B6ECBFF0000000C0000000300000000000000000000
              0002000000077A7FDAFFBCCFF5FFA2B5F0FFA2B4F0FFA2B4F0FFA2B4F0FFA1B3
              EFFFA0B2EFFFA0B1EFFF9FB0EFFF9EAEEFFF9CADEEFF9BABEDFF99AAEDFF98A8
              EDFF96A7EDFF93A5ECFF91A3ECFF8FA1EBFF8D9FEAFF8B9DEAFF889BE9FF8598
              E8FF8396E8FF798CE4FF7074CFFF0000000B0000000300000000000000000000
              0001000000057277C8E8BECAF1FFD4E1F9FFD3E0F9FFD2DFF9FFD0DEF9FFCFDD
              F9FFCDDBF7FFCBDAF7FFC9D7F7FFC6D5F6FFC4D4F6FFC2D1F6FFC0CFF5FFBDCE
              F5FFBACBF4FFB7C9F4FFB5C6F3FFB2C3F3FFAEC1F1FFABBFF1FFA8BCF1FFA5B9
              F1FFA2B7F0FF90A0E6FF6A6EBEE9000000080000000200000000000000000000
              00010000000320223744656BAFCB8087DEFF7F87DEFF7F87DDFF7F86DDFF7F86
              DCFF7F85DCFF7E84DCFF7D84DCFF7D84DBFF7D84DBFF7C83DBFF7C82D9FF7C82
              D9FF7B82D9FF7A81D9FF7A81D8FF7A80D8FF7A80D7FF797FD7FF787FD6FF7A7F
              D6FF787ED6FF5F63A8CD1E1F3447000000050000000100000000000000000000
              0000000000010000000200000004000000050000000600000006000000060000
              0006000000070000000700000007000000070000000700000007000000080000
              0008000000080000000800000008000000080000000800000009000000090000
              0009000000080000000700000004000000020000000000000000000000000000
              0000000000000000000100000001000000010000000100000001000000010000
              0001000000010000000200000002000000020000000200000002000000020000
              0002000000020000000200000002000000020000000200000002000000020000
              0002000000020000000200000001000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000}
            OptionsImage.Layout = blGlyphTop
            TabOrder = 1
            OnClick = btnCancelarClick
          end
          object btnExcel: TcxButton
            Left = 1080
            Top = 10
            Width = 49
            Height = 55
            Caption = 'Exportar '
            OptionsImage.Glyph.Data = {
              36100000424D3610000000000000360000002800000020000000200000000100
              2000000000000010000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000000000000010000000100000001000000010000
              0001000000010000000100000001000000010000000100000005000000140000
              001F000000210000002100000021000000220000002200000022000000230000
              0023000000230000002300000021000000160000000500000000000000000000
              0000000000000000000100000002000000040000000600000007000000070000
              000700000007000000070000000700000007000000070000001946332CCC6045
              3BFF644A41FFBD8150FFBC7E4DFFB97949FFB67646FFB37141FFB06D3DFFAD68
              39FFAB6535FF553A34FF593D35FF392621CE0000001500000000000000000000
              00000000000000000002000000070000000F00000016000000190000001A0000
              001A0000001B0000001B0000001B0000001B0000001B00000032664A40FF8165
              5AFF6A4F46FFE8C28BFFE7C088FFE6BD85FFE5BB81FFE4B87CFFE3B579FFE2B2
              76FFE2B273FF5A3F37FF664940FF523730FF0000001E00000000000000000000
              000000000001000000040000000F78554AC1A57666FFA57565FFA57465FFA574
              64FFA37463FFA47363FFA37362FFA37262FFA27162FFBDA79FFF6A4E42FF8369
              5FFF70564BFFD9B27DFFD8B07BFFD7AE77FFD7AB74FFD6A970FFD5A66DFFD4A5
              6AFFD4A268FF5E433CFF6F5147FF543931FF0000001D00000000000000000000
              0000000000010000000500000014A77868FFFDFCFAFFFBF8F6FFFBF8F5FFFBF7
              F4FFFBF7F4FFFAF7F4FFFAF6F3FFFAF6F2FFFAF5F2FFE4E0DCFF6E5246FF866C
              63FF765C50FFFFFFFFFFF7F1EBFFF7F0EBFFF7F0EBFFF7EFEBFFF6EFEAFFF6EF
              EAFFF6EFE9FF644940FF715349FF563B33FF0000001B00000000000000000000
              0000000000010000000600000016AA796AFFFDFCFBFFF6ECE6FFF6ECE6FFF6EC
              E6FFF6ECE5FFF4EBE5FFF4EBE5FFF4EBE4FFF4EBE4FFE1D9D2FF725648FF8A70
              65FF7B6154FFFFFFFFFFF8F2EDFFF8F1EDFFF7F1EDFFF7F0EDFFF8F1EBFFF7F0
              EBFFF7F0ECFF6A4F46FF72554BFF5A3E36FF0000001900000000000000000000
              0000000000010000000500000015AA7C6CFFFDFCFBFFF7EDE8FFF7EDE8FFF6EC
              E6FFF4EDE6FFF4ECE6FFF4ECE6FFF6ECE5FFF4ECE4FFE3DAD4FF755A4CFF8E75
              6AFF7F6458FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFF70564BFF75584EFF5C4138FF0000001700000000000000000000
              0000000000010000000500000015AC7D6FFFFEFDFCFFF7EFE9FFF7EDE8FFF7ED
              E9FFF6EDE8FFF6EDE6FFF6EDE6FFF6ECE6FFF6ECE5FFE7DDD7FF7A5E50FF9078
              6EFF82685BFF82675AFF806659FF7F6558FF7E6357FF7D6356FF7A6055FF795F
              53FF775D52FF765B50FF765A50FF5F443BFF0000001500000000000000000000
              0000000000000000000500000014AD7F70FFFEFDFCFFF7F0EAFFF7EFE9FFF7EF
              E9FFF7EFE9FFF7EFE8FFF6EDE8FFF6EDE8FFF6EDE6FFE8E0D9FF7E6253FF947C
              71FF674E44FF654B42FF634A41FF61473FFF5F473EFF5C443CFF5B433AFF5941
              39FF584038FF573F37FF775C52FF63473DFF0000001300000000000000000000
              0000000000000000000500000013AE8172FFFEFDFCFFF7F0EAFFF7F0EAFFF7F0
              E9FFF6EFE9FFF7EFE9FFF7EFE8FFF7EFE9FFF6EDE8FFEAE1DCFF816656FF9680
              75FF6B5248FFF4ECE6FFE9DACEFFE9D8CDFFE9D8CCFFE9D8CBFFE8D7CAFFF3EA
              E2FFF3E9E2FF5A4139FF795E54FF664B40FF0000001100000000000000000000
              0000000000000000000400000012AF8475FFFEFDFDFFF8F1EBFFF8F1EBFFF8F0
              EBFFF7F0EBFFF7F0EAFFF7F0EAFFF7F0E9FFF7EFE9FFEBE5DFFF856A59FF9983
              79FF70564DFFF4ECE6FFEBDACFFFEADACFFFE9D9CDFFE9D9CCFF513730FF6549
              3EFFF3EAE3FF5D453CFF7B6156FF6A4F43FF0000000F00000000000000000000
              0000000000000000000400000011B18676FFFEFDFDFFF8F1EDFFF8F2ECFFF8F1
              EBFFF7F1EBFFF7F1EBFFF7F0EBFFF8F0EAFFF7F0EAFFEEE7E2FF896E5CFF9C87
              7DFF755A50FFF5EDE8FFEBDCD2FFEADCD0FFEADACFFFEAD9CEFF49312BFF5D40
              39FFF4EBE4FF60483FFF7D6358FF6E5346FF0000000D00000000000000000000
              0000000000000000000400000010B48878FFFEFDFDFFF9F2EDFFF8F2ECFFF8F2
              ECFFF8F1ECFFF8F1ECFFF7F1ECFFF7F0EBFFF8F1EAFFF1E9E4FF8D7260FF9F8A
              81FF795E54FFF5EEE9FFECDED4FFEBDCD2FFEADCD1FFEADBD0FF452D27FF472E
              29FFE9D9CDFF644C43FF7F655AFF72574AFF0000000B00000000000000000000
              000000000000000000040000000FB58979FFFEFEFDFFF9F3F0FFF8F2EDFFF8F2
              EDFFF8F2EDFFF8F2EDFFF8F1EDFFF8F1ECFFF8F1ECFFF2EBE5FF917663FFA28D
              83FF7C6157FFF5EFEAFFF5EEE9FFF5EEE9FFF5EDE8FFF5EDE7FFF5ECE6FFF4EC
              E6FFF4ECE6FF695046FF998278FF765B4DFF0000000900000000000000000000
              000000000000000000030000000EB78C7DFFFEFEFEFFF9F4F0FFF9F3F0FFF9F3
              EFFFF8F2EFFFF8F2EDFFF8F2EDFFF8F2EDFFF8F1ECFFF5EEEAFFAC9686FF9377
              64FF7F645AFF998178FF967F75FF937A72FF8E786DFF8B7269FF866E64FF8269
              5FFF7D645BFF6E544AFF7C6052FF5B463BC20000000500000000000000000000
              000000000000000000030000000DDBC7BFFFFEFEFEFFF9F4F1FFF9F4F0FFF9F3
              F1FFF9F3F0FFF8F3EFFFF8F2EFFFF9F2EFFFF8F3EFFFF8F2EFFFF6EFEBFFF5EE
              E9FFF4EDE8FFF4EDE8FFF4EDE7FFF4EDE7FFF2EBE6FFF2EBE6FFF2EAE5FFF7F3
              F1FFD2BCB4FF000000190000000B00000004000000010000000000000000035C
              22B7047F2FFF047E2FFF047D2EFF037C2EFF037B2DFF027A2CFF02792CFF0279
              2BFF01782BFF01772BFF01772AFF00762AFF469961FFF9F3EFFFF8F2EFFFF9F2
              EFFFF8F2EFFFF8F2EDFFF8F1ECFFF8F1ECFFF7F1ECFFF7F0EBFFF7F0EBFFFAF6
              F3FFAE8373FF0000001200000004000000000000000000000000000000000581
              30FF0A9A46FF079139FF069038FF068F38FF058D37FF058C36FF038B34FF038A
              34FF028933FF028632FF028532FF018531FF01762AFFF9F3F0FFF9F3EFFFF9F3
              EFFFF9F2EFFFF9F3EDFFF9F2EDFFF8F1EDFFF8F1ECFFF8F1ECFFF6EFE9FFF8F4
              F2FFB08374FF0000001100000004000000000000000000000000000000000581
              31FF0C9C48FFFAF6F6FF079139FFF9F5F3FF068E38FFF8F2F1FFF7F2EFFF7EBE
              91FFF7EFEDFFF6EEEBFF7CBB8EFF028631FF01762AFFF9F4F0FFF9F3F1FFF9F4
              F0FFF9F3F0FFF9F3F0FFF9F3EFFFF9F3EFFFF8F2EFFFF6F0EAFFF5EDE7FFF6F1
              EEFFB38576FF0000001000000004000000000000000000000000000000000682
              32FF0E9E49FFFBF7F6FF08923AFFF9F5F4FF079038FFF8F4F2FF058D37FF038C
              36FF038B34FF7DBD90FFF6EEEBFF028632FF01772AFFFAF6F2FFFAF6F1FFF9F4
              F1FFF8F3F0FFF9F4F1FFF8F3EFFFF8F2EEFFF6F0EBFFF4EDE8FFF2E9E5FFF3EC
              E9FFB38978FF0000000F00000004000000000000000000000000000000000683
              32FF0E9F4BFF82C799FFFBF7F6FF81C498FF079139FFF9F5F2FF068E37FF058D
              36FF7FBF93FFF6F1EEFF038A33FF028833FF01782BFFFAF6F2FFF9F6F3FFFAF6
              F2FFF8F2EFFFF6EFEBFFF5EDE9FFF3EAE6FFF0E5E2FFEEE2DDFFEBDED9FFECE1
              DDFFB5897AFF0000000E00000004000000000000000000000000000000000785
              33FF0FA24CFFFCF8F8FF08943BFFFBF7F6FF08923AFFFAF5F3FF068F39FF068E
              37FFF8F3F0FF058C34FF038B34FF038A34FF01792BFFFAF7F4FFF9F5F2FFF9F5
              F1FFF5EEE9FFEEE2DCFFE6D8D0FFE1D2CAFFE0CEC7FFDECAC2FFDBC7BEFFDCC8
              C2FFB78C7DFF0000000D00000003000000000000000000000000000000000785
              33FF11A34EFFFBFAF9FF09953CFFFAF8F7FF08933AFFFAF7F5FF079139FF078F
              39FF7FC194FFF8F2F0FFF7F2EFFF038B34FF02792CFFFAF7F4FFFBF7F4FFF8F2
              EFFFEFE6DFFFB38B7CFFA57766FFA47564FFA47464FFA47363FFA37363FFCFB5
              ACFFB78C7DFF0000000A00000003000000000000000000000000000000000786
              34FF11A551FF0FA350FF0FA24CFF0E9F4BFF0E9E48FF0C9C46FF0A9A44FF0A99
              43FF089741FF089540FF07933EFF06913DFF027A2CFFFBF8F6FFFAF8F4FFF7F2
              EFFFECDFDAFFAB7E6DFFFFFFFFFFFFFEFEFFFFFDFCFFFEFCFAFFFCF9F7FFCAAF
              A6FF4C352D860000000600000002000000000000000000000000000000000564
              27BD088634FF078534FF078433FF068332FF068332FF068231FF058031FF0480
              30FF057E2FFF037E2EFF047C2EFF037C2DFF429C61FFFBF9F6FFFBF8F6FFF6F1
              EDFFEBDFDBFFB08574FFFFFEFEFFFEFBFAFFFDF9F7FFFCF6F3FFCEB2A8FF4F38
              3086000000080000000300000001000000000000000000000000000000000000
              0000000000000000000100000005E1CEC7FFFFFFFFFFFEFAF9FFFDFAFAFFFDFB
              F9FFFDFAF9FFFDFAF8FFFDFAF9FFFDF9F8FFFBF9F7FFFBF9F8FFF9F6F4FFF6F0
              ECFFECE1DBFFB68C7DFFFFFEFEFFFDF9F6FFFBF6F3FFD1B5ACFF533B33860000
              0008000000030000000100000000000000000000000000000000000000000000
              0000000000000000000100000004C49E90FFFFFFFFFFFDFBFAFFFDFBFAFFFDFB
              F9FFFDFBF9FFFEFAF9FFFDFAF9FFFDFAF8FFFDFAF8FFF9F7F6FFF9F4F2FFF5ED
              EBFFEBE1DDFFBC9584FFFFFEFEFFFBF6F3FFD4BAAFFF563F3685000000070000
              0003000000010000000000000000000000000000000000000000000000000000
              0000000000000000000100000004C5A190FFFFFFFFFFFEFDFBFFFDFBFBFFFDFD
              FBFFFEFBFAFFFEFBFAFFFDFBF9FFFDFBF9FFFBF7F6FFF9F5F3FFF7F1EEFFF3EB
              E7FFEDE1DCFFC19B8BFFFFFEFEFFD6BCB2FF59423A8400000006000000020000
              0001000000000000000000000000000000000000000000000000000000000000
              0000000000000000000100000003C6A191FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
              FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFBF9F9FFF9F6F4FFF6F1F0FFF2EC
              E9FFEEE3E0FFE4D2CBFFDBC5BDFF5A453E830000000500000002000000010000
              0000000000000000000000000000000000000000000000000000000000000000
              000000000000000000000000000293776CBEC6A291FFC6A192FFC6A191FFC59F
              91FFC69F92FFC59F91FFC59F90FFC59F91FFC49F90FFC49E90FFC49D8FFFC49E
              8EFFC39D8EFFC39D8EFF5D484182000000040000000200000001000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              0000000000000000000000000001000000010000000200000002000000030000
              0003000000030000000300000003000000030000000300000004000000040000
              0004000000040000000400000003000000010000000000000000000000000000
              0000000000000000000000000000000000000000000000000000}
            PaintStyle = bpsGlyph
            TabOrder = 3
            OnClick = btnExcelClick
          end
          object cMmCelda: TcxMemo
            Left = 172
            Top = 10
            Properties.WantReturns = False
            Style.BorderColor = clWindowFrame
            Style.BorderStyle = ebs3D
            Style.HotTrack = False
            TabOrder = 2
            OnEnter = cMmCeldaEnter
            OnExit = cMmCeldaExit
            OnKeyPress = cMmCeldaKeyPress
            Height = 89
            Width = 185
          end
          object dxLayoutControl3Group_Root: TdxLayoutGroup
            AlignHorz = ahClient
            AlignVert = avClient
            ButtonOptions.Buttons = <>
            Hidden = True
            LayoutDirection = ldHorizontal
            ShowBorder = False
            Index = -1
          end
          object dxLayoutControl3Item2: TdxLayoutItem
            Parent = dxLayoutControl3Group_Root
            CaptionOptions.Text = 'cxButton1'
            CaptionOptions.Visible = False
            Control = btnGuardar
            ControlOptions.ShowBorder = False
            Index = 0
          end
          object dxLayoutControl3Item3: TdxLayoutItem
            Parent = dxLayoutControl3Group_Root
            CaptionOptions.Text = 'cxButton1'
            CaptionOptions.Visible = False
            Control = btnCancelar
            ControlOptions.ShowBorder = False
            Index = 1
          end
          object dxLayoutControl3Item1: TdxLayoutItem
            Parent = dxLayoutControl3Group_Root
            CaptionOptions.Visible = False
            Control = btnExcel
            ControlOptions.ShowBorder = False
            Index = 3
          end
          object dxLayoutControl3Item4: TdxLayoutItem
            Parent = dxLayoutControl3Group_Root
            AlignHorz = ahClient
            AlignVert = avClient
            Control = cMmCelda
            ControlOptions.ShowBorder = False
            Index = 2
          end
        end
      end
      object GBx4: TcxGroupBox
        Left = 0
        Top = 105
        Align = alClient
        Caption = 'GBx4'
        TabOrder = 1
        Height = 459
        Width = 1143
        object SprShBkDatos: TcxSpreadSheetBook
          Left = 2
          Top = 18
          Width = 1139
          Height = 439
          BufferedPaint = True
          RowsAutoHeight = False
          Align = alClient
          DefaultStyle.Font.Name = 'Tahoma'
          DefaultStyle.WordBreak = True
          HeaderFont.Charset = DEFAULT_CHARSET
          HeaderFont.Color = clWindowText
          HeaderFont.Height = -11
          HeaderFont.Name = 'Tahoma'
          HeaderFont.Style = []
          PainterType = ptCustom
          PageCount = 4
          OnActiveCellChanging = SprShBkDatosActiveCellChanging
        end
      end
    end
  end
  object QActa: TZQuery
    Connection = connection.zConnection
    AfterScroll = QActaAfterScroll
    SQL.Strings = (
      'select * from acta_entrega where sContrato=:Contrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
    Left = 288
    Top = 64
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
  end
  object dsActa: TDataSource
    AutoEdit = False
    DataSet = QActa
    Left = 376
    Top = 72
  end
  object QrFolios: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select sContrato,sIdFolio,sNumeroOrden,sUbicacion from ordenesde' +
        'trabajo where sContrato=:Contrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
    Left = 272
    Top = 240
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
  end
  object dsFolios: TDataSource
    DataSet = QrFolios
    Left = 224
    Top = 240
  end
  object RptActa: TfrxReport
    Version = '4.10.3'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Por defecto'
    PrintOptions.PrintOnSheet = 0
    ReportOptions.CreateDate = 42460.581254791700000000
    ReportOptions.LastChange = 42509.418173009260000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'procedure Memo126OnBeforePrint(Sender: TfrxComponent);'
      'var'
      '       Suma,resH,resM:Double;'
      '       tmpHrs:String;'
      '       Hrs:string;'
      '       Min:String;                               '
      '                      '
      'begin'
      '       suma:=0;                           '
      
        '      // Suma:=(SUM((iif(<Td_Distribucion_detalle2."sHoraFinal">' +
        '='#39'24:00'#39',1,StrToTime(<Td_Distribucion_detalle2."sHoraFinal">))-i' +
        'if(<Td_Distribucion_detalle2."sHoraFinal">='#39'00:00'#39',0, StrToTime(' +
        '<Td_Distribucion_detalle2."sHoraInicio">))),DetailData2));      ' +
        '         '
      '      '
      
        '     // Suma:=(SUM((iif(<Td_Distribucion_detalle2."sHoraFinal">=' +
        #39'24:00'#39',1, StrToTime(<Td_Distribucion_detalle2."sHoraFinal">) ) ' +
        '-    iif(<Td_Distribucion_detalle2."sHoraInicio">='#39'00:00'#39',0,StrT' +
        'oTime(<Td_Distribucion_detalle2."sHoraInicio">) )          ),Det' +
        'ailData2));                  '
      '      '
      
        '     //  Suma:=(SUM((iif(<Td_Distribucion_detalle2."sHoraFinal">' +
        '='#39'24:00'#39',1, StrToTime(<Td_Distribucion_detalle2."sHoraFinal">) )' +
        ' -    iif(<Td_Distribucion_detalle2."sHoraInicio">='#39'00:00'#39',0,Str' +
        'ToTime(<Td_Distribucion_detalle2."sHoraInicio">) )          ),De' +
        'tailData2));                  '
      ''
      
        '        Suma:=(SUM((iif(<Td_Distribucion_detalle2."sIdClasificac' +
        'ion">='#39'NOTA'#39',0,iif(<Td_Distribucion_detalle2."sHoraFinal">='#39'24:0' +
        '0'#39',1, StrToTime(<Td_Distribucion_detalle2."sHoraFinal">) ) -    ' +
        'iif(<Td_Distribucion_detalle2."sHoraInicio">='#39'00:00'#39',0,StrToTime' +
        '(<Td_Distribucion_detalle2."sHoraInicio">) ))),DetailData2));   ' +
        '           '
      '      // showmessage(FloatToStr(Suma));'
      '       '
      '       '
      '       '
      '       '
      '       '
      '       resH:=trunc(suma*24);'
      ''
      '      // showmessage(Hrs);                     '
      
        '       resM:=(suma - (trunc(suma*24)/24) ) * 1440;              ' +
        '                     '
      '       if round(resM)=60 then'
      '       begin'
      '        resH:=resH + 1;  '
      '        Min:='#39'00'#39';                '
      '       end'
      '       else'
      '                Min:=floatToStr(round(resM));'
      '      '
      ''
      ''
      '       Hrs:=floatToStr(resH);'
      '       if length(Hrs)=1 then'
      '        Hrs:='#39'0'#39'+Hrs;'
      ''
      '       if length(Min)=1 then'
      '        Min:='#39'0'#39'+Min;  '
      '                            '
      
        '      Memo126.text:=Hrs +'#39':'#39' +min;                              ' +
        ' '
      'end;'
      ''
      'begin'
      ''
      'end.')
    OnGetValue = RptActaGetValue
    OnReportPrint = 'no '
    Left = 392
    Top = 240
    Datasets = <
      item
        DataSet = frmDiarioTurno2.Td_duracion
        DataSetName = 'ds_Duracion'
      end
      item
        DataSet = frmReportePeriodo.dsConfiguracion
        DataSetName = 'dsConfiguracion'
      end
      item
        DataSet = frmDiarioTurno2.dsEmbarcacion
        DataSetName = 'dsEmbarcacion'
      end
      item
        DataSet = frmDiarioTurno2.Td_AvancesPartidas
        DataSetName = 'Td_AvancesPartidas'
      end
      item
        DataSet = frmDiarioTurno2.Td_contrato
        DataSetName = 'Td_contrato'
      end
      item
        DataSet = frmDiarioTurno2.Td_Distribucion_detalle
        DataSetName = 'Td_Distribucion_detalle'
      end
      item
        DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
        DataSetName = 'Td_Distribucion_detalle2'
      end
      item
        DataSet = frmDiarioTurno2.Td_movFolios
        DataSetName = 'Td_movFolios'
      end
      item
        DataSet = frmDiarioTurno2.Td_partidas
        DataSetName = 'Td_partidas'
      end
      item
        DataSet = frmDiarioTurno2.Td_PartidasAnexo
        DataSetName = 'Td_PartidasAnexo'
      end
      item
        DataSet = frmDiarioTurno2.Td_resumenMaterial
        DataSetName = 'Td_resumenMaterial'
      end
      item
        DataSet = frmDiarioTurno2.Td_ResumenPersonal
        DataSetName = 'Td_ResumenPersonal'
      end>
    Variables = <>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      Orientation = poLandscape
      PaperWidth = 279.400000000000000000
      PaperHeight = 215.900000000000000000
      PaperSize = 1
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object MasterData1: TfrxMasterData
        Height = 22.677180000000000000
        Top = 495.118430000000000000
        Width = 980.410082000000000000
        DataSet = frmDiarioTurno2.Td_partidas
        DataSetName = 'Td_partidas'
        RowCount = 0
        Stretched = True
        object Memo19: TfrxMemoView
          Width = 83.149660000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftLeft]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."sCsu"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo21: TfrxMemoView
          Left = 83.488188980000000000
          Width = 105.826771650000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_partidas
          DataSetName = 'Td_partidas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."sIdFolio"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo23: TfrxMemoView
          Left = 189.692913390000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."Plataforma"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo24: TfrxMemoView
          Left = 283.803149610000000000
          Width = 94.488188980000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."Ubicacion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo25: TfrxMemoView
          Left = 378.291338580000000000
          Width = 245.669320630000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft]
          HAlign = haBlock
          Memo.UTF8W = (
            '[Td_partidas."mDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo26: TfrxMemoView
          Left = 624.606680000000000000
          Width = 60.472480000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DisplayFormat.FormatStr = 'dd-mmm-yy'
          DisplayFormat.Kind = fkDateTime
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."dFiProgramado"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo27: TfrxMemoView
          Left = 684.638220000000000000
          Width = 60.472436060000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DisplayFormat.FormatStr = 'dd-mmm-yy'
          DisplayFormat.Kind = fkDateTime
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."dFfProgramado"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo30: TfrxMemoView
          Left = 744.906000000000000000
          Width = 64.252010000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight]
          HAlign = haCenter
          Memo.UTF8W = (
            '[FormatFloat('#39'0.00'#39',<Td_partidas."Avance">)]%')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo31: TfrxMemoView
          Left = 807.921770000000000000
          Width = 170.078850000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_MovFolios."mObservaciones"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object PageFooter1: TfrxPageFooter
        Height = 22.677180000000000000
        Top = 793.701300000000000000
        Width = 980.410082000000000000
        object Memo1: TfrxMemoView
          Left = 642.520100000000000000
          Width = 75.590600000000000000
          Height = 18.897650000000000000
          ShowHint = False
          HAlign = haRight
          Memo.UTF8W = (
            '[Page#]')
        end
      end
      object PageHeader3: TfrxPageHeader
        Height = 336.378196850000000000
        Top = 18.897650000000000000
        Width = 980.410082000000000000
        object Picture5: TfrxPictureView
          Left = 814.598950000000000000
          Top = 4.779530000000000000
          Width = 143.622091180000000000
          Height = 68.031513150000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmReportePeriodo.dsConfiguracion
          DataSetName = 'dsConfiguracion'
          Frame.Color = clNone
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Picture6: TfrxPictureView
          Left = 3.779530000000000000
          Width = 136.063006770000000000
          Height = 49.133863150000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmDiarioTurno2.Td_contrato
          DataSetName = 'Td_contrato'
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Memo3: TfrxMemoView
          Left = 140.842610000000000000
          Top = 37.456710000000000000
          Width = 661.417750000000000000
          Height = 37.795300000000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            'ACTA DE ENTREGA'
            'DE ACTIVIDADES REALIZADAS')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo2: TfrxMemoView
          Left = 139.842610000000000000
          Top = 2.779530000000000000
          Width = 661.417750000000000000
          Height = 37.795300000000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            'GERENCIA DE CONFIABILIDAD DE INSTALACIONES MARINAS'
            'COORDINACION DE MANTENIMIENTO INTEGRAL ABKATUN-POOL-CHUC')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo10: TfrxMemoView
          Left = 605.189027640000000000
          Top = 266.960764170000000000
          Width = 122.834672520000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          HAlign = haRight
          Memo.UTF8W = (
            'FECHA DE ACTA:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo11: TfrxMemoView
          Left = 758.259940160000000000
          Top = 254.842644170000000000
          Width = 211.653616540000000000
          Height = 28.346466460000000000
          ShowHint = False
          DisplayFormat.FormatStr = 'dd "de" mmmm "de" yyyy'
          DisplayFormat.Kind = fkDateTime
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_MovFolios."dFecha"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo12: TfrxMemoView
          Left = 5.000000000000000000
          Top = 99.252044170000000000
          Width = 142.110236220000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            'CONTRATO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo13: TfrxMemoView
          Left = 147.291745530000000000
          Top = 91.252044170000000000
          Width = 438.425235910000000000
          Height = 20.787406460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            '[dsConfiguracion."sContratoBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo20: TfrxMemoView
          Left = 5.000000000000000000
          Top = 112.370078740000000000
          Width = 142.110236220000000000
          Height = 39.685039370000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            'OBRA:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo22: TfrxMemoView
          Left = 147.110236220000000000
          Top = 113.370078740000000000
          Width = 438.425196850000000000
          Height = 43.464569370000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            '[dsConfiguracion."mDescripcionBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo28: TfrxMemoView
          Left = 5.000000000000000000
          Top = 164.401567480000000000
          Width = 142.110236220000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            'ORDEN DE TRABAJO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo29: TfrxMemoView
          Left = 605.102362200000000000
          Top = 169.385797480000000000
          Width = 122.834645670000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          HAlign = haRight
          Memo.UTF8W = (
            'No DE ACTA:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo35: TfrxMemoView
          Left = 147.110230690000000000
          Top = 164.401567480000000000
          Width = 438.425196850000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            '[td_contrato."sLabelContrato"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo38: TfrxMemoView
          Left = 752.968650390000000000
          Top = 139.149557480000000000
          Width = 214.677165350000000000
          Height = 69.921296460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_MovFolios."sNoActa"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo39: TfrxMemoView
          Left = 5.000000000000000000
          Top = 244.629913940000000000
          Width = 142.110236220000000000
          Height = 17.007876460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            'ESPECIALIDAD:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo40: TfrxMemoView
          Left = 147.110219130000000000
          Top = 244.629913940000000000
          Width = 438.425196850000000000
          Height = 24.566936460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            '[Td_MovFolios."sEspecialidad"]')
          ParentFont = False
        end
        object Memo16: TfrxMemoView
          Left = 147.110236220000000000
          Top = 185.063016540000000000
          Width = 438.425196850000000000
          Height = 51.023629370000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haBlock
          Memo.UTF8W = (
            '[dsConfiguracion."mDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo41: TfrxMemoView
          Left = 5.000000000000000000
          Top = 282.874150000000000000
          Width = 142.110236220000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            'CENTRO DE PROCESO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo42: TfrxMemoView
          Left = 147.110219130000000000
          Top = 279.094620000000000000
          Width = 438.425196850000000000
          Height = 24.566936460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftBottom]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            '[Td_partidas."Plataforma"]')
          ParentFont = False
        end
        object Shape1: TfrxShapeView
          Top = 77.771800000000000000
          Width = 978.898270000000000000
          Height = 241.889920000000000000
          ShowHint = False
          Frame.Color = 14211288
          Frame.Style = fsDouble
        end
        object Memo217: TfrxMemoView
          Left = 798.535870000000000000
          Top = 89.929190000000000000
          Width = 107.338582680000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          HAlign = haRight
          Memo.UTF8W = (
            'ACTA PARCIAL')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo218: TfrxMemoView
          Left = 798.535870000000000000
          Top = 112.606370000000000000
          Width = 107.338582680000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Width = 0.500000000000000000
          HAlign = haRight
          Memo.UTF8W = (
            'ACTA TOTAL')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo219: TfrxMemoView
          Left = 925.984850000000000000
          Top = 86.929190000000000000
          Width = 41.574830000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Calibri'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Highlight.Font.Charset = DEFAULT_CHARSET
          Highlight.Font.Color = clWhite
          Highlight.Font.Height = -13
          Highlight.Font.Name = 'Calibri'
          Highlight.Font.Style = []
          Highlight.Condition = '<Td_movFolios."eTipo">='#39'Total'#39
          Memo.UTF8W = (
            'X')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo220: TfrxMemoView
          Left = 925.984850000000000000
          Top = 108.826840000000000000
          Width = 41.574830000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Calibri'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Highlight.Font.Charset = DEFAULT_CHARSET
          Highlight.Font.Color = clWhite
          Highlight.Font.Height = -13
          Highlight.Font.Name = 'Calibri'
          Highlight.Font.Style = []
          Highlight.Condition = '<Td_movFolios."eTipo">='#39'Parcial'#39
          Memo.UTF8W = (
            'X')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupHeader1: TfrxGroupHeader
        Height = 56.692950000000000000
        Top = 415.748300000000000000
        Width = 980.410082000000000000
        Condition = 'Td_partidas."sContrato"'
        object Memo4: TfrxMemoView
          Top = 7.559060000000000000
          Width = 83.149660000000000000
          Height = 49.133890000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'No. SIPOM '
            'O CSU')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo5: TfrxMemoView
          Left = 83.488188980000000000
          Top = 7.559060000000000000
          Width = 105.826840000000000000
          Height = 49.133890000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PARTIDA/FOLIO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo6: TfrxMemoView
          Left = 189.771800000000000000
          Top = 7.559060000000000000
          Width = 94.488250000000000000
          Height = 49.133890000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PLATAFORMA')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo7: TfrxMemoView
          Left = 283.921460000000000000
          Top = 7.559060000000000000
          Width = 94.488188980000000000
          Height = 49.133890000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'UBICACI'#211'N')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo8: TfrxMemoView
          Left = 378.291590000000000000
          Top = 7.559060000000000000
          Width = 245.669450000000000000
          Height = 49.133890000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'DESCRIPCI'#211'N')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo9: TfrxMemoView
          Left = 624.606680000000000000
          Top = 7.559060000000000000
          Width = 120.188976380000000000
          Height = 26.456710000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PERIODO DE '
            'EJECUCION')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo14: TfrxMemoView
          Left = 624.606680000000000000
          Top = 34.015770000000000000
          Width = 60.472480000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'INICIO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo15: TfrxMemoView
          Left = 684.637795280000000000
          Top = 34.015770000000000000
          Width = 60.472436060000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'T'#201'RMINO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo17: TfrxMemoView
          Left = 744.906000000000000000
          Top = 7.559060000000000000
          Width = 64.252010000000000000
          Height = 49.133890000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'AVANCE %')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo18: TfrxMemoView
          Left = 807.921770000000000000
          Top = 7.559060000000000000
          Width = 170.078850000000000000
          Height = 49.133890000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'OBSERVACIONES Y/O OFICIOS DE REFERENCIA')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupFooter1: TfrxGroupFooter
        Height = 192.756030000000000000
        Top = 540.472790000000000000
        Width = 980.410082000000000000
        object Memo32: TfrxMemoView
          Width = 978.898270000000000000
          Height = 11.338590000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Style = fsDouble
          Frame.Typ = [ftTop]
          HAlign = haCenter
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo191: TfrxMemoView
          Top = 37.795300000000000000
          Width = 980.409605040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'RESUMEN DE COSTOS')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo194: TfrxMemoView
          Top = 56.692950000000000000
          Width = 980.409605040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ANEXOS "C"')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo195: TfrxMemoView
          Top = 75.590600000000000000
          Width = 201.826425040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ANEXO C')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo199: TfrxMemoView
          Left = 200.315090000000000000
          Top = 75.590600000000000000
          Width = 632.692845040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'DESCRIPCION')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo200: TfrxMemoView
          Left = 831.496600000000000000
          Top = 75.590600000000000000
          Width = 148.913005040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'IMPORTE TOTAL')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo201: TfrxMemoView
          Top = 94.488250000000000000
          Width = 201.826425040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          HAlign = haCenter
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo202: TfrxMemoView
          Left = 200.315090000000000000
          Top = 94.488250000000000000
          Width = 632.692845040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo203: TfrxMemoView
          Left = 831.496600000000000000
          Top = 94.488250000000000000
          Width = 73.322405040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'M.N.')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo204: TfrxMemoView
          Left = 907.087200000000000000
          Top = 94.488250000000000000
          Width = 73.322405040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'U.S.D.')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo205: TfrxMemoView
          Top = 113.385900000000000000
          Width = 201.826425040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Memo.UTF8W = (
            'ANEXO C-5 "PERSONAL OPT."')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo206: TfrxMemoView
          Left = 204.094620000000000000
          Top = 113.385900000000000000
          Width = 628.913315040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Memo.UTF8W = (
            'PERSONAL')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo207: TfrxMemoView
          Left = 831.496600000000000000
          Top = 113.385900000000000000
          Width = 73.322405040000000000
          Height = 18.897650000000000000
          ShowHint = False
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[PERSONALMN]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo208: TfrxMemoView
          Left = 907.087200000000000000
          Top = 113.385900000000000000
          Width = 73.322405040000000000
          Height = 18.897650000000000000
          ShowHint = False
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[PERSONALDLL]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo209: TfrxMemoView
          Top = 132.283550000000000000
          Width = 201.826425040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Memo.UTF8W = (
            'ANEXO C-5 "EQUIPO OPT."')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo210: TfrxMemoView
          Left = 204.094620000000000000
          Top = 132.283550000000000000
          Width = 628.913315040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Memo.UTF8W = (
            'EQUIPO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo211: TfrxMemoView
          Left = 831.496600000000000000
          Top = 132.283550000000000000
          Width = 73.322405040000000000
          Height = 18.897650000000000000
          ShowHint = False
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[EQUIPOMN]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo212: TfrxMemoView
          Left = 907.087200000000000000
          Top = 132.283550000000000000
          Width = 73.322405040000000000
          Height = 18.897650000000000000
          ShowHint = False
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[EQUIPODLL]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo213: TfrxMemoView
          Top = 151.181200000000000000
          Width = 201.826425040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          HAlign = haCenter
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo214: TfrxMemoView
          Left = 200.315090000000000000
          Top = 151.181200000000000000
          Width = 632.692845040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            'TOTALES')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo215: TfrxMemoView
          Left = 831.496600000000000000
          Top = 151.181200000000000000
          Width = 73.322405040000000000
          Height = 18.897650000000000000
          ShowHint = False
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[<PERSONALMN> + <EQUIPOMN>]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo216: TfrxMemoView
          Left = 907.087200000000000000
          Top = 151.181200000000000000
          Width = 73.322405040000000000
          Height = 18.897650000000000000
          ShowHint = False
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[<PERSONALDLL> + <EQUIPODLL>]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
    end
    object Page6: TfrxReportPage
      PaperWidth = 215.900000000000000000
      PaperHeight = 279.400000000000000000
      PaperSize = 1
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object PageHeader1: TfrxPageHeader
        Height = 253.228536850000000000
        Top = 18.897650000000000000
        Width = 740.409927000000000000
        OnAfterPrint = 'PageHeader1OnAfterPrint'
        object Picture1: TfrxPictureView
          Left = 591.606680000000000000
          Top = 4.779530000000000000
          Width = 143.622091180000000000
          Height = 45.354333150000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmReportePeriodo.dsConfiguracion
          DataSetName = 'dsConfiguracion'
          Frame.Color = clNone
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Picture2: TfrxPictureView
          Left = 3.779530000000000000
          Width = 136.063006770000000000
          Height = 49.133863150000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmDiarioTurno2.Td_contrato
          DataSetName = 'Td_contrato'
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Memo33: TfrxMemoView
          Left = 140.842610000000000000
          Width = 449.764070000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sNombre"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo34: TfrxMemoView
          Left = 139.842610000000000000
          Top = 27.456710000000000000
          Width = 449.764070000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            'RESUMEN DE COSTOS POR ACTIVIDADES')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo36: TfrxMemoView
          Left = 337.133858270000000000
          Top = 66.031540000000000000
          Width = 192.755905510000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'EMBARCACION:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo37: TfrxMemoView
          Left = 528.755905510000000000
          Top = 66.031540000000000000
          Width = 210.519685040000000000
          Height = 13.228346460000000000
          OnBeforePrint = 'Memo7OnBeforePrint'
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsEmbarcacion."sDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo45: TfrxMemoView
          Top = 66.252044170000000000
          Width = 142.110236220000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'CONTRATO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo46: TfrxMemoView
          Left = 142.291745530000000000
          Top = 66.252044170000000000
          Width = 195.779527560000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sContratoBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo70: TfrxMemoView
          Top = 79.370078740000000000
          Width = 142.110236220000000000
          Height = 39.685039370000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'DESCRIPCI'#211'N:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo72: TfrxMemoView
          Left = 141.732283460000000000
          Top = 79.370078740000000000
          Width = 597.543307090000000000
          Height = 39.685039370000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            '[dsConfiguracion."mDescripcionBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo179: TfrxMemoView
          Top = 118.677148270000000000
          Width = 142.110236220000000000
          Height = 39.685056460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'ORDEN DE TRABAJO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo1028: TfrxMemoView
          Top = 207.937034720000000000
          Width = 142.110236220000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'PERIODO DE EJECUCI'#211'N:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo1081: TfrxMemoView
          Left = 142.110219130000000000
          Top = 208.047244090000000000
          Width = 195.779478740000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."dFiProgramado"] AL [Td_partidas."dFfProgramado"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo1082: TfrxMemoView
          Left = 141.732107720000000000
          Top = 119.165366540000000000
          Width = 597.543485280000000000
          Height = 39.307086610000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haBlock
          Memo.UTF8W = (
            
              '[td_contrato."sLabelContrato"]. -[dsConfiguracion."mDescripcion"' +
              ']')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo43: TfrxMemoView
          Top = 158.740260000000000000
          Width = 142.110236220000000000
          Height = 20.787406460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'PART./FOLIO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo44: TfrxMemoView
          Left = 141.732107720000000000
          Top = 159.228478270000000000
          Width = 597.543307090000000000
          Height = 48.377952760000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haBlock
          Memo.UTF8W = (
            '[Td_partidas."mDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo50: TfrxMemoView
          Left = 337.157700000000000000
          Top = 208.047244090000000000
          Width = 192.756030000000000000
          Height = 28.346456690000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'INSTALACION: [Td_partidas."Plataforma"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo55: TfrxMemoView
          Left = 528.756186220000000000
          Top = 208.047244090000000000
          Width = 210.519685040000000000
          Height = 28.346456690000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ACTIVO: [Td_partidas."Ubicacion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo56: TfrxMemoView
          Top = 179.417440000000000000
          Width = 142.110236220000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."sIdFolio"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object PageFooter2: TfrxPageFooter
        Height = 22.677180000000000000
        Top = 684.094930000000000000
        Width = 740.409927000000000000
      end
      object GroupHeader2: TfrxGroupHeader
        Height = 22.677180000000000000
        Top = 332.598640000000000000
        Width = 740.409927000000000000
        Condition = 'Td_AvancesPartidas."sNumeroActividad"'
        ReprintOnNewPage = True
        object Memo47: TfrxMemoView
          Top = 3.779530000000000000
          Width = 738.519685040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ACTIVIDAD')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object MasterData2: TfrxMasterData
        Height = 22.677180000000000000
        Top = 377.953000000000000000
        Width = 740.409927000000000000
        DataSet = frmDiarioTurno2.Td_AvancesPartidas
        DataSetName = 'Td_AvancesPartidas'
        RowCount = 0
        Stretched = True
        object Memo48: TfrxMemoView
          Width = 113.385900000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_AvancesPartidas
          DataSetName = 'Td_AvancesPartidas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_AvancesPartidas."sNumeroActividad"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo49: TfrxMemoView
          Left = 113.385900000000000000
          Width = 625.133858270000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_AvancesPartidas
          DataSetName = 'Td_AvancesPartidas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Memo.UTF8W = (
            '[Td_AvancesPartidas."mDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupFooter2: TfrxGroupFooter
        Height = 71.811070000000000000
        Top = 551.811380000000000000
        Width = 740.409927000000000000
        AllowSplit = True
        object Memo59: TfrxMemoView
          Left = 264.567100000000000000
          Top = 36.133890000000000000
          Width = 283.464750000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            'COSTO TOTAL DE LA ACTIVIDAD:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo60: TfrxMemoView
          Left = 548.031850000000000000
          Top = 35.811070000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle
          DataSetName = 'Td_Distribucion_detalle'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[SUM(<Td_Distribucion_detalle."dImporteMN">,DetailData1)]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo61: TfrxMemoView
          Left = 642.520100000000000000
          Top = 35.811070000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle
          DataSetName = 'Td_Distribucion_detalle'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[SUM(<Td_Distribucion_detalle."dImporteDll">,DetailData1)]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo62: TfrxMemoView
          Left = 548.031850000000000000
          Top = 17.236240000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop]
          HAlign = haRight
          Memo.UTF8W = (
            'IMP MN')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo63: TfrxMemoView
          Left = 642.520100000000000000
          Top = 17.236240000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop]
          HAlign = haRight
          Memo.UTF8W = (
            'IMP USD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo222: TfrxMemoView
          Width = 314.456722200000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Memo.UTF8W = (
            'NOTA: PARA ESTA ACTIVIDAD NO APLICAN MATERIALES')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object DetailData1: TfrxDetailData
        Height = 22.677180000000000000
        Top = 472.441250000000000000
        Width = 740.409927000000000000
        DataSet = frmDiarioTurno2.Td_Distribucion_detalle
        DataSetName = 'Td_Distribucion_detalle'
        RowCount = 0
        Stretched = True
        object Memo51: TfrxMemoView
          Width = 343.937230000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle
          DataSetName = 'Td_Distribucion_detalle'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          Memo.UTF8W = (
            '[Td_Distribucion_detalle."sDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo52: TfrxMemoView
          Left = 343.937230000000000000
          Width = 204.094620000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle
          DataSetName = 'Td_Distribucion_detalle'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_Distribucion_detalle."sImporte"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo53: TfrxMemoView
          Left = 548.031850000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle
          DataSetName = 'Td_Distribucion_detalle'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_Distribucion_detalle."dImporteMn"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo54: TfrxMemoView
          Left = 642.520100000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle
          DataSetName = 'Td_Distribucion_detalle'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_Distribucion_detalle."dImporteDll"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupHeader3: TfrxGroupHeader
        Height = 26.456710000000000000
        Top = 423.307360000000000000
        Width = 740.409927000000000000
        Condition = 'Td_Distribucion_detalle."sNumeroActividad"'
        object Memo57: TfrxMemoView
          Left = 548.031850000000000000
          Top = 7.559060000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop]
          HAlign = haRight
          Memo.UTF8W = (
            'IMP MN')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo58: TfrxMemoView
          Left = 642.520100000000000000
          Top = 7.559060000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop]
          HAlign = haRight
          Memo.UTF8W = (
            'IMP USD')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupFooter3: TfrxGroupFooter
        Height = 11.338590000000000000
        Top = 517.795610000000000000
        Width = 740.409927000000000000
      end
    end
    object Page3: TfrxReportPage
      PaperWidth = 215.900000000000000000
      PaperHeight = 279.400000000000000000
      PaperSize = 1
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object PageHeader2: TfrxPageHeader
        Height = 245.669476850000000000
        Top = 18.897650000000000000
        Width = 740.409927000000000000
        OnAfterPrint = 'PageHeader1OnAfterPrint'
        object Picture3: TfrxPictureView
          Left = 591.606680000000000000
          Top = 4.779530000000000000
          Width = 143.622091180000000000
          Height = 45.354333150000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmReportePeriodo.dsConfiguracion
          DataSetName = 'dsConfiguracion'
          Frame.Color = clNone
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Picture4: TfrxPictureView
          Left = 3.779530000000000000
          Width = 136.063006770000000000
          Height = 49.133863150000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmDiarioTurno2.Td_contrato
          DataSetName = 'Td_contrato'
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Memo65: TfrxMemoView
          Left = 139.842610000000000000
          Top = 27.590551181102400000
          Width = 449.764070000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            'NOTA DE CAMPO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo66: TfrxMemoView
          Left = 337.133858270000000000
          Top = 66.031540000000000000
          Width = 192.755905510000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'EMBARCACION:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo67: TfrxMemoView
          Left = 528.755905510000000000
          Top = 66.031540000000000000
          Width = 210.519685040000000000
          Height = 13.228346460000000000
          OnBeforePrint = 'Memo7OnBeforePrint'
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsEmbarcacion."sDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo68: TfrxMemoView
          Top = 66.252044170000000000
          Width = 142.110236220000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'CONTRATO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo69: TfrxMemoView
          Left = 142.291745530000000000
          Top = 66.252044170000000000
          Width = 195.779527560000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sContratoBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo71: TfrxMemoView
          Top = 79.370078740000000000
          Width = 142.110236220000000000
          Height = 39.685039370000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'DESCRIPCI'#211'N:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo73: TfrxMemoView
          Left = 141.732283460000000000
          Top = 79.370078740000000000
          Width = 597.543307090000000000
          Height = 39.685039370000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            '[dsConfiguracion."mDescripcionBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo74: TfrxMemoView
          Top = 118.677148270000000000
          Width = 142.110236220000000000
          Height = 39.685056460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'ORDEN DE TRABAJO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo75: TfrxMemoView
          Top = 207.937034720000000000
          Width = 142.110236220000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'PERIODO DE EJECUCI'#211'N:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo76: TfrxMemoView
          Left = 142.110219130000000000
          Top = 208.047244090000000000
          Width = 195.779478740000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."dFiProgramado"] AL [Td_partidas."dFfProgramado"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo77: TfrxMemoView
          Left = 141.732107720000000000
          Top = 119.165366540000000000
          Width = 597.543485280000000000
          Height = 39.307086610000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haBlock
          Memo.UTF8W = (
            
              '[td_contrato."sLabelContrato"]. -[dsConfiguracion."mDescripcion"' +
              ']')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo78: TfrxMemoView
          Top = 158.740260000000000000
          Width = 142.110236220000000000
          Height = 20.787406460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'PART./FOLIO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo79: TfrxMemoView
          Left = 141.732107720000000000
          Top = 159.228478270000000000
          Width = 597.543307090000000000
          Height = 48.377952760000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haBlock
          Memo.UTF8W = (
            '[Td_partidas."mDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo80: TfrxMemoView
          Left = 337.157700000000000000
          Top = 208.047244090000000000
          Width = 192.756030000000000000
          Height = 28.346456690000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'INSTALACION: [Td_partidas."Plataforma"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo81: TfrxMemoView
          Left = 528.756186220000000000
          Top = 208.047244090000000000
          Width = 210.519685040000000000
          Height = 28.346456690000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ACTIVO: [Td_partidas."Ubicacion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo82: TfrxMemoView
          Top = 179.417440000000000000
          Width = 142.110236220000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."sIdFolio"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo64: TfrxMemoView
          Left = 140.842610000000000000
          Width = 449.764070000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sNombre"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object PageFooter3: TfrxPageFooter
        Height = 22.677180000000000000
        Top = 990.236860000000000000
        Width = 740.409927000000000000
      end
      object GroupHeader4: TfrxGroupHeader
        Height = 22.677180000000000000
        Top = 325.039580000000000000
        Width = 740.409927000000000000
        Condition = 'Td_PartidasAnexo."sNumeroActividad"'
        PrintChildIfInvisible = True
        object Memo83: TfrxMemoView
          Left = 154.960730000000000000
          Top = 3.779530000000000000
          Width = 583.558955040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ACTIVIDAD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo123: TfrxMemoView
          Top = 3.779530000000000000
          Width = 154.204751260000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PARTIDA')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object MasterData3: TfrxMasterData
        Height = 22.677180000000000000
        Top = 370.393940000000000000
        Width = 740.409927000000000000
        DataSet = frmDiarioTurno2.Td_PartidasAnexo
        DataSetName = 'Td_PartidasAnexo'
        PrintChildIfInvisible = True
        RowCount = 0
        Stretched = True
        object Memo84: TfrxMemoView
          Width = 154.960730000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_AvancesPartidas
          DataSetName = 'Td_AvancesPartidas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_PartidasAnexo."sNumeroActividad"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo85: TfrxMemoView
          Left = 154.960730000000000000
          Width = 583.559028270000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_AvancesPartidas
          DataSetName = 'Td_AvancesPartidas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Memo.UTF8W = (
            '[Td_PartidasAnexo."mDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupFooter4: TfrxGroupFooter
        Height = 22.677180000000000000
        Top = 801.260360000000000000
        Width = 740.409927000000000000
        PrintChildIfInvisible = True
        object Memo127: TfrxMemoView
          Left = 416.503937007874000000
          Width = 182.551181102362000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'IMPORTE [Td_resumenMaterial."sTipo"]:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo128: TfrxMemoView
          Left = 599.055118110000000000
          Width = 71.433070870000000000
          Height = 18.897650000000000000
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[SUM(<Td_resumenMaterial."dImporteMn">,DetailData3)]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo129: TfrxMemoView
          Left = 670.866141730000000000
          Width = 67.275590550000000000
          Height = 18.897650000000000000
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[SUM(<Td_resumenMaterial."dImporteDll">,DetailData3)]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object DetailData2: TfrxDetailData
        Height = 18.897650000000000000
        Top = 525.354670000000000000
        Width = 740.409927000000000000
        DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
        DataSetName = 'Td_Distribucion_detalle2'
        RowCount = 0
        Stretched = True
        object Memo87: TfrxMemoView
          Width = 111.118110236220000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataField = 'dIdFecha'
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_Distribucion_detalle2."dIdFecha"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo88: TfrxMemoView
          Left = 111.488250000000000000
          Width = 75.590551180000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_Distribucion_detalle2."sHoraInicio"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo89: TfrxMemoView
          Left = 186.960730000000000000
          Width = 75.590551180000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_Distribucion_detalle2."sHoraFinal"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo90: TfrxMemoView
          Left = 263.433210000000000000
          Width = 109.606370000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_Distribucion_detalle2."sIdClasificacion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo91: TfrxMemoView
          Left = 372.803340000000000000
          Width = 83.149660000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = 'hh:mm'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            
              '[ IIf((<Td_Distribucion_detalle2."sHoraFinal">='#39'24:00'#39') and ( <T' +
              'd_Distribucion_detalle2."sHoraInicio">='#39'00:00'#39'),'#39'24:00'#39',FormatDa' +
              'teTime('#39'hh:mm'#39',(iif(<Td_Distribucion_detalle2."sHoraFinal">='#39'24:' +
              '00'#39',1,StrToTime(<Td_Distribucion_detalle2."sHoraFinal">)) - iif(' +
              '<Td_Distribucion_detalle2."sHoraInicio">='#39'00:00'#39',0,StrToTime(<Td' +
              '_Distribucion_detalle2."sHoraInicio">)))))]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo92: TfrxMemoView
          Left = 456.393940000000000000
          Width = 94.488188980000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n"%"'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            
              '[FormatFloat('#39'0.00'#39',<Td_Distribucion_detalle2."AvanceAnterior">)' +
              ']%')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo101: TfrxMemoView
          Left = 550.543600000000000000
          Width = 94.488188980000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[FormatFloat('#39'0.00'#39',<Td_Distribucion_detalle2."dAvance">)]%')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo102: TfrxMemoView
          Left = 644.693260000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            
              '[FormatFloat('#39'0.00'#39',<Td_Distribucion_detalle2."dAvance"> + <Td_D' +
              'istribucion_detalle2."AvanceAnterior">)]%')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupHeader5: TfrxGroupHeader
        Height = 86.929190000000000000
        Top = 415.748300000000000000
        Width = 740.409927000000000000
        Condition = 'Td_Distribucion_detalle2."sContrato"'
        Stretched = True
        object Memo86: TfrxMemoView
          Left = 257.008040000000000000
          Top = 18.897650000000000000
          Width = 306.141930000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Memo.UTF8W = (
            'PERIODOS DE EJECUCION DE LA ACTIVIDAD')
        end
        object Memo93: TfrxMemoView
          Top = 56.692950000000000000
          Width = 111.118110236220000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'FECHA')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo94: TfrxMemoView
          Left = 111.488250000000000000
          Top = 56.692950000000000000
          Width = 75.590551180000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'INICIO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo95: TfrxMemoView
          Left = 186.960730000000000000
          Top = 56.692950000000000000
          Width = 75.590551180000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'TERMINO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo96: TfrxMemoView
          Left = 263.433210000000000000
          Top = 56.692950000000000000
          Width = 109.606370000000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'AFECTACI'#211'N')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo97: TfrxMemoView
          Left = 372.803340000000000000
          Top = 56.692950000000000000
          Width = 83.149660000000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'INTERVALO'
            'TIEMPO')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo98: TfrxMemoView
          Left = 456.393940000000000000
          Top = 56.692950000000000000
          Width = 94.488188980000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'AVANCE'
            'ANTERIOR')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo99: TfrxMemoView
          Left = 550.543600000000000000
          Top = 56.692950000000000000
          Width = 94.488188980000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'AVANCE'
            'ACTUAL')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo100: TfrxMemoView
          Left = 644.693260000000000000
          Top = 56.692950000000000000
          Width = 94.488250000000000000
          Height = 30.236240000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'AVANCE'
            'ACUMULADO')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupFooter5: TfrxGroupFooter
        Height = 110.283550000000000000
        Top = 566.929500000000000000
        Width = 740.409927000000000000
        object Memo103: TfrxMemoView
          Left = 154.960730000000000000
          Top = 45.354360000000000000
          Width = 583.558955040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ACTIVIDAD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo104: TfrxMemoView
          Top = 64.929190000000000000
          Width = 154.960730000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_AvancesPartidas
          DataSetName = 'Td_AvancesPartidas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_PartidasAnexo."sNumeroActividad"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo105: TfrxMemoView
          Left = 154.960730000000000000
          Top = 64.929190000000000000
          Width = 583.559028270000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_AvancesPartidas
          DataSetName = 'Td_AvancesPartidas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Memo.UTF8W = (
            '[Td_PartidasAnexo."mDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo124: TfrxMemoView
          Top = 45.354360000000000000
          Width = 154.204751260000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PARTIDA')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo125: TfrxMemoView
          Width = 373.417322830000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'DURACION TIEMPO EFECTIVO (HRS):')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo126: TfrxMemoView
          Left = 372.661417320000000000
          Width = 83.149606300000000000
          Height = 18.897650000000000000
          OnBeforePrint = 'Memo126OnBeforePrint'
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = 'hh:mm'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            
              '[(SUM((iif(<Td_Distribucion_detalle2."sIdClasificacion">='#39'NOTA'#39',' +
              '0,iif(<Td_Distribucion_detalle2."sHoraFinal">='#39'24:00'#39',1, StrToTi' +
              'me(<Td_Distribucion_detalle2."sHoraFinal">) ) -    iif(<Td_Distr' +
              'ibucion_detalle2."sHoraInicio">='#39'00:00'#39',0,StrToTime(<Td_Distribu' +
              'cion_detalle2."sHoraInicio">) ))),DetailData2))]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupHeader6: TfrxGroupHeader
        Height = 37.795300000000000000
        Top = 699.213050000000000000
        Width = 740.409927000000000000
        Condition = 'Td_resumenMaterial."sTipo"'
        object Memo106: TfrxMemoView
          Top = 1.000000000000000000
          Width = 738.519685040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_resumenMaterial."sTipo"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo108: TfrxMemoView
          Top = 18.897650000000000000
          Width = 86.929190000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PARTIDA')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo109: TfrxMemoView
          Left = 87.488250000000000000
          Top = 18.897650000000000000
          Width = 265.322832200000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'DESCRIPCI'#211'N')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo110: TfrxMemoView
          Left = 476.645950000000000000
          Top = 18.897650000000000000
          Width = 60.472440940000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PU MN')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo111: TfrxMemoView
          Left = 538.236550000000000000
          Top = 18.897650000000000000
          Width = 60.472440940000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PU USD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo112: TfrxMemoView
          Left = 598.827150000000000000
          Top = 18.897650000000000000
          Width = 71.433070866141700000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'IMP MN')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo113: TfrxMemoView
          Left = 671.047244090000000000
          Top = 18.897650000000000000
          Width = 67.275590551181090000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'IMP USD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo114: TfrxMemoView
          Left = 354.141930000000000000
          Top = 18.897650000000000000
          Width = 61.984247090000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'UNIDAD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo115: TfrxMemoView
          Left = 416.503937010000000000
          Top = 18.897650000000000000
          Width = 59.338582680000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'CANTIDAD')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupFooter6: TfrxGroupFooter
        Height = 83.149660000000000000
        Top = 846.614720000000000000
        Width = 740.409927000000000000
        object Memo130: TfrxMemoView
          Left = 354.141732283465000000
          Top = 41.574830000000000000
          Width = 244.157480314961000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'COSTO TOTAL DE LA ACTIVIDAD:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo131: TfrxMemoView
          Left = 599.055118110000000000
          Top = 41.574830000000000000
          Width = 71.433070866141700000
          Height = 18.897650000000000000
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUM(<Td_resumenMaterial."dImporteMn">,DetailData3)]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo132: TfrxMemoView
          Left = 670.866141730000000000
          Top = 41.574830000000000000
          Width = 67.275590551181090000
          Height = 18.897650000000000000
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUM(<Td_resumenMaterial."dImporteDll">,DetailData3)]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo221: TfrxMemoView
          Top = 7.559060000000000000
          Width = 314.456722200000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Memo.UTF8W = (
            'NOTA: PARA ESTA ACTIVIDAD NO APLICAN MATERIALES')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object DetailData3: TfrxDetailData
        Height = 18.897650000000000000
        Top = 759.685530000000000000
        Width = 740.409927000000000000
        DataSet = frmDiarioTurno2.Td_resumenMaterial
        DataSetName = 'Td_resumenMaterial'
        RowCount = 0
        Stretched = True
        object Memo107: TfrxMemoView
          Width = 86.929190000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_resumenMaterial."sIdRecurso"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo116: TfrxMemoView
          Left = 87.488250000000000000
          Width = 265.322832200000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Memo.UTF8W = (
            '[Td_resumenMaterial."sdescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo117: TfrxMemoView
          Left = 354.141930000000000000
          Width = 61.984247090000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_resumenMaterial."sMedida"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo118: TfrxMemoView
          Left = 416.503937007874000000
          Width = 59.338582677165400000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2f'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_resumenMaterial."dCantidad"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo119: TfrxMemoView
          Left = 476.535433070000000000
          Width = 60.472440940000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_resumenMaterial."dVentaMn"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo120: TfrxMemoView
          Left = 538.251968500000000000
          Width = 60.472440940000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_resumenMaterial."dVentaDll"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo121: TfrxMemoView
          Left = 598.716535430000000000
          Width = 71.433070870000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_resumenMaterial."dImporteMn"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo122: TfrxMemoView
          Left = 671.047244090000000000
          Width = 67.275590550000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_resumenMaterial."dImporteDLL"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
    end
    object Page4: TfrxReportPage
      PaperWidth = 215.900000000000000000
      PaperHeight = 279.400000000000000000
      PaperSize = 1
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object PageHeader4: TfrxPageHeader
        Height = 245.669476850000000000
        Top = 18.897650000000000000
        Width = 740.409927000000000000
        OnAfterPrint = 'PageHeader1OnAfterPrint'
        object Picture7: TfrxPictureView
          Left = 591.606680000000000000
          Top = 4.779530000000000000
          Width = 143.622091180000000000
          Height = 45.354333150000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmReportePeriodo.dsConfiguracion
          DataSetName = 'dsConfiguracion'
          Frame.Color = clNone
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Picture8: TfrxPictureView
          Left = 3.779530000000000000
          Width = 136.063006770000000000
          Height = 49.133863150000000000
          ShowHint = False
          DataField = 'bImagen'
          DataSet = frmDiarioTurno2.Td_contrato
          DataSetName = 'Td_contrato'
          HightQuality = False
          Transparent = False
          TransparentColor = clWhite
        end
        object Memo133: TfrxMemoView
          Left = 140.842610000000000000
          Width = 449.764070000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sNombre"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo134: TfrxMemoView
          Left = 139.842610000000000000
          Top = 27.590551181102400000
          Width = 449.764070000000000000
          Height = 22.677180000000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          HAlign = haCenter
          Memo.UTF8W = (
            'DESGLOSE DE COSTOS')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo135: TfrxMemoView
          Left = 337.133858270000000000
          Top = 66.031540000000000000
          Width = 192.755905510000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'EMBARCACION:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo136: TfrxMemoView
          Left = 528.755905510000000000
          Top = 66.031540000000000000
          Width = 210.519685040000000000
          Height = 13.228346460000000000
          OnBeforePrint = 'Memo7OnBeforePrint'
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsEmbarcacion."sDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo137: TfrxMemoView
          Top = 66.252044170000000000
          Width = 142.110236220000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'CONTRATO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo138: TfrxMemoView
          Left = 142.291745530000000000
          Top = 66.252044170000000000
          Width = 195.779527560000000000
          Height = 13.228346460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[dsConfiguracion."sContratoBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo139: TfrxMemoView
          Top = 79.370078740000000000
          Width = 142.110236220000000000
          Height = 39.685039370000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'DESCRIPCI'#211'N:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo140: TfrxMemoView
          Left = 141.732283460000000000
          Top = 79.370078740000000000
          Width = 597.543307090000000000
          Height = 39.685039370000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Frame.Width = 0.500000000000000000
          Memo.UTF8W = (
            '[dsConfiguracion."mDescripcionBarco"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo141: TfrxMemoView
          Top = 118.677148270000000000
          Width = 142.110236220000000000
          Height = 39.685056460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'ORDEN DE TRABAJO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo142: TfrxMemoView
          Top = 207.937034720000000000
          Width = 142.110236220000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'PERIODO DE EJECUCI'#211'N:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo143: TfrxMemoView
          Left = 142.110219130000000000
          Top = 208.047244090000000000
          Width = 195.779478740000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."dFiProgramado"] AL [Td_partidas."dFfProgramado"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo144: TfrxMemoView
          Left = 141.732107720000000000
          Top = 119.165366540000000000
          Width = 597.543485280000000000
          Height = 39.307086610000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haBlock
          Memo.UTF8W = (
            
              '[td_contrato."sLabelContrato"]. -[dsConfiguracion."mDescripcion"' +
              ']')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo145: TfrxMemoView
          Top = 158.740260000000000000
          Width = 142.110236220000000000
          Height = 20.787406460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            'PART./FOLIO:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo146: TfrxMemoView
          Left = 141.732107720000000000
          Top = 159.228478270000000000
          Width = 597.543307090000000000
          Height = 48.377952760000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haBlock
          Memo.UTF8W = (
            '[Td_partidas."mDescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo147: TfrxMemoView
          Left = 337.157700000000000000
          Top = 208.047244090000000000
          Width = 192.756030000000000000
          Height = 28.346456690000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'INSTALACION: [Td_partidas."Plataforma"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo148: TfrxMemoView
          Left = 528.756186220000000000
          Top = 208.047244090000000000
          Width = 210.519685040000000000
          Height = 28.346456690000000000
          ShowHint = False
          StretchMode = smMaxHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'ACTIVO: [Td_partidas."Ubicacion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo149: TfrxMemoView
          Top = 179.417440000000000000
          Width = 142.110236220000000000
          Height = 28.346466460000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          Frame.Width = 0.500000000000000000
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_partidas."sIdFolio"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object PageFooter4: TfrxPageFooter
        Height = 22.677180000000000000
        Top = 623.622450000000000000
        Width = 740.409927000000000000
      end
      object GroupFooter7: TfrxGroupFooter
        Height = 22.677180000000000000
        Top = 476.220780000000000000
        Width = 740.409927000000000000
        PrintChildIfInvisible = True
        object Memo150: TfrxMemoView
          Left = 402.897637800000000000
          Width = 192.000000000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'IMPORTE [Td_ResumenPersonal."sTipo"]:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo151: TfrxMemoView
          Left = 594.897637800000000000
          Width = 71.811023622047210000
          Height = 18.897650000000000000
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUM(<Td_ResumenPersonal."dImporteMn">,MasterData4)]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo152: TfrxMemoView
          Left = 667.086614170000000000
          Width = 71.811023622047210000
          Height = 18.897650000000000000
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUM(<Td_ResumenPersonal."dImporteDll">,Masterdata4)]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupHeader7: TfrxGroupHeader
        Height = 37.795300000000000000
        Top = 370.393940000000000000
        Width = 740.409927000000000000
        Condition = 'Td_ResumenPersonal."sTipo"'
        object Memo159: TfrxMemoView
          Top = 1.000000000000000000
          Width = 738.519685040000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_ResumenPersonal."sTipo"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo160: TfrxMemoView
          Top = 18.897650000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PARTIDA')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo161: TfrxMemoView
          Left = 94.488250000000000000
          Top = 18.897650000000000000
          Width = 242.645652200000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'DESCRIPCI'#211'N')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo162: TfrxMemoView
          Left = 463.645950000000000000
          Top = 18.897650000000000000
          Width = 64.251970940000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PU MN')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo163: TfrxMemoView
          Left = 527.236550000000000000
          Top = 18.897650000000000000
          Width = 68.031500940000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'PU USD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo164: TfrxMemoView
          Left = 594.827150000000000000
          Top = 18.897650000000000000
          Width = 71.811030940000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'IMP MN')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo165: TfrxMemoView
          Left = 667.047244090000000000
          Top = 18.897650000000000000
          Width = 71.811016300000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'IMP USD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo166: TfrxMemoView
          Left = 337.141930000000000000
          Top = 18.897650000000000000
          Width = 65.763777090000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'UNIDAD')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo167: TfrxMemoView
          Left = 403.173470000000000000
          Top = 18.897650000000000000
          Width = 60.472480000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'CANTIDAD')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupFooter9: TfrxGroupFooter
        Height = 41.574830000000000000
        Top = 521.575140000000000000
        Width = 740.409927000000000000
        object Memo168: TfrxMemoView
          Left = 337.133858270000000000
          Top = 11.338590000000000000
          Width = 257.008040000000000000
          Height = 18.897650000000000000
          ShowHint = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            'COSTO TOTAL DE LA ACTIVIDAD:')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo169: TfrxMemoView
          Left = 594.897637800000000000
          Top = 11.338590000000000000
          Width = 71.811023622047210000
          Height = 18.897650000000000000
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUM(<Td_ResumenPersonal."dImporteMn">,MasterData4)]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo170: TfrxMemoView
          Left = 667.086614170000000000
          Top = 11.338590000000000000
          Width = 71.811023622047210000
          Height = 18.897650000000000000
          ShowHint = False
          DataSet = frmDiarioTurno2.Td_Distribucion_detalle2
          DataSetName = 'Td_Distribucion_detalle2'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[SUM(<Td_ResumenPersonal."dImporteDll">,Masterdata4)]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
      object GroupHeader8: TfrxGroupHeader
        Height = 22.677180000000000000
        Top = 325.039580000000000000
        Width = 740.409927000000000000
        Condition = 'Td_ResumenPersonal."sFolio"'
      end
      object MasterData4: TfrxMasterData
        Height = 22.677180000000000000
        Top = 430.866420000000000000
        Width = 740.409927000000000000
        DataSet = frmDiarioTurno2.Td_ResumenPersonal
        DataSetName = 'Td_ResumenPersonal'
        RowCount = 0
        Stretched = True
        object Memo171: TfrxMemoView
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftLeft, ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_ResumenPersonal."sIdRecurso"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo172: TfrxMemoView
          Left = 94.488250000000000000
          Width = 242.645652200000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          Memo.UTF8W = (
            '[Td_ResumenPersonal."sdescripcion"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo173: TfrxMemoView
          Left = 337.141930000000000000
          Width = 65.763777090000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_ResumenPersonal."sMedida"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo174: TfrxMemoView
          Left = 403.062992130000000000
          Width = 60.472440940000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2f'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haCenter
          Memo.UTF8W = (
            '[Td_ResumenPersonal."dCantidad"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo175: TfrxMemoView
          Left = 463.535433070000000000
          Width = 64.251970940000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_ResumenPersonal."dVentaMn"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo176: TfrxMemoView
          Left = 527.125984250000000000
          Width = 68.031500940000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_ResumenPersonal."dVentaDll"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo177: TfrxMemoView
          Left = 594.716535430000000000
          Width = 71.811030940000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_ResumenPersonal."dImporteMn"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo178: TfrxMemoView
          Left = 667.047244090000000000
          Width = 71.811021180000000000
          Height = 18.897650000000000000
          ShowHint = False
          StretchMode = smMaxHeight
          DataSet = frmDiarioTurno2.Td_resumenMaterial
          DataSetName = 'Td_resumenMaterial'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = [ftRight, ftTop, ftBottom]
          HAlign = haRight
          Memo.UTF8W = (
            '[Td_ResumenPersonal."dImporteDLL"]')
          ParentFont = False
          VAlign = vaCenter
        end
      end
    end
  end
  object QrImprimir: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select ae.sContrato, ae.sContrato as sOrden,ae.iIdActa,ae.sNoAct' +
        'a,ae.sNumeroOrden,ae.dFecha,ae.dFecha as dIdfecha,ae.etipo,ae.sE' +
        'specialidad,'
      
        '"" as sIdTurno,ae.mObservaciones,ae.sFirma1,ae.sFirma2,ae.lperno' +
        'cta,ae.lmaterial,ae.lpaginado,ae.lPartidas,ae.sActivo,ae.sCentro' +
        'Proceso,c.eLugarOT'
      'from acta_entrega ae '
      'inner join contratos c'
      'on(c.sContrato=ae.sContrato)'
      'where iIdActa=:Acta')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Acta'
        ParamType = ptUnknown
      end>
    Left = 448
    Top = 248
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Acta'
        ParamType = ptUnknown
      end>
  end
  object pmDatos: TPopupMenu
    OnPopup = pmDatosPopup
    Left = 856
    Top = 72
    object mniAdd: TMenuItem
      Caption = 'Insertar'
      OnClick = mniAddClick
    end
    object mniDelete: TMenuItem
      Caption = 'Eliminar'
    end
  end
  object cxStyleRepository1: TcxStyleRepository
    PixelsPerInch = 96
    object cxStyle1: TcxStyle
      AssignedValues = [svColor]
      Color = 15138559
    end
  end
  object Md1: TdxMemData
    Indexes = <>
    SortOptions = []
    Left = 360
    Top = 240
  end
  object JMry1: TJvMemoryData
    FieldDefs = <>
    Left = 344
    Top = 240
  end
  object QMateriales: TZQuery
    Connection = connection.zConnection
    UpdateObject = UdSqlMateriales
    AfterInsert = QMaterialesAfterInsert
    BeforePost = QMaterialesBeforePost
    SQL.Strings = (
      'select am.*,i.mdescripcion from acta_material am'
      'inner join insumos i'
      'on(i.sContrato=am.sContrato and i.sIdInsumo=am.sIdInsumo)'
      'where am.sContrato=:Contrato and am.iIdActa=:Acta'
      'order by am.sNumeroActividad, i.mDescripcion')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Acta'
        ParamType = ptUnknown
      end>
    Left = 344
    Top = 296
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Acta'
        ParamType = ptUnknown
      end>
  end
  object dsMateriales: TDataSource
    DataSet = QMateriales
    Left = 432
    Top = 272
  end
  object QrActividades: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select * from actividadesxorden where sContrato=:Contrato and sN' +
        'umeroOrden=:Orden and sTipoActividad='#39'Actividad'#39
      'order by iItemOrden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Orden'
        ParamType = ptUnknown
      end>
    Left = 152
    Top = 344
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Orden'
        ParamType = ptUnknown
      end>
  end
  object QrInsumos: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select * from insumos where sContrato=:Contrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
    Left = 216
    Top = 344
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
  end
  object dsActividades: TDataSource
    DataSet = QrActividades
    Left = 152
    Top = 376
  end
  object dsInsumos: TDataSource
    DataSet = QrInsumos
    Left = 216
    Top = 376
  end
  object UdSqlMateriales: TZUpdateSQL
    DeleteSQL.Strings = (
      
        'delete from acta_material where sContrato=:sContrato and iIdActa' +
        '=:iIdActa and sNumeroActividad=:sNumeroActividad and sIdInsumo=:' +
        'sIdInsumo')
    InsertSQL.Strings = (
      
        'insert into acta_material values(:sContrato,:iIdActa,:sNumeroAct' +
        'ividad,:sIdInsumo,:sMedida,:dCantidad,:sTrazabilidad)')
    ModifySQL.Strings = (
      
        'update acta_material set sNumeroActividad=:sNumeroActividad, sId' +
        'Insumo=:sIdInsumo, sMedida=:sMedida, dCantidad=:dCantidad, sTraz' +
        'abilidad=:sTrazabilidad'
      
        'where sContrato=:sContrato and iIdActa=:iIdActa and sNumeroActiv' +
        'idad=:Old_sNumeroActividad and sIdInsumo=:Old_sIdInsumo')
    UseSequenceFieldForRefreshSQL = False
    Left = 504
    Top = 312
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'sNumeroActividad'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sIdInsumo'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sMedida'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'dCantidad'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sTrazabilidad'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'sContrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'iIdActa'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Old_sNumeroActividad'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Old_sIdInsumo'
        ParamType = ptUnknown
      end>
  end
  object dlgSaveGuardar: TSaveDialog
    DefaultExt = '.xls'
    Filter = 'Libro de Excel 97-2003|*.xls'
    OnTypeChange = dlgSaveGuardarTypeChange
    Left = 368
    Top = 376
  end
  object pmActa: TPopupMenu
    Left = 424
    Top = 152
    object mniRecalcular: TMenuItem
      Caption = 'Recalcular'
      OnClick = mniRecalcularClick
    end
  end
end
