object FrmNotaCampoObservaciones: TFrmNotaCampoObservaciones
  Left = 0
  Top = 0
  Caption = 'FrmNotaCampoObservaciones'
  ClientHeight = 602
  ClientWidth = 651
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -15
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  Scaled = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 17
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 651
    Height = 115
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object Label2: TLabel
      Left = 3
      Top = 31
      Width = 133
      Height = 17
      Caption = 'Fecha de Impresion:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
    end
    object Label13: TLabel
      Left = 3
      Top = 56
      Width = 142
      Height = 17
      Caption = 'Periodo de Ejecucion:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
    end
    object Panel3: TPanel
      Left = 0
      Top = 0
      Width = 651
      Height = 21
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 0
      object Label1: TLabel
        Left = 0
        Top = 0
        Width = 97
        Height = 21
        Align = alLeft
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'Folio:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -17
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lblFolio: TLabel
        Left = 97
        Top = 0
        Width = 554
        Height = 21
        Align = alClient
        Alignment = taCenter
        Caption = 'Folio:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -17
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        ExplicitWidth = 44
        ExplicitHeight = 19
      end
    end
    object AdvDBDateTimePicker1: TAdvDBDateTimePicker
      Left = 147
      Top = 30
      Width = 148
      Height = 25
      Date = 41901.528449074070000000
      Time = 41901.528449074070000000
      Color = 15138559
      DoubleBuffered = True
      Kind = dkDate
      ParentDoubleBuffered = False
      TabOrder = 1
      OnEnter = AdvDBDateTimePicker1Enter
      OnExit = AdvDBDateTimePicker1Exit
      OnKeyPress = AdvDBDateTimePicker1KeyPress
      BorderStyle = bsSingle
      Ctl3D = True
      DateTime = 41901.528449074070000000
      Version = '1.1.0.1'
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clWindowText
      LabelFont.Height = -13
      LabelFont.Name = 'Tahoma'
      LabelFont.Style = []
      DataField = 'dFecha'
      DataSource = dsNotaCampo
    end
    object DBMemo1: TDBMemo
      Left = 147
      Top = 53
      Width = 499
      Height = 58
      Color = 15138559
      DataField = 'sPeriodo'
      DataSource = dsNotaCampo
      TabOrder = 2
      OnEnter = DBMemo1Enter
      OnExit = DBMemo1Exit
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 115
    Width = 651
    Height = 487
    Align = alClient
    BevelOuter = bvLowered
    TabOrder = 1
    object AdvPageControl1: TAdvPageControl
      Left = 1
      Top = 214
      Width = 649
      Height = 272
      ActivePage = AdvTabSheet1
      ActiveFont.Charset = DEFAULT_CHARSET
      ActiveFont.Color = clWindowText
      ActiveFont.Height = -13
      ActiveFont.Name = 'Tahoma'
      ActiveFont.Style = []
      Align = alBottom
      TabBackGroundColor = clBtnFace
      TabMargin.RightMargin = 0
      TabOverlap = 0
      TabStyle = tsDelphi
      Version = '1.7.1.0'
      TabOrder = 1
      object AdvTabSheet1: TAdvTabSheet
        Caption = 'Observaciones'
        Color = clBtnFace
        ColorTo = clNone
        TabColor = clBtnFace
        TabColorTo = clNone
        object Label12: TLabel
          Left = 24
          Top = 40
          Width = 87
          Height = 17
          Caption = 'Observacion:'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
        end
        object AvDbDtpFecha: TAdvDBDateTimePicker
          Left = 80
          Top = 9
          Width = 186
          Height = 25
          Date = 41901.539224537040000000
          Time = 41901.539224537040000000
          Color = 15138559
          DoubleBuffered = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          Kind = dkDate
          ParentDoubleBuffered = False
          ParentFont = False
          TabOrder = 0
          OnEnter = AvDbDtpFechaEnter
          OnExit = AvDbDtpFechaExit
          OnKeyPress = AvDbDtpFechaKeyPress
          BorderStyle = bsSingle
          Ctl3D = True
          DateTime = 41901.539224537040000000
          Version = '1.1.0.1'
          LabelCaption = 'Fecha:  '
          LabelPosition = lpLeftCenter
          LabelTransparent = True
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -13
          LabelFont.Name = 'Arial'
          LabelFont.Style = []
          DataField = 'dFecha'
          DataSource = dsNotaObservaciones
        end
        object dbmmoObs: TDBMemo
          Left = 24
          Top = 56
          Width = 614
          Height = 190
          Color = 15138559
          DataField = 'sObservacion'
          DataSource = dsNotaObservaciones
          ScrollBars = ssVertical
          TabOrder = 1
          OnEnter = dbmmoObsEnter
          OnExit = dbmmoObsExit
        end
      end
      object AdvTabSheet2: TAdvTabSheet
        Caption = 'Firmantes'
        Color = clBtnFace
        ColorTo = clNone
        TabColor = clBtnFace
        TabColorTo = clNone
        object Label3: TLabel
          Left = 174
          Top = 6
          Width = 100
          Height = 17
          Caption = 'No. Firmantes: '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
        end
        object Label4: TLabel
          Left = 66
          Top = 67
          Width = 56
          Height = 17
          Caption = 'Nombre '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
        end
        object Label45: TLabel
          Left = 66
          Top = 39
          Width = 46
          Height = 17
          Caption = 'Puesto'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
        end
        object Label5: TLabel
          Left = 19
          Top = 17
          Width = 140
          Height = 16
          Caption = 'Firmante 1 (Izquierda)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object Label6: TLabel
          Left = 66
          Top = 135
          Width = 56
          Height = 17
          Caption = 'Nombre '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
        end
        object Label7: TLabel
          Left = 66
          Top = 112
          Width = 46
          Height = 17
          Caption = 'Puesto'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
        end
        object Label8: TLabel
          Left = 19
          Top = 90
          Width = 133
          Height = 16
          Caption = 'Firmante 2 (Derecha)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object Label9: TLabel
          Left = 66
          Top = 208
          Width = 56
          Height = 17
          Caption = 'Nombre '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
        end
        object Label10: TLabel
          Left = 66
          Top = 184
          Width = 46
          Height = 17
          Caption = 'Puesto'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
        end
        object Label11: TLabel
          Left = 19
          Top = 163
          Width = 122
          Height = 16
          Caption = 'Firmante 3 (Centro)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object tsFirma1: TDBAdvEditBtn
          Left = 129
          Top = 63
          Width = 452
          Height = 25
          Flat = False
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          Enabled = True
          TabOrder = 2
          Visible = True
          OnEnter = tsFirma1Enter
          OnExit = tsFirma1Exit
          OnKeyPress = tsFirma1KeyPress
          Version = '1.3.2.8'
          ButtonStyle = bsButton
          ButtonWidth = 16
          Etched = False
          ButtonCaption = '...'
          DataField = 'sFirmante1'
          DataSource = dsNotaCampo
        end
        object tsPuesto1: TDBEdit
          Left = 129
          Top = 37
          Width = 452
          Height = 25
          Cursor = crDrag
          Hint = 'Superintendente de Obra'
          CharCase = ecUpperCase
          Color = 14342874
          DataField = 'sPuesto1'
          DataSource = dsNotaCampo
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnEnter = tsPuesto1Enter
          OnExit = tsPuesto1Exit
          OnKeyPress = tsPuesto1KeyPress
        end
        object tsFirma2: TDBAdvEditBtn
          Left = 129
          Top = 132
          Width = 452
          Height = 25
          Flat = False
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          Enabled = True
          TabOrder = 4
          Visible = True
          OnEnter = tsFirma1Enter
          OnExit = tsFirma1Exit
          OnKeyPress = tsFirma2KeyPress
          Version = '1.3.2.8'
          ButtonStyle = bsButton
          ButtonWidth = 16
          Etched = False
          ButtonCaption = '...'
          DataField = 'sFirmante2'
          DataSource = dsNotaCampo
        end
        object tsPuesto2: TDBEdit
          Left = 129
          Top = 108
          Width = 452
          Height = 25
          Cursor = crDrag
          Hint = 'Superintendente de Obra'
          CharCase = ecUpperCase
          Color = 14342874
          DataField = 'sPuesto2'
          DataSource = dsNotaCampo
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnEnter = tsPuesto1Enter
          OnExit = tsPuesto1Exit
          OnKeyPress = tsPuesto2KeyPress
        end
        object tsFirma3: TDBAdvEditBtn
          Left = 129
          Top = 205
          Width = 452
          Height = 25
          Flat = False
          LabelFont.Charset = DEFAULT_CHARSET
          LabelFont.Color = clWindowText
          LabelFont.Height = -11
          LabelFont.Name = 'Tahoma'
          LabelFont.Style = []
          Lookup.Separator = ';'
          Color = 15138559
          Enabled = True
          TabOrder = 6
          Visible = True
          OnEnter = tsFirma1Enter
          OnExit = tsFirma1Exit
          Version = '1.3.2.8'
          ButtonStyle = bsButton
          ButtonWidth = 16
          Etched = False
          ButtonCaption = '...'
          DataField = 'sFirmante3'
          DataSource = dsNotaCampo
        end
        object tsPuesto3: TDBEdit
          Left = 129
          Top = 181
          Width = 452
          Height = 25
          Cursor = crDrag
          Hint = 'Superintendente de Obra'
          CharCase = ecUpperCase
          Color = 14342874
          DataField = 'sPuesto3'
          DataSource = dsNotaCampo
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
          OnEnter = tsPuesto1Enter
          OnExit = tsPuesto1Exit
          OnKeyPress = tsPuesto3KeyPress
        end
        object JDbCmbFirmantes: TJvDBComboBox
          Left = 278
          Top = 3
          Width = 78
          Height = 25
          Color = 15138559
          DataField = 'iNumFirmante'
          DataSource = dsNotaCampo
          Items.Strings = (
            '2'
            '3')
          TabOrder = 0
          Values.Strings = (
            '2'
            '3')
          ListSettings.OutfilteredValueFont.Charset = DEFAULT_CHARSET
          ListSettings.OutfilteredValueFont.Color = clRed
          ListSettings.OutfilteredValueFont.Height = -13
          ListSettings.OutfilteredValueFont.Name = 'Tahoma'
          ListSettings.OutfilteredValueFont.Style = []
          OnEnter = JDbCmbFirmantesEnter
          OnExit = JDbCmbFirmantesExit
          OnKeyPress = JDbCmbFirmantesKeyPress
        end
      end
    end
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 649
      Height = 213
      Align = alClient
      Caption = 'Panel4'
      TabOrder = 0
      inline frmBarra1: TfrmBarra
        Left = 1
        Top = 1
        Width = 73
        Height = 211
        VertScrollBar.Style = ssHotTrack
        Align = alLeft
        TabOrder = 0
        ExplicitLeft = 1
        ExplicitTop = 1
        ExplicitWidth = 73
        ExplicitHeight = 211
        inherited AdvPanel1: TAdvPanel
          Width = 73
          Height = 211
          ExplicitWidth = 73
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
      end
      object JvDBUltimGrid1: TJvDBUltimGrid
        Left = 74
        Top = 1
        Width = 574
        Height = 211
        Align = alClient
        Color = 15138559
        DataSource = dsNotaObservaciones
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -15
        TitleFont.Name = 'Arial'
        TitleFont.Style = [fsBold]
        SelectColumnsDialogStrings.Caption = 'Select columns'
        SelectColumnsDialogStrings.OK = '&OK'
        SelectColumnsDialogStrings.NoSelectionWarning = 'At least one column must be visible!'
        EditControls = <>
        RowsHeight = 21
        TitleRowHeight = 22
        Columns = <
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'dFecha'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Arial'
            Font.Style = []
            Title.Alignment = taCenter
            Title.Caption = 'Fecha'
            Width = 80
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'sObservacion'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Arial'
            Font.Style = []
            Title.Alignment = taCenter
            Title.Caption = 'Observacion'
            Width = 450
            Visible = True
          end>
      end
    end
  end
  object QNotaCampo: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select * from notacampo_general where sContrato=:Contrato and sN' +
        'umeroOrden=:Folio')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Folio'
        ParamType = ptUnknown
      end>
    Left = 320
    Top = 16
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Folio'
        ParamType = ptUnknown
      end>
  end
  object dsNotaCampo: TDataSource
    DataSet = QNotaCampo
    Left = 368
    Top = 16
  end
  object QNotaObservaciones: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select * from notacampo_observaciones where iIdNota=:Nota')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Nota'
        ParamType = ptUnknown
      end>
    Left = 320
    Top = 288
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Nota'
        ParamType = ptUnknown
      end>
  end
  object dsNotaObservaciones: TDataSource
    AutoEdit = False
    DataSet = QNotaObservaciones
    Left = 360
    Top = 288
  end
end
