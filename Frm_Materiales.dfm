object FrmAltaMAterial: TFrmAltaMAterial
  Left = 0
  Top = 0
  BorderStyle = bsSizeToolWin
  Caption = 'Registro de Material'
  ClientHeight = 379
  ClientWidth = 655
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GBx1: TcxGroupBox
    Left = 0
    Top = 318
    Align = alBottom
    TabOrder = 0
    ExplicitTop = 294
    ExplicitWidth = 603
    Height = 61
    Width = 655
  end
  object GBx2: TcxGroupBox
    Left = 0
    Top = 0
    Align = alClient
    PanelStyle.Active = True
    TabOrder = 1
    ExplicitTop = -6
    ExplicitWidth = 603
    ExplicitHeight = 288
    Height = 318
    Width = 655
    object CxGrd1: TcxGrid
      Left = 2
      Top = 2
      Width = 651
      Height = 314
      Align = alClient
      TabOrder = 0
      ExplicitLeft = 200
      ExplicitTop = 56
      ExplicitWidth = 250
      ExplicitHeight = 200
      object CxGrdDbTblVMateriales: TcxGridDBTableView
        Navigator.Buttons.CustomButtons = <>
        DataController.DataModeController.GridMode = True
        DataController.DataSource = dsInsumos
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        FilterRow.Visible = True
        OptionsData.Appending = True
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsView.CellAutoHeight = True
        OptionsView.ColumnAutoWidth = True
        OptionsView.GroupByBox = False
        object CxGrdDbTblVMaterialesColumn1: TcxGridDBColumn
          Caption = 'Id Almacen'
          DataBinding.FieldName = 'sIdAlmacen'
          PropertiesClassName = 'TcxLookupComboBoxProperties'
          Properties.KeyFieldNames = 'sIdAlmacen'
          Properties.ListColumns = <
            item
              Caption = 'Id Almacen'
              Width = 70
              FieldName = 'sIdAlmacen'
            end
            item
              Caption = 'Almacen'
              Width = 200
              FieldName = 'sDescripcion'
            end>
          Properties.ListSource = dsAlmacenes
          HeaderAlignmentHorz = taCenter
          Width = 70
        end
        object CxGrdDbTblVMaterialesColumn2: TcxGridDBColumn
          Caption = 'Id Material'
          DataBinding.FieldName = 'sIdInsumo'
          HeaderAlignmentHorz = taCenter
          Width = 70
        end
        object CxGrdDbTblVMaterialesColumn3: TcxGridDBColumn
          Caption = 'Material'
          DataBinding.FieldName = 'mDescripcion'
          PropertiesClassName = 'TcxMemoProperties'
          Properties.VisibleLineCount = 2
          HeaderAlignmentHorz = taCenter
          Width = 250
        end
        object CxGrdDbTblVMaterialesColumn4: TcxGridDBColumn
          Caption = 'U. Medida'
          DataBinding.FieldName = 'sMedida'
          HeaderAlignmentHorz = taCenter
          Width = 80
        end
        object CxGrdDbTblVMaterialesColumn5: TcxGridDBColumn
          Caption = 'Trazabilidad'
          DataBinding.FieldName = 'sTrazabilidad'
          HeaderAlignmentHorz = taCenter
          Width = 150
        end
      end
      object CxGLvlGrid1Level1: TcxGridLevel
        GridView = CxGrdDbTblVMateriales
      end
    end
  end
  object QInsumos: TZQuery
    Connection = connection.zConnection
    AfterInsert = QInsumosAfterInsert
    BeforePost = QInsumosBeforePost
    SQL.Strings = (
      
        'select sContrato,sIdInsumo,sIdAlmacen,mDescripcion,sTrazabilidad' +
        ',sMedida,sColumnaAux from insumos where sContrato=:Contrato')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
    Left = 176
    Top = 176
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end>
  end
  object dsInsumos: TDataSource
    DataSet = QInsumos
    Left = 232
    Top = 184
  end
  object QrAlmacen: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      'select * from almacenes')
    Params = <>
    Left = 320
    Top = 192
  end
  object dsAlmacenes: TDataSource
    DataSet = QrAlmacen
    Left = 376
    Top = 192
  end
end
