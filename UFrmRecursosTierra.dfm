object FrmRecursosTierra: TFrmRecursosTierra
  Left = 0
  Top = 0
  Caption = 'FrmRecursosTierra'
  ClientHeight = 444
  ClientWidth = 756
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GBxClient: TcxGroupBox
    Left = 0
    Top = 0
    Align = alClient
    Caption = 'GBxClient'
    TabOrder = 0
    Height = 208
    Width = 756
    object cxGrid1: TcxGrid
      Left = 2
      Top = 18
      Width = 752
      Height = 188
      Align = alClient
      TabOrder = 0
      object cxGrid1DBTableView1: TcxGridDBTableView
        Navigator.Buttons.CustomButtons = <>
        DataController.DataSource = dsConceptos
        DataController.DetailKeyFieldNames = 'iIddiario'
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        FilterRow.Visible = True
        OptionsData.Deleting = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsView.GridLines = glHorizontal
        OptionsView.GroupByBox = False
        Styles.Header = cxStyle1
        object cxGrid1DBTableView1Column1: TcxGridDBColumn
          Caption = 'Id'
          DataBinding.FieldName = 'iIdDiario'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          HeaderAlignmentHorz = taCenter
          Styles.Header = cxStyle1
          Width = 50
        end
        object cxGrid1DBTableView1Column4: TcxGridDBColumn
          Caption = 'Horario'
          DataBinding.FieldName = 'Nota'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          HeaderAlignmentHorz = taCenter
          Width = 100
        end
        object cxGrid1DBTableView1Column2: TcxGridDBColumn
          Caption = 'Descripcion/Tipo'
          DataBinding.FieldName = 'mDescripcion'
          HeaderAlignmentHorz = taCenter
          Width = 500
        end
        object cxGrid1DBTableView1Column3: TcxGridDBColumn
          Caption = 'Avance'
          DataBinding.FieldName = 'dAvance'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Visible = False
          HeaderAlignmentHorz = taCenter
          Width = 80
        end
        object cxGrid1DBTableView1Column6: TcxGridDBColumn
          DataBinding.FieldName = 'iIdActividad'
          Visible = False
        end
        object cxGrid1DBTableView1Column7: TcxGridDBColumn
          DataBinding.FieldName = 'iHermano'
          Visible = False
        end
      end
      object CxGrdDbTblVGrid1DBTableView2: TcxGridDBTableView
        Navigator.Buttons.CustomButtons = <>
        OnSelectionChanged = CxGrdDbTblVGrid1DBTableView2SelectionChanged
        DataController.DataSource = dsNotasCortes
        DataController.DetailKeyFieldNames = 'iIddiarioNota'
        DataController.KeyFieldNames = 'iIddiario'
        DataController.MasterKeyFieldNames = 'iIddiario'
        DataController.Options = [dcoAssignGroupingValues, dcoAssignMasterDetailKeys, dcoSaveExpanding, dcoGroupsAlwaysExpanded]
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsData.Deleting = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsSelection.CellSelect = False
        OptionsSelection.MultiSelect = True
        OptionsView.FocusRect = False
        OptionsView.ExpandButtonsForEmptyDetails = False
        OptionsView.GridLines = glHorizontal
        OptionsView.GroupByBox = False
        OptionsView.Header = False
        OptionsView.IndicatorWidth = 5
        object CxGrdDbTblVGrid1DBTableView2Column1: TcxGridDBColumn
          DataBinding.FieldName = 'iidDiario'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taRightJustify
          Width = 40
        end
        object CxGrdDbTblVGrid1DBTableView2Column3: TcxGridDBColumn
          DataBinding.FieldName = 'Nota'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taRightJustify
          Width = 100
        end
        object CxGrdDbTblVGrid1DBTableView2Column2: TcxGridDBColumn
          DataBinding.FieldName = 'sIdClasificacion'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Width = 70
        end
        object CxGrdDbTblVGrid1DBTableView2Column4: TcxGridDBColumn
          DataBinding.FieldName = 'Movimiento'
          Width = 300
        end
        object CxGrdDbTblVGrid1DBTableView2Column5: TcxGridDBColumn
          DataBinding.FieldName = 'dAvance'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Visible = False
          Width = 70
        end
        object CxGrdDbTblVGrid1DBTableView2Column6: TcxGridDBColumn
          DataBinding.FieldName = 'iIdActividad'
          Visible = False
        end
      end
      object cxGrid1Level1: TcxGridLevel
        GridView = cxGrid1DBTableView1
        object CxGLvlGrid1Level2: TcxGridLevel
          GridView = CxGrdDbTblVGrid1DBTableView2
          Options.DetailFrameColor = clNone
          Options.DetailFrameWidth = 0
          Options.TabsForEmptyDetails = False
        end
      end
    end
  end
  object GBxBottom: TcxGroupBox
    Left = 0
    Top = 216
    Align = alBottom
    Caption = 'GBxBottom'
    TabOrder = 1
    Height = 228
    Width = 756
    object CxPageRecursos: TcxPageControl
      Left = 2
      Top = 18
      Width = 752
      Height = 208
      Align = alClient
      TabOrder = 0
      Properties.ActivePage = cTsPersonal
      Properties.CustomButtons.Buttons = <>
      ClientRectBottom = 204
      ClientRectLeft = 4
      ClientRectRight = 748
      ClientRectTop = 24
      object cTsPersonal: TcxTabSheet
        Caption = 'cTsPersonal'
        ImageIndex = 0
        object cxGrid3: TcxGrid
          Left = 0
          Top = 0
          Width = 744
          Height = 180
          Align = alClient
          TabOrder = 0
          object cxGrid3DBTableView1: TcxGridDBTableView
            Navigator.Buttons.CustomButtons = <>
            DataController.DataSource = dsPersonal
            DataController.Summary.DefaultGroupSummaryItems = <>
            DataController.Summary.FooterSummaryItems = <>
            DataController.Summary.SummaryGroups = <>
            FilterRow.Visible = True
            OptionsView.GroupByBox = False
            object cxGrid3DBTableView1Column1: TcxGridDBColumn
              Caption = 'C'#243'digo'
              DataBinding.FieldName = 'sIdPersonal'
              Width = 80
            end
            object cxGrid3DBTableView1Column2: TcxGridDBColumn
              Caption = 'Personal'
              DataBinding.FieldName = 'sDescripcion'
              Options.Editing = False
              Width = 300
            end
            object cxGrid3DBTableView1Column3: TcxGridDBColumn
              Caption = 'Cant. Reportar'
              DataBinding.FieldName = 'dCantidad'
              Width = 80
            end
            object cxGrid3DBTableView1Column4: TcxGridDBColumn
              Caption = 'Cant. Solicitada'
              DataBinding.FieldName = 'dSolicitado'
              Options.Editing = False
              Width = 80
            end
          end
          object cxGrid3Level1: TcxGridLevel
            GridView = cxGrid3DBTableView1
          end
        end
      end
      object cTsEquipo: TcxTabSheet
        Caption = 'cTsEquipo'
        ImageIndex = 1
        object cxGrid4: TcxGrid
          Left = 0
          Top = 0
          Width = 744
          Height = 180
          Align = alClient
          TabOrder = 0
          object cxGridDBTableView1: TcxGridDBTableView
            Navigator.Buttons.CustomButtons = <>
            DataController.DataSource = dsEquipos
            DataController.Summary.DefaultGroupSummaryItems = <>
            DataController.Summary.FooterSummaryItems = <>
            DataController.Summary.SummaryGroups = <>
            FilterRow.Visible = True
            OptionsView.GroupByBox = False
            object cxGridDBColumn1: TcxGridDBColumn
              Caption = 'sIdEquipo'
              DataBinding.FieldName = 'sIdPersonal'
            end
            object cxGridDBColumn2: TcxGridDBColumn
              DataBinding.FieldName = 'sDescripcion'
            end
            object cxGridDBColumn3: TcxGridDBColumn
              DataBinding.FieldName = 'dCantidad'
            end
            object cxGridDBColumn4: TcxGridDBColumn
              DataBinding.FieldName = 'dSolicitado'
            end
            object cxGridDBColumn5: TcxGridDBColumn
            end
            object cxGridDBColumn6: TcxGridDBColumn
            end
          end
          object cxGridLevel1: TcxGridLevel
            GridView = cxGridDBTableView1
          end
        end
      end
    end
  end
  object SplPrincipal: TcxSplitter
    Left = 0
    Top = 208
    Width = 756
    Height = 8
    AlignSplitter = salBottom
  end
  object FramedPanel1: TFramedPanel
    Left = 240
    Top = 280
    Width = 393
    Height = 121
    BorderOuter.Size = 1
    BorderMiddle.Size = 5
    BorderInner.Size = 1
    TabOrder = 3
    Visible = False
    object GBx1: TcxGroupBox
      Left = 13
      Top = 72
      Align = alBottom
      Caption = 'GBx1'
      TabOrder = 0
      Height = 36
      Width = 367
    end
    object cxGrid2: TcxGrid
      Left = 13
      Top = 13
      Width = 367
      Height = 59
      Align = alClient
      TabOrder = 1
      object cxGrid2DBTableView1: TcxGridDBTableView
        Navigator.Buttons.CustomButtons = <>
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsView.GroupByBox = False
      end
      object cxGrid2Level1: TcxGridLevel
        GridView = cxGrid2DBTableView1
      end
    end
  end
  object QrNotasCortes: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select ba.iIdDiario, ba.mDescripcion, ba.sHoraInicio, ba.sHoraFi' +
        'nal, ba.sConceptoGerencial, ba.sIdClasificacion, ba.dCantidad, f' +
        'ormat(ba.dAvance,2) as dAvance, ba.sWbs,'
      
        'concat(ba.sHoraInicio," - ",ba.sHoraFinal) as Nota, ba.lImprime,' +
        'ba.iHermano,ba.iIdTarea,ba.iIdActividad, ba.eTipoActividad,ba.sI' +
        'dTipoMovimiento,ba.iIddiarioNota,tm.sDescripcion as Movimiento  '
      'from bitacoradeactividades ba'
      
        'left join tiposdemovimiento tm on (tm.sContrato=:barco and ba.sI' +
        'dClasificacion=tm.sIdTipoMovimiento)'
      ' where ba.sContrato =:Contrato and ba.sIdConvenio =:Convenio '
      
        'and ba.dIdFecha =:Fecha and ba.snumeroorden = :folio and Find_In' +
        '_Set(ba.sIdTipoMovimiento,"ED,EN") order by ba.sHoraInicio, ba.s' +
        'HoraFinal')
    Params = <
      item
        DataType = ftUnknown
        Name = 'barco'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Convenio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'folio'
        ParamType = ptUnknown
      end>
    Left = 352
    Top = 96
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'barco'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Convenio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'folio'
        ParamType = ptUnknown
      end>
  end
  object dsNotasCortes: TDataSource
    DataSet = QrNotasCortes
    Left = 432
    Top = 96
  end
  object QrConceptos: TZReadOnlyQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select iIdDiario, mDescripcion, sHoraInicio, sHoraFinal, sConcep' +
        'toGerencial, sIdClasificacion, dCantidad, format(dAvance,2) as d' +
        'Avance, sWbs,'
      
        'concat(sHoraInicio," - ",sHoraFinal) as Nota, lImprime,iHermano,' +
        'iIdTarea,iIdActividad, eTipoActividad,sIdTipoMovimiento,iIddiari' +
        'oNota '
      
        'from bitacoradeactividades where sContrato =:Contrato and sIdCon' +
        'venio =:Convenio '
      
        'and dIdFecha =:Fecha and snumeroorden = :folio and Find_In_Set(s' +
        'IdTipoMovimiento,"E") order by sHoraInicio, sHoraFinal')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Convenio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'folio'
        ParamType = ptUnknown
      end>
    Left = 352
    Top = 136
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Convenio'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'folio'
        ParamType = ptUnknown
      end>
  end
  object dsConceptos: TDataSource
    DataSet = QrConceptos
    Left = 424
    Top = 136
  end
  object cxStyleRepository1: TcxStyleRepository
    PixelsPerInch = 96
    object cxStyle1: TcxStyle
      AssignedValues = [svFont]
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
    end
    object cxStyle2: TcxStyle
    end
  end
  object QPersonal: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select * from bitacoradepersonal where sContrato=:Contrato and d' +
        'IdFecha=:fecha and iIdDiario=:Diario and iIdActividad=:Actividad'
      'order by iItemOrden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Diario'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Actividad'
        ParamType = ptUnknown
      end>
    Left = 248
    Top = 216
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Diario'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Actividad'
        ParamType = ptUnknown
      end>
  end
  object dsPersonal: TDataSource
    DataSet = QPersonal
    Left = 304
    Top = 216
  end
  object QEquipos: TZQuery
    Connection = connection.zConnection
    SQL.Strings = (
      
        'select * from bitacoradeequipos where sContrato=:Contrato and dI' +
        'dFecha=:fecha and iIdDiario=:Diario and iIdActividad=:Actividad'
      'order by iItemOrden')
    Params = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Diario'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Actividad'
        ParamType = ptUnknown
      end>
    Left = 424
    Top = 216
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Contrato'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'fecha'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Diario'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Actividad'
        ParamType = ptUnknown
      end>
  end
  object dsEquipos: TDataSource
    DataSet = QEquipos
    Left = 376
    Top = 224
  end
end
