unit frm_CalculoAvancesxPartida;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, utilerias, frm_connection, DBCtrls, StdCtrls, Buttons, DB, Global,
  ZDataset, ZAbstractRODataset, Grids,
  DBGrids, RXDBCtrl, Gauges, unitexcepciones, udbgrid, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxStyles, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom,
  dxSkinDarkSide, dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
  dxSkinGlassOceans, dxSkinHighContrast, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMetropolis,
  dxSkinMetropolisDark, dxSkinMoneyTwins, dxSkinOffice2007Black,
  dxSkinOffice2007Blue, dxSkinOffice2007Green, dxSkinOffice2007Pink,
  dxSkinOffice2007Silver, dxSkinOffice2010Black, dxSkinOffice2010Blue,
  dxSkinOffice2010Silver, dxSkinOffice2013DarkGray, dxSkinOffice2013LightGray,
  dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic,
  dxSkinSharp, dxSkinSharpPlus, dxSkinSilver, dxSkinSpringTime, dxSkinStardust,
  dxSkinSummer2008, dxSkinTheAsphaltWorld, dxSkinsDefaultPainters,
  dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint, dxSkinXmas2008Blue,
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit,
  cxNavigator, cxDBData, cxTextEdit, cxMaskEdit, cxDBLookupComboBox,
  cxDropDownEdit, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxGridCustomView, cxGrid, ExtCtrls;

type
  TfrmCalculoAvancesxPartida = class(TForm)
    ds_ordenesdetrabajo: TDataSource;
    ds_actividadesxorden: TDataSource;
    ordenesdetrabajo: TZReadOnlyQuery;
    ActividadesxOrden: TZReadOnlyQuery;
    ActividadesxOrdensContrato: TStringField;
    ActividadesxOrdensIdConvenio: TStringField;
    ActividadesxOrdensNumeroOrden: TStringField;
    ActividadesxOrdeniNivel: TIntegerField;
    ActividadesxOrdensSimbolo: TStringField;
    ActividadesxOrdensWbs: TStringField;
    ActividadesxOrdensWbsAnterior: TStringField;
    ActividadesxOrdensPaquete: TStringField;
    ActividadesxOrdensNumeroActividad: TStringField;
    ActividadesxOrdensTipoActividad: TStringField;
    ActividadesxOrdeniItemOrden: TStringField;
    ActividadesxOrdenmDescripcion: TMemoField;
    ActividadesxOrdendPonderado: TFloatField;
    ActividadesxOrdensMedida: TStringField;
    ActividadesxOrdendCantidad: TFloatField;
    ActividadesxOrdendCargado: TFloatField;
    ActividadesxOrdendInstalado: TFloatField;
    ActividadesxOrdendExcedente: TFloatField;
    ActividadesxOrdendVentaMN: TFloatField;
    ActividadesxOrdendVentaDLL: TFloatField;
    ActividadesxOrdeniColor: TIntegerField;
    ActividadesxOrdensWbsSpace: TStringField;
    Bitacora: TZReadOnlyQuery;
    MaximoDiario: TZReadOnlyQuery;
    ActividadesxOrdendFechaInicio: TDateField;
    ActividadesxOrdendFechaFinal: TDateField;
    grid_actividades: TcxGrid;
    BVW_Actividades: TcxGridDBTableView;
    sSimbolo: TcxGridDBColumn;
    sWbsSpace: TcxGridDBColumn;
    sNumeroActividad: TcxGridDBColumn;
    dCantidad: TcxGridDBColumn;
    sMedida: TcxGridDBColumn;
    dVentaMN: TcxGridDBColumn;
    dInstalado: TcxGridDBColumn;
    dExcedente: TcxGridDBColumn;
    dPonderado: TcxGridDBColumn;
    grid_actividadesLevel1: TcxGridLevel;
    Panel1: TPanel;
    Label1: TLabel;
    tsNumeroOrden: TDBLookupComboBox;
    btnOk: TBitBtn;
    btnTodos: TBitBtn;
    Panel2: TPanel;
    Progress: TGauge;
    cxStyleRepository1: TcxStyleRepository;
    cxstylePaquete: TcxStyle;
    cxColColor: TcxGridDBColumn;
    cxstyleExedente: TcxStyle;
    cxstyleInstaladoIgualCantidad: TcxStyle;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnOkClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure ActividadesxOrdenCalcFields(DataSet: TDataSet);
    procedure grid_actividadesGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure btnTodosClick(Sender: TObject);
    procedure ActividadesxOrdenAfterScroll(DataSet: TDataSet);
    procedure grid_actividadesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_actividadesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_actividadesTitleClick(Column: TColumn);
    procedure BVW_ActividadesStylesGetContentStyle(
      Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
      AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCalculoAvancesxPartida: TfrmCalculoAvancesxPartida;
  lRegistrarUno: boolean;
  utgrid: ticdbgrid;
  SavePlace     : TBookmark;
implementation

{$R *.dfm}

procedure TfrmCalculoAvancesxPartida.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree;
end;

procedure TfrmCalculoAvancesxPartida.btnOkClick(Sender: TObject);
var
  iDiario: Integer;
  dInstalado: Double;
  dCantidad: Double;
  EsPartidaTerminal: Boolean;
  sWbsKardex: string;
begin
  try

      if ActividadesxOrden.RecordCount > 0 then
      begin
         if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Actividad' then
         begin
            dInstalado := 0;

            // Inicializo Query . Para solo enviar Parametros en el ciclo
            Bitacora.Active := False;
            Bitacora.SQL.Clear;
            Bitacora.SQL.Add('select iIdDiario, dIdFecha, sIdTurno, dCantidad, dAvance, mDescripcion from bitacoradeactividades ' +
              'where sContrato = :contrato and sNumeroOrden = :Orden and dCantidad > 0 And sWbs = :Wbs And sIdTipoMovimiento = "E" and ' +
              'sNumeroActividad = :Actividad and lAlcance = "No" order by dIdFecha, sIdTurno, iIdDiario');
            Bitacora.Params.ParamByName('Contrato').DataType  := ftString;
            Bitacora.Params.ParamByName('Contrato').Value     := global_contrato;
            Bitacora.Params.ParamByName('Orden').DataType     := ftString;
            Bitacora.Params.ParamByName('Orden').Value        := ActividadesxOrden.FieldValues['sNumeroOrden'];
            Bitacora.Params.ParamByName('Wbs').DataType       := ftString;
            Bitacora.Params.ParamByName('Wbs').Value          := ActividadesxOrden.FieldValues['sWbs'];
            Bitacora.Params.ParamByName('Actividad').DataType := ftString;
            Bitacora.Params.ParamByName('Actividad').Value    := ActividadesxOrden.FieldValues['sNumeroActividad'];
            Bitacora.Open;

            if Bitacora.RecordCount > 0 then
            begin
                Progress.Visible := True;
                Progress.MaxValue := Bitacora.RecordCount;
                Progress.MinValue := 1;
            end;

            while not Bitacora.Eof do
            begin
                Progress.Progress := Bitacora.RecNo;
                dInstalado := dInstalado + Bitacora.FieldValues['dCantidad'];
                Bitacora.Next;
            end;
            Progress.Progress := 0;
            Progress.Visible := False;
          end;

          //Ahora los acumulados se almacenan en los catalogos principales ...
          // Primero en el programa de trabajo
          try
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('UPDATE actividadesxorden SET dInstalado = :Instalado, dExcedente = :Excedente ' +
                      'where sContrato = :contrato And sIdConvenio = :Convenio And sNumeroOrden = :Orden And sWbs = :Wbs And sNumeroActividad = :Actividad And sTipoActividad = "Actividad"');
            connection.zCommand.Params.ParamByName('contrato').DataType  := ftString;
            connection.zCommand.Params.ParamByName('contrato').value     := global_contrato;
            connection.zCommand.Params.ParamByName('convenio').DataType  := ftString;
            connection.zCommand.Params.ParamByName('convenio').value     := global_convenio;
            connection.zCommand.Params.ParamByName('Orden').DataType     := ftString;
            connection.zCommand.Params.ParamByName('Orden').value        := ActividadesxOrden.FieldValues['sNumeroOrden'];
            connection.zCommand.Params.ParamByName('Wbs').DataType       := ftString;
            connection.zCommand.Params.ParamByName('Wbs').value          := ActividadesxOrden.FieldValues['sWbs'];
            connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
            connection.zCommand.Params.ParamByName('Actividad').value    := ActividadesxOrden.FieldValues['sNumeroActividad'];
            if (dInstalado > ActividadesxOrden.FieldValues['dCantidad']) then
            begin
              connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('Instalado').value    := ActividadesxOrden.FieldValues['dCantidad'];
              connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('Excedente').value    := dInstalado - ActividadesxOrden.FieldValues['dCantidad'];
            end
            else
            begin
              connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('Instalado').value    := dInstalado;
              connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('Excedente').value    := 0;
            end;
            connection.zCommand.ExecSQL;
          except
            MessageDlg('ERROR: Ocurrio un error al actualizar en el programa la partida No. ' + ActividadesxOrden.FieldValues['sWbs'] + ', notificar al administrador del sistema', mtWarning, [mbOk], 0);
          end;

              // Ahora Ajusto la Partida del Anexo ....
          Connection.qryBusca.Active := False;
          Connection.qryBusca.SQL.Clear;
          Connection.qryBusca.SQL.Add('Select Sum(dInstalado + dExcedente) as dTotal From actividadesxorden Where sContrato = :Contrato And sIdConvenio = :Convenio And ' +
            'sNumeroActividad = :Actividad And sTipoActividad = "Actividad" Group By sNumeroActividad');
          Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
          Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad'];
          Connection.qryBusca.Open;
          if Connection.qryBusca.RecordCount > 0 then
          begin
            Connection.qryBusca2.Active := False;
            Connection.qryBusca2.SQL.Clear;
            Connection.qryBusca2.SQL.Add('Select dCantidadAnexo from actividadesxanexo Where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroActividad = :Actividad');
            Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString;
            Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
            Connection.qryBusca2.Params.ParamByName('Convenio').DataType := ftString;
            Connection.qryBusca2.Params.ParamByName('Convenio').Value := global_convenio;
            Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString;
            Connection.qryBusca2.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad'];
            Connection.qryBusca2.Open;

            if Connection.qryBusca2.RecordCount > 0 then
            begin
              try
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('UPDATE actividadesxanexo SET dInstalado = :Instalado, dExcedente = :Excedente ' +
                  'where sContrato = :contrato And sIdConvenio = :Convenio And sNumeroActividad = :Actividad');
                connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                connection.zCommand.Params.ParamByName('contrato').value := global_contrato;
                connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                connection.zCommand.Params.ParamByName('convenio').value := global_convenio;
                connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
                connection.zCommand.Params.ParamByName('Actividad').value := ActividadesxOrden.FieldValues['sNumeroActividad'];
                if Connection.qryBusca.FieldValues['dTotal'] > Connection.qryBusca2.FieldValues['dCantidadAnexo'] then
                begin
                  connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
                  connection.zCommand.Params.ParamByName('Instalado').value := Connection.qryBusca2.FieldValues['dCantidadAnexo'];
                  connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
                  connection.zCommand.Params.ParamByName('Excedente').value := Connection.qryBusca.FieldValues['dTotal'] - Connection.qryBusca2.FieldValues['dCantidadAnexo'];
                end
                else
                begin
                  connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
                  connection.zCommand.Params.ParamByName('Instalado').value := Connection.qryBusca.FieldValues['dTotal'];
                  connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
                  connection.zCommand.Params.ParamByName('Excedente').value := 0;
                end;
                connection.zCommand.ExecSQL;
              except
                MessageDlg('ERROR: Ocurrio un error al actualizar en el programa la partida No. ' + ActividadesxOrden.FieldValues['sWbs'] + ', notificar al administrador del sistema', mtWarning, [mbOk], 0);
              end
            end
          end ;

         if lRegistrarUno then
         begin
            //registrar la operacion en el kardex
            sWbsKardex := ActividadesxOrden.FieldByName('sWbs').AsString;
            Kardex('Regeneraciones', 'Concepto regenerado', sWbsKardex, 'Partida', tsNumeroOrden.Text, '', '','Tarifa Diaria','Regeneracion de Avance x Cencepto');
         end;

         SavePlace := BVW_Actividades.DataController.DataSource.DataSet.GetBookmark;
         ActividadesxOrden.Active := False ;
         ActividadesxOrden.Open ;
         Try
            BVW_Actividades.DataController.DataSource.DataSet.GotoBookmark(SavePlace);
         Except
         Else
            BVW_Actividades.DataController.DataSet.FreeBookmark(SavePlace);
         End ;
      end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Regeneracion de Avances x Concepto', 'Al regenerar concepto seleccionado', 0);
    end;
  end;

end;

procedure TfrmCalculoAvancesxPartida.FormShow(Sender: TObject);
begin
  try
    connection.configuracion.refresh;
    OrdenesdeTrabajo.Active := False;
    OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString;
    OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := Global_Contrato;
    ordenesdetrabajo.Params.ParamByName('status').DataType := ftString;
    ordenesdetrabajo.Params.ParamByName('status').Value := connection.configuracion.FieldValues['cStatusProceso'];
    OrdenesdeTrabajo.Open;
    if OrdenesdeTrabajo.RecordCount > 0 then
      tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldValues['sNumeroOrden'];
    Progress.Visible := False;
    tsNumeroOrden.SetFocus;
    lRegistrarUno := True;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Regeneracion de Avances x Concepto', 'Al hacer las consultas de inicio', 0);
    end;
  end;
end;

procedure TfrmCalculoAvancesxPartida.tsNumeroOrdenExit(Sender: TObject);
begin
  ActividadesxOrden.Active := False;
  ActividadesxOrden.Params.ParamByName('Contrato').DataType := ftString;
  ActividadesxOrden.Params.ParamByName('Contrato').Value := Global_Contrato;
  ActividadesxOrden.Params.ParamByName('Convenio').DataType := ftString;
  ActividadesxOrden.Params.ParamByName('Convenio').Value := Global_Convenio;
  ActividadesxOrden.Params.ParamByName('Orden').DataType := ftString;
  ActividadesxOrden.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
  ActividadesxOrden.Open;
  tsNumeroOrden.Color := global_color_salida
end;

procedure TfrmCalculoAvancesxPartida.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    grid_actividades.SetFocus
end;

procedure TfrmCalculoAvancesxPartida.tsNumeroOrdenEnter(Sender: TObject);
begin
  tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmCalculoAvancesxPartida.ActividadesxOrdenCalcFields(
  DataSet: TDataSet);
begin
  if ActividadesxOrden.FieldValues['sWbs'] <> Null then
    ActividadesxOrdensWbsSpace.Text := espaces(ActividadesxOrden.FieldValues['iNivel']) + ActividadesxOrden.FieldValues['sWbs']
end;

procedure TfrmCalculoAvancesxPartida.grid_actividadesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x, y);
end;

procedure TfrmCalculoAvancesxPartida.grid_actividadesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TfrmCalculoAvancesxPartida.grid_actividadesTitleClick(
  Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmCalculoAvancesxPartida.btnTodosClick(Sender: TObject);
begin
  try
    lRegistrarUno := False;
    ActividadesxOrden.First;
    while not ActividadesxOrden.Eof do
    begin
      btnOk.Click;
      ActividadesxOrden.Next;
    end;
    //registrar el conjunto en un solo registro de kardex
    Kardex('Regeneraciones', 'Regenera Orden', 'Todas', 'Partida', tsNumeroOrden.Text, '', '','Tarifa Diaria','Regeneracion de Avances x Cencepto');
    lRegistrarUno := True;
    ActividadesxOrden.Refresh;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Regeneracion de Avances x Concepto', 'Al regenerar todos los conceptos', 0);
    end;
  end;
end;

procedure TfrmCalculoAvancesxPartida.BVW_ActividadesStylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
begin
  if ( ActividadesxOrden.State in [dsBrowse] ) and ( ActividadesxOrden.RecordCount > 0 ) then
  begin

    if ARecord.Values[ dInstalado.Index ] = ARecord.Values[ dCantidad.Index ] then
      AStyle := cxstyleInstaladoIgualCantidad;

    if ARecord.Values[ cxColColor.Index ] = 12 then
      AStyle := cxstylePaquete;

    if ARecord.Values[ dExcedente.Index ] > 0 then
      AStyle := cxstyleExedente;

  end;

end;

procedure TfrmCalculoAvancesxPartida.grid_actividadesGetCellParams(
  Sender: TObject; Field: TField; AFont: TFont; var Background: TColor;
  Highlight: Boolean);
begin
  try
    if (Sender as TrxDBGrid).DataSource.DataSet.State = dsBrowse then
      if ActividadesxOrden.RecordCount > 0 then
      begin
        AFont.Color := esColor(ActividadesxOrden.FieldValues['iColor']);
        if (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sTipoActividad').AsString = 'Paquete' then
          Afont.Style := [fsBold]
        else
          if ((Sender as TrxDBGrid).DataSource.DataSet.FieldByName('dExcedente').AsFloat > 0) then
          begin
            Afont.Style := [fsBold, fsItalic];
            AFont.Color := clRed;
          end
      end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Regeneracion de Avances x Concepto', 'Al cambiar de registro', 0);
    end;
  end;
end;

procedure TfrmCalculoAvancesxPartida.ActividadesxOrdenAfterScroll(
  DataSet: TDataSet);
begin
  Grid_Actividades.Hint := ActividadesxOrden.FieldValues['mDescripcion']
end;

end.

