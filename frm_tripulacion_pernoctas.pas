unit frm_tripulacion_pernoctas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms,
  Dialogs, StdCtrls, Mask, DBCtrls, ComCtrls, Grids, DBGrids, global, frm_connection, utilerias,
  DB, ADODB, Buttons, ExtCtrls, frxClass, frxDBSet, ZAbstractRODataset,
  ZDataset, ZAbstractDataset, Controls, Menus, UnitExcepciones, udbgrid, UFunctionsGHH,
  DBDateTimePicker, UnitValidacion, rxToolEdit, rxCurrEdit, RXDBCtrl,
  cxGraphics, cxLookAndFeels, cxLookAndFeelPainters, dxSkinsCore,
  dxSkinDevExpressStyle, dxSkinFoggy, cxButtons;

type
  TfrmTripulacionPernoctas = class(TForm)
    grid_tripulacion: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    tsIdCategoria: TDBLookupComboBox;
    ds_tripulacion: TDataSource;
    btnUpdate: TBitBtn;
    Panel1: TPanel;
    ttripulacion_nacionales: TEdit;
    ttripulacion_extranjeros: TEdit;
    DBTripulacion: TfrxDBDataset;
    btnPrinter: TBitBtn;
    DBTotalesxCategoria: TfrxDBDataset;
    frxTripulacion: TfrxReport;
    ds_categorias: TDataSource;
    categorias: TZReadOnlyQuery;
    tripulaciondiaria: TZQuery;
    btnNuevo: TBitBtn;
    btnDelete: TBitBtn;
    qry_Tripulacion: TZQuery;
    Panel2: TPanel;
    lblTripulacion: TLabel;
    Label4: TLabel;
    txtCantidad: TEdit;
    Label5: TLabel;
    Tripulacion: TZReadOnlyQuery;
    cmdSalir: TButton;
    PopupMenu1: TPopupMenu;
    ripulacinDiariaDiaAnterior1: TMenuItem;
    ZLookTripulacion: TZQuery;
    ds_looktripulacion: TDataSource;
    tsIdTripulacion: TDBLookupComboBox;
    tdIdFecha: TDBDateTimePicker;
    cmbTurnos: TDBLookupComboBox;
    Label3: TLabel;
    ds_turnos: TDataSource;
    QryTurnos: TZQuery;
    qryTripulacionPernoctaFuera: TZQuery;
    dsTripulacionPernoctaFuera: TDataSource;
    GroupBox1: TGroupBox;
    dCantidadPernoctaFuera: TRxDBCalcEdit;
    Label6: TLabel;
    Label7: TLabel;
    DBMemo1: TDBMemo;
    RxDBCalcEdit1: TRxDBCalcEdit;
    Label8: TLabel;
    qryTripulacionPernoctaFuerasContrato: TStringField;
    qryTripulacionPernoctaFuerasIdTurno: TStringField;
    qryTripulacionPernoctaFueradIdFecha: TDateField;
    qryTripulacionPernoctaFueradCantidad: TFloatField;
    qryTripulacionPernoctaFueradCantidadBordo: TFloatField;
    qryTripulacionPernoctaFueramNotas: TMemoField;
    ActualizaPersonal: TMenuItem;
    Label9: TLabel;
    iEspacios: TRxDBCalcEdit;
    qryTripulacionPernoctaFueramEspacios: TMemoField;
    qryTripulacionPernoctaFueraiEspacios: TIntegerField;
    cmdAgregar: TButton;
    GroupBox2: TGroupBox;
    tdIdCategoria: TEdit;
    Label10: TLabel;
    Label11: TLabel;
    tdCategoria: TEdit;
    tAceptar: TButton;
    tCancelar: TButton;
    tAgregar: TcxButton;
    Label12: TLabel;
    tsIdFolio: TDBLookupComboBox;
    tGuardar: TcxButton;
    ds_folios: TDataSource;
    Folios: TZReadOnlyQuery;
    qry_TripulacionsContrato: TStringField;
    qry_TripulaciondIdFecha: TDateField;
    qry_TripulacionsNumeroOrden: TStringField;
    qry_TripulacionsIdCuenta: TStringField;
    qry_TripulacioniIdDiario: TIntegerField;
    qry_TripulaciondCantidad: TFloatField;
    qry_TripulacionsIdCategoria: TStringField;
    qry_TripulacionsDescripcion: TStringField;
    qry_TripulacionsDescripcionFolio: TStringField;
    pernoctas: TZReadOnlyQuery;
    ds_pernoctas: TDataSource;
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdCategoriaKeyPress(Sender: TObject; var Key: Char);
    procedure ttripulacion_nacionalesKeyPress(Sender: TObject;
      var Key: Char);
    procedure ttripulacion_extranjerosKeyPress(Sender: TObject;
      var Key: Char);
    procedure tdIdFechaExit(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnUpdateClick(Sender: TObject);
    procedure qry_tripulacionAfterInsert(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tdIdFechaEnter(Sender: TObject);
    procedure qry_tripulacionCalcFields(DataSet: TDataSet);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnNuevoClick(Sender: TObject);
    procedure tsIdCategoriaExit(Sender: TObject);
    procedure cmdAgregarClick(Sender: TObject);
    procedure Panel2Click(Sender: TObject);
    procedure cmdSalirClick(Sender: TObject);
    procedure btnPrinterClick(Sender: TObject);
    procedure ripulacinDiariaDiaAnterior1Click(Sender: TObject);
    procedure frxTripulacionGetValue(const VarName: string; var Value: Variant);
    procedure tsIdTripulacionExit(Sender: TObject);
    procedure tsIdCategoriaEnter(Sender: TObject);
    procedure tsIdTripulacionEnter(Sender: TObject);
    procedure txtCantidadEnter(Sender: TObject);
    procedure txtCantidadExit(Sender: TObject);
    procedure grid_tripulacionTitleClick(Column: TColumn);
    procedure qry_TripulacionBeforePost(DataSet: TDataSet);
    procedure txtCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure txtCantidadChange(Sender: TObject);
    procedure cmbTurnosExit(Sender: TObject);
    procedure tGuardarClick(Sender: TObject);
    procedure qryTripulacionPernoctaFueraAfterScroll(DataSet: TDataSet);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure ActualizaPersonalClick(Sender: TObject);
    procedure tsIdTripulacionKeyPress(Sender: TObject; var Key: Char);
    procedure tAgregarClick(Sender: TObject);
    procedure tCancelarClick(Sender: TObject);
    procedure tdIdCategoriaKeyPress(Sender: TObject; var Key: Char);
    procedure tdIdCategoriaEnter(Sender: TObject);
    procedure tdIdCategoriaExit(Sender: TObject);
    procedure tdCategoriaEnter(Sender: TObject);
    procedure tdCategoriaExit(Sender: TObject);
    procedure tdCategoriaKeyPress(Sender: TObject; var Key: Char);
    procedure tAceptarClick(Sender: TObject);
    procedure qry_TripulacionAfterScroll(DataSet: TDataSet);
    procedure tsIdFolioKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdFolioEnter(Sender: TObject);
    procedure tsIdFolioExit(Sender: TObject);
    procedure qry_TripulacionBeforeDelete(DataSet: TDataSet);
    procedure qry_TripulacionBeforeEdit(DataSet: TDataSet);
    procedure qry_TripulacionBeforeInsert(DataSet: TDataSet);



  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTripulacionPernoctas: TfrmTripulacionPernoctas;
  sFirma_PEP, sPuesto_PEP, sFirma_Contratista, sPuesto_Contratista: string;
  fechaAntes: tDate;
  utgrid: ticdbgrid;
implementation

uses frm_tripulacion_diaria;

{$R *.dfm}

procedure TfrmTripulacionPernoctas.tdIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
  {if key = #13 then
    tsIdCategoria.SetFocus  }
end;

procedure TfrmTripulacionPernoctas.tsIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
   if key = #13 then
      grid_tripulacion.SetFocus   
end;

procedure TfrmTripulacionPernoctas.tsIdFolioEnter(Sender: TObject);
begin
    tsIdFolio.Color := global_color_entrada;
end;

procedure TfrmTripulacionPernoctas.tsIdFolioExit(Sender: TObject);
begin
    tsIdFolio.Color := global_color_salida;
end;

procedure TfrmTripulacionPernoctas.tsIdFolioKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key =#13 then
       grid_tripulacion.SetFocus;

    qry_tripulacion.Active := False;
    qry_tripulacion.params.ParamByName('contrato').AsString := global_contrato;
    qry_tripulacion.params.ParamByName('fecha').AsDate   := tdIdFecha.Date;
    qry_tripulacion.params.ParamByName('categoria').DataType := ftString;
    if tsIdCategoria.KeyValue = 0 then
       qry_tripulacion.params.ParamByName('categoria').Value := '%'
    else
       qry_tripulacion.params.ParamByName('categoria').Value := tsIdCategoria.KeyValue;
    qry_tripulacion.params.ParamByName('folio').DataType := ftString;
    if tsIdFolio.KeyValue = '<todos>' then
       qry_tripulacion.params.ParamByName('folio').Value := '%'
    else
       qry_tripulacion.params.ParamByName('folio').Value := tsIdFolio.KeyValue;
    qry_tripulacion.Open;

    if qry_Tripulacion.RecordCount = 0 then
    begin
      ttripulacion_nacionales.Text := '0';
      ttripulacion_extranjeros.Text := '0';
    end;
end;

procedure TfrmTripulacionPernoctas.tsIdTripulacionEnter(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_entrada;
end;

procedure TfrmTripulacionPernoctas.tsIdTripulacionExit(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_salida;
  if tsidtripulacion.KeyValue <> null then
    lblTripulacion.Caption := 'Tripulacion ' + tsIdTripulacion.KeyValue;
end;

procedure TfrmTripulacionPernoctas.tsIdTripulacionKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       txtCantidad.SetFocus;
end;

procedure TfrmTripulacionPernoctas.ttripulacion_nacionalesKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    ttripulacion_extranjeros.SetFocus
end;

procedure TfrmTripulacionPernoctas.txtCantidadChange(Sender: TObject);
begin
  tEditChangef(txtCantidad, 'Cantidad');
end;

procedure TfrmTripulacionPernoctas.txtCantidadEnter(Sender: TObject);
begin
  txtcantidad.Color := global_color_entrada;
end;

procedure TfrmTripulacionPernoctas.txtCantidadExit(Sender: TObject);
begin
  txtcantidad.Color := global_color_salida
end;

procedure TfrmTripulacionPernoctas.txtCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTedit(txtCantidad, key) then
    key := #0
  else
    if key =#13 then
       cmdAgregar.SetFocus;
  

end;

procedure TfrmTripulacionPernoctas.ttripulacion_extranjerosKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdidFecha.SetFocus
end;

//Inicio LA forma

procedure TfrmTripulacionPernoctas.FormShow(Sender: TObject);
begin

  UtGrid := TicdbGrid.create(grid_tripulacion);
  tdIdFecha.Date := Date();
  tdIdFecha.SetFocus;

  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('select dFechaVigencia from categorias group by dFechaVigencia order by dFechaVigencia ASC');
  connection.zCommand.Open;

  fechaAntes := date;
  if connection.zCommand.RecordCount > 0 then
  begin
    fechaAntes := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
    while not connection.zCommand.Eof do
    begin
      if tdIdFecha.Date >= connection.zCommand.FieldByName('dFechaVigencia').AsDateTime then
        fechaAntes := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
      connection.zCommand.Next;
    end;
  end;

  Categorias.Active := False;
  Categorias.ParamByName('fecha').AsDate := fechaAntes;
  Categorias.Open;

  Folios.Active := False;
  Folios.ParamByName('contrato').AsString := global_contrato;
  Folios.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  Folios.Open;

  QryTurnos.Active := False;
  QryTurnos.ParamByName('Contrato').AsString := global_contrato;
  QryTurnos.Open;

  cmbTurnos.KeyValue := 'A';

  Tripulacion.Active := False;
  Tripulacion.SQL.Clear;
  Tripulacion.SQL.Add('select * from tripulacion where sContrato = :Contrato and dFechaVigencia =:Fecha order by sIdTripulacion');
  Tripulacion.params.ParamByName('Contrato').DataType := ftString;
  Tripulacion.params.ParamByName('Contrato').Value := Global_Contrato;
  Tripulacion.params.ParamByName('fecha').DataType := ftDate;
  Tripulacion.params.ParamByName('fecha').Value := fechaAntes;
  Tripulacion.Open;

  ZLookTripulacion.Active := False;
  ZLookTripulacion.Open;


  if global_fecha_rd = 0 then
     tdIdFecha.Date := now
  else
     tdIdFecha.Date := global_fecha_rd;
  tdIdFecha.OnExit(sender);
  
  grid_tripulacion.SetFocus;
end;

procedure TfrmTripulacionPernoctas.frxTripulacionGetValue(const VarName: string;
  var Value: Variant);
begin
  if CompareText(VarName, 'FECHA_REPORTE') = 0 then
    Value := global_fecha_barco;

  if CompareText(VarName, 'DIAS_TRANSCURRIDOS') = 0 then
    Value := global_dias_por_transcurrir;

  if CompareText(VarName, 'DIAS_POR_TRANSCURRIR') = 0 then
    Value := global_dias_transcurridos;

  if CompareText(VarName, 'SUPERINTENDENTE') = 0 then
    Value := sSuperIntendente;

  if CompareText(VarName, 'SUPERVISOR') = 0 then
    Value := sSupervisor;

  if CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
    Value := sSupervisorTierra;

  if CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
    Value := sPuestoSuperIntendente;

  if CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
    Value := sPuestoSupervisor;

  if CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
    Value := sPuestoSupervisorTierra;

  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select sIdEmbarcacion from embarcacion_vigencia ' +
    'where sContrato =:Contrato and dFechaInicio <= :Fecha and dFechaFinal >=:Fecha order by dFechaInicio');
  connection.QryBusca.ParamByName('Contrato').AsString := global_contrato_barco;
  connection.QryBusca.ParamByName('Fecha').AsDate := tdIdFecha.date;
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
  begin
    connection.QryBusca2.Active := False;
    connection.QryBusca2.SQL.Clear;
    connection.QryBusca2.SQL.Add('select sDescripcion from pernoctan where sIdPernocta =:Embarcacion ');
    connection.QryBusca2.ParamByName('Embarcacion').AsString := connection.QryBusca.FieldValues['sIdEmbarcacion'];
    connection.QryBusca2.Open;
    if CompareText(VarName, 'EMBARCACION') = 0 then
      Value := connection.QryBusca2.FieldValues['sDescripcion'];
  end;

  //Aqui consultamos que las ordenes esten autorizadas
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select * from reportediario where dIdFecha =:fecha and lStatus  <> "Autorizado" '+
                              'and sContrato <> :Contrato');
  connection.QryBusca.ParamByName('Contrato').AsString := global_contrato_barco;
  connection.QryBusca.ParamByName('Fecha').AsDate      := tdIdFecha.date;
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
  begin
      if CompareText(VarName, 'REPORTES_AUTORIZADOS') = 0 then
         Value := 'No';
  end
  else
  begin
      if CompareText(VarName, 'REPORTES_AUTORIZADOS') = 0 then
         Value := 'Si';
  end;
end;

procedure TfrmTripulacionPernoctas.grid_tripulacionTitleClick(Column: TColumn);
begin
//if grid_tripulacion.datasource.DataSet.IsEmpty=false  then
  if grid_tripulacion.DataSource.DataSet.RecordCount>0 then
   UtGrid.DbGridTitleClick(Column);
end;

//Consultar la Fecha

procedure TfrmTripulacionPernoctas.tdIdFechaExit(Sender: TObject);
begin
  tdidfecha.Color := global_color_salida;
  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('select dFechaVigencia from categorias group by dFechaVigencia order by dFechaVigencia ASC');
  connection.zCommand.Open;

  fechaAntes := date;
  if connection.zCommand.RecordCount > 0 then
  begin
    fechaAntes := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
    while not connection.zCommand.Eof do
    begin
      if tdIdFecha.Date >= connection.zCommand.FieldByName('dFechaVigencia').AsDateTime then
         fechaAntes := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
      connection.zCommand.Next;
    end;
  end;

  Tripulacion.Active := False;
  Tripulacion.SQL.Clear;
  Tripulacion.SQL.Add('select * from tripulacion where sContrato = :Contrato and dFechaVigencia =:Fecha order by sIdTripulacion');
  Tripulacion.params.ParamByName('Contrato').DataType := ftString;
  Tripulacion.params.ParamByName('Contrato').Value    := Global_Contrato;
  Tripulacion.params.ParamByName('fecha').DataType    := ftDate;
  Tripulacion.params.ParamByName('fecha').Value       := fechaAntes;
  Tripulacion.Open;

  Categorias.Active := False;
  Categorias.ParamByName('fecha').AsDate := fechaAntes;
  Categorias.Open;

  Folios.Active := False;
  Folios.ParamByName('contrato').AsString := global_contrato;
  Folios.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  Folios.Open;

  tsIdCategoria.KeyValue := 0;
  tsIdFolio.KeyValue := '<todos>';

  qry_tripulacion.Active := False;
  qry_tripulacion.params.ParamByName('contrato').asString  := global_contrato;
  qry_tripulacion.params.ParamByName('fecha').AsDate     := tdIdFecha.Date;
  qry_tripulacion.params.ParamByName('categoria').DataType := ftString;
  if tsIdCategoria.KeyValue = 0 then
     qry_tripulacion.params.ParamByName('categoria').Value := '%'
  else
     qry_tripulacion.params.ParamByName('categoria').Value := tsIdCategoria.KeyValue;
  qry_tripulacion.params.ParamByName('folio').DataType := ftString;
  if tsIdFolio.KeyValue = '<todos>' then
     qry_tripulacion.params.ParamByName('folio').Value := '%'
  else
     qry_tripulacion.params.ParamByName('folio').Value := tsIdFolio.KeyValue;
  qry_tripulacion.Open;

  if qry_Tripulacion.RecordCount = 0 then
  begin
    ttripulacion_nacionales.Text := '0';
    ttripulacion_extranjeros.Text := '0';
    qry_Tripulacion.Close;
  end;


  qryTripulacionPernoctaFuera.Active := false;
  qryTripulacionPernoctaFuera.params.ParamByName('contrato').DataType := ftString;
  qryTripulacionPernoctaFuera.params.ParamByName('contrato').Value := global_contrato_barco;
  qryTripulacionPernoctaFuera.params.ParamByName('fecha').DataType := ftDate;
  qryTripulacionPernoctaFuera.params.ParamByName('fecha').Value := tdIdFecha.Date;
  qryTripulacionPernoctaFuera.params.ParamByName('turno').DataType := ftString;
  qryTripulacionPernoctaFuera.params.ParamByName('turno').Value := QryTurnos.FieldValues['sIdTurno'];
  qryTripulacionPernoctaFuera.Open;

  //Consultamos los tipos de personal que se deben mostrar en personal y equipo..
  connection.QryBusca.Active := FalsE;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select * from reportediario where dIdFecha =:Fecha and sContrato <> :contrato group by sContrato ');
  connection.QryBusca.ParamByName('fecha').AsDate      := tdIdfecha.Date;
  connection.QryBusca.ParamByName('contrato').AsString := connection.contrato.FieldValues['sCodigo'];
  connection.QryBusca.Open;

  grid_tripulacion.Columns[2].PickList.add('');
  while not connection.QryBusca.Eof do
  begin
      grid_tripulacion.Columns[2].PickList.add(connection.QryBusca.FieldValues['sContrato']);
      connection.QryBusca.Next;
  end;

end;


//Consultar la Categoria

procedure TfrmTripulacionPernoctas.ActualizaPersonalClick(Sender: TObject);
var
    zPersonal, zInserta, zBusca, zTripulacion  : tzReadOnlyQuery;
    lEncuentra : boolean;
    Categoria  : string;
begin
    zPersonal := tzReadOnlyQuery.Create(self);
    zPersonal.Connection := connection.zConnection;

    zInserta := tzReadOnlyQuery.Create(self);
    zInserta.Connection := connection.zConnection;

    zBusca := tzReadOnlyQuery.Create(self);
    zBusca.Connection := connection.zConnection;

    zTripulacion := tzReadOnlyQuery.Create(self);
    zTripulacion.Connection := connection.zConnection;

    if qry_tripulacion.RecordCount > 0 then
    begin
        zTripulacion.Active := False;
        zTripulacion.SQL.Clear;
        zTripulacion.SQL.Add('select * from tripulaciondiaria where sContrato =:Contrato and sIdTurno =:Turno and dIdFecha =:fecha ');
        zTripulacion.ParamByName('contrato').AsString := global_contrato_barco;
        zTripulacion.ParamByName('turno').AsString    := QryTurnos.FieldValues['sIdTurno'];
        zTripulacion.ParamByName('fecha').AsDate      := tdIdFecha.Date;
        zTripulacion.Open;

        zPersonal.Active := False;
        zPersonal.SQL.Clear;
        zPersonal.SQL.Add('select p.sAgrupaPersonal, sum(b.dCantidad) as dCantidad, p.sDescripcion from bitacoradepersonal b '+
                          'inner join personal p on(b.sContrato = p.sContrato  and p.sIdPersonal = b.sIdPersonal) '+
                          'where b.dIdFecha  = :fecha and sIdTipoPersonal = "PE-C" group by b.sAgrupaPersonal order by p.iItemOrden');
        zPersonal.ParamByName('fecha').AsDate := tdIdFecha.Date;
        zPersonal.Open;

        //Personal de reportes diarios..
        while not zPersonal.Eof do
        begin
            //Personal de tripulacion..
            lEncuentra :=  False;
            zTripulacion.First;
            while not zTripulacion.Eof do
            begin
                if zTripulacion.FieldValues['sIdTripulacion'] = zPersonal.FieldValues['sAgrupaPersonal']  then
                begin
                    lEncuentra := True;
                    Categoria  := zTripulacion.FieldValues['sIdCategoria'];
                    zTripulacion.Last;
                end;
                zTripulacion.Next;
            end;

            //Verificamos si se encontro la categoria de personal en la zTripulacion diaria,
            if lEncuentra  then
            begin
                zInserta.Active := False;
                zInserta.SQL.Clear;
                zInserta.SQL.Add('Update tripulaciondiaria set iNacionales =:cantidad where sContrato =:Contrato and sIdTurno =:Turno '+
                                  'and dIdFecha =:Fecha and sIdCategoria =:Categoria and sIdTripulacion =:Tripulacion');
                zInserta.ParamByName('contrato').AsString    := global_contrato_barco;
                zInserta.ParamByName('turno').AsString       := QryTurnos.FieldValues['sIdTurno'];
                zInserta.ParamByName('fecha').AsDate         := tdIdFecha.Date;
                zInserta.ParamByName('categoria').AsString   := Categoria;
                zInserta.ParamByName('cantidad').AsFloat     := zPersonal.FieldValues['dCantidad'];
                zInserta.ParamByName('tripulacion').AsString := zPersonal.FieldValues['sAgrupaPersonal'];
                zInserta.ExecSQL;
            end
            else
            begin
                //Buscamos la categoria de la tripulacion que sea del Anexo de Personal,,
                zBusca.Active := False;
                zBusca.SQL.Clear;
                zBusca.SQL.Add('select sIdCategoria from categorias where dFechaVigencia =:fecha and lPersonalAnexo = "Si" ');
                zBusca.ParamByName('fecha').AsDate   := fechaAntes;
                zBusca.Open;

                if zBusca.RecordCount > 0 then
                   categoria := zBusca.FieldValues['sIdCategoria']
                else
                begin
                    messageDLG('No existe una Categoría que Considere Personal de Anexo "Si", Ver Administración de Catálogos [Categorías]', mtInformation, [mbOk], 0);
                    exit;
                end;

                //Ahora buscamos la categoria en el catalogo de tripulacion..
                zBusca.Active := False;
                zBusca.SQL.Clear;
                zBusca.SQL.Add('select sIdTripulacion from tripulacion where sContrato =:Contrato and dFechaVigencia =:fecha and sIdCategoria =:Categoria and sIdTripulacion =:Tripulacion ');
                zBusca.ParamByName('contrato').AsString    := global_contrato_barco;
                zBusca.ParamByName('fecha').AsDate         := fechaAntes;
                zBusca.ParamByName('categoria').AsString   := categoria;
                zBusca.ParamByName('tripulacion').AsString := zPersonal.FieldValues['sAgrupaPersonal'];
                zBusca.Open;

                //Si no se encuentra la damos de alta en el catalogo..
                if zBusca.RecordCount = 0 then
                begin
                    zInserta.Active := False;
                    zInserta.SQL.Clear;
                    zInserta.SQL.Add('Insert into tripulacion (sContrato, dFechaVigencia, sIdCategoria, sIdTripulacion, sDescripcion, iNacionales, iExtranjeros ) '+
                                      'values (:Contrato, :fecha, :categoria, :tripulacion, :descripcion, 0, 0)');
                    zInserta.ParamByName('contrato').DataType    := ftString;
                    zInserta.ParamByName('contrato').value       := Global_Contrato_barco;
                    zInserta.ParamByName('fecha').DataType       := ftDate;
                    zInserta.ParamByName('fecha').value          := fechaAntes;
                    zInserta.ParamByName('categoria').DataType   := ftString;
                    zInserta.ParamByName('categoria').value      := categoria;
                    zInserta.ParamByName('tripulacion').DataType := ftString;
                    zInserta.ParamByName('tripulacion').value    := zPersonal.FieldValues['sAgrupaPersonal'];
                    zInserta.ParamByName('Descripcion').DataType := ftString;
                    zInserta.ParamByName('Descripcion').value    := zPersonal.FieldValues['sAgrupaPersonal'] +' '+zPersonal.FieldValues['sDescripcion'];
                    zInserta.ExecSQL;
                end;

                //Insertamos la tripulaicon diaria..
                zInserta.Active := False;
                zInserta.SQL.Clear;
                zInserta.SQL.Add('Insert into tripulaciondiaria (sContrato, sIdTurno, dIdFecha, sIdCategoria, sIdTripulacion, iNacionales, iExtranjeros ) '+
                                  'values (:Contrato, :turno, :fecha, :categoria, :tripulacion, :nacionales, :extranjeros)');
                zInserta.ParamByName('contrato').DataType    := ftString;
                zInserta.ParamByName('contrato').value       := Global_Contrato_barco;
                zInserta.ParamByName('Turno').DataType       := ftString;
                zInserta.ParamByName('Turno').value          := QryTurnos.FieldValues['sIdTurno'];
                zInserta.ParamByName('fecha').DataType       := ftDate;
                zInserta.ParamByName('fecha').value          := tdIdFecha.Date;
                zInserta.ParamByName('categoria').DataType   := ftString;
                zInserta.ParamByName('categoria').value      := categoria;
                zInserta.ParamByName('tripulacion').DataType := ftString;
                zInserta.ParamByName('tripulacion').value    := zPersonal.FieldValues['sAgrupaPersonal'];
                zInserta.ParamByName('Nacionales').DataType  := ftInteger;
                zInserta.ParamByName('Nacionales').value     := zPersonal.FieldValues['dCantidad'];
                zInserta.ParamByName('Extranjeros').DataType := ftInteger;
                zInserta.ParamByName('Extranjeros').value    := 0;
                zInserta.ExecSQL;
            end;
            zPersonal.Next;
        end;

        //Sino se encontro personal reportado y se hizo el recalculo..
        if zPersonal.RecordCount = 0 then
        begin
            zPersonal.Active := False;
            zPersonal.SQL.Clear;
            zPersonal.SQL.Add('select sAgrupaPersonal from personal where sContrato =:Contrato ');
            zPersonal.ParamByName('Contrato').AsString := global_contrato_barco;
            zPersonal.Open;

            zBusca.Active := False;
            zBusca.SQL.Clear;
            zBusca.SQL.Add('select * from tripulaciondiaria where sContrato =:Contrato and sIdTurno =:Turno and dIdFecha =:fecha ');
            zBusca.ParamByName('contrato').AsString := global_contrato_barco;
            zBusca.ParamByName('fecha').AsDate      := tdIdFecha.Date;
            zBusca.ParamByName('Turno').AsString    := Qryturnos.FieldValues['sIdTurno'] ;
            zBusca.Open;

            while not zBusca.Eof do
            begin
                zPersonal.First;
                while not zPersonal.Eof do
                begin
                    if zPersonal.FieldValues['sAgrupaPersonal'] = zBusca.FieldValues['sIdTripulacion'] then
                    begin
                        zInserta.Active := False;
                        zInserta.SQL.Clear;
                        zInserta.SQL.Add('Update tripulaciondiaria set iNacionales =:cantidad where sContrato =:Contrato and sIdTurno =:Turno '+
                                         'and dIdFecha =:Fecha and sIdCategoria =:Categoria and sIdTripulacion =:Tripulacion');
                        zInserta.ParamByName('contrato').AsString    := global_contrato_barco;
                        zInserta.ParamByName('turno').AsString       := zBusca.FieldValues['sIdTurno'];
                        zInserta.ParamByName('fecha').AsDate         := zBusca.FieldValues['dIdFecha'];
                        zInserta.ParamByName('categoria').AsString   := zBusca.FieldValues['sIdCategoria'];
                        zInserta.ParamByName('cantidad').AsFloat     := 0;
                        zInserta.ParamByName('tripulacion').AsString := zBusca.FieldValues['sIdTripulacion'];
                        zInserta.ExecSQL;
                    end;
                    zPersonal.Next;
                end;
                zBusca.Next;
            end;
        end;
    end
    else
       messageDLG('Primero debe Agregar Personal de Tripulacion!', mtInformation, [mbOk], 0);

    btnUpdate.Click;
    Qry_Tripulacion.Refresh;
    Qry_tripulacion.First;

    zPersonal.Destroy;
    zInserta.Destroy;
    zBusca.Destroy;
    zTripulacion.Destroy;
end;

procedure TfrmTripulacionPernoctas.btnCancelClick(Sender: TObject);
begin
  close;
end;

procedure TfrmTripulacionPernoctas.btnUpdateClick(Sender: TObject);
begin
  if (not qry_Tripulacion.active) or (qry_Tripulacion.RecordCount < 1) then
    exit;
  ttripulacion_nacionales.Text := '0';
  ttripulacion_extranjeros.Text := '0';
  qry_tripulacion.First;
  try
    Connection.zcommand.Active := False;
    Connection.zcommand.SQL.Clear;
    Connection.zcommand.SQL.Add('begin');
    Connection.zcommand.ExecSQL;
    while not qry_tripulacion.eof do
    begin
      ttripulacion_nacionales.Text := inttostr(strtoint(ttripulacion_nacionales.Text) + qry_tripulacion.FieldValues['iNacionales']);
      ttripulacion_extranjeros.Text := inttostr(strtoint(ttripulacion_extranjeros.Text) + qry_tripulacion.FieldValues['iExtranjeros']);

      Connection.zcommand.Active := False;
      Connection.zcommand.SQL.Clear;
      Connection.zcommand.SQL.Add('UPDATE tripulaciondiaria SET iNacionales = :nacionales , ' +
        'iExtranjeros = :extranjeros WHERE sContrato = :contrato and ' +
        'dIdFecha = :fecha and sIdCategoria = :categoria and sIdTripulacion = :tripulacion and sIdTurno =:Turno ');
      Connection.zcommand.params.ParamByName('nacionales').DataType := ftInteger;
      Connection.zcommand.params.ParamByName('nacionales').value := qry_tripulacion.FieldValues['iNacionales'];
      Connection.zcommand.params.ParamByName('extranjeros').DataType := ftInteger;
      Connection.zcommand.params.ParamByName('extranjeros').value := qry_tripulacion.FieldValues['iExtranjeros'];
      Connection.zcommand.params.ParamByName('contrato').DataType := ftString;
      Connection.zcommand.params.ParamByName('contrato').value := global_contrato_barco;
      Connection.zcommand.params.ParamByName('Turno').DataType := ftString;
      Connection.zcommand.params.ParamByName('Turno').value := QryTurnos.FieldValues['sIdTurno'];
      Connection.zcommand.params.ParamByName('fecha').DataType := ftDate;
      Connection.zcommand.params.ParamByName('fecha').value := tdIdFecha.Date;
      Connection.zcommand.params.ParamByName('categoria').DataType := ftString;
      Connection.zcommand.params.ParamByName('categoria').value := tsIdCategoria.KeyValue;
      Connection.zcommand.params.ParamByName('tripulacion').DataType := ftString;
      Connection.zcommand.params.ParamByName('tripulacion').value := qry_tripulacion.FieldValues['sIdTripulacion'];
      Connection.zcommand.ExecSQL;
      qry_tripulacion.Next
    end;
    Connection.zcommand.Active := False;
    Connection.zcommand.SQL.Clear;
    Connection.zcommand.SQL.Add('commit');
    Connection.zcommand.ExecSQL;
  except;
    ShowMessage('Error al actualizar !!');
    Connection.zcommand.Active := False;
    Connection.zcommand.SQL.Clear;
    Connection.zcommand.SQL.Add('rollback');
    Connection.zcommand.ExecSQL;
  end;
  qry_tripulacion.Refresh;
end;

procedure TfrmTripulacionPernoctas.tGuardarClick(Sender: TObject);
var
   espacios : string;
   i        : integer;
begin
  try

    btnUpdate.Click;

    if qryTripulacionPernoctaFuera.RecordCount > 0 then
      qryTripulacionPernoctaFuera.Edit
    else
    begin
       // qryTripulacionPernoctaFuera.Append;
       // qryTripulacionPernoctaFuera.FieldValues['sContrato'] := global_contrato;
       // qryTripulacionPernoctaFuera.FieldValues['sIdTurno'] := QryTurnos.FieldValues['sIdTurno'];
       // qryTripulacionPernoctaFuera.FieldValues['dIdFecha'] := tdIdFecha.Date;
    end;

    i := 1;
    espacios := '';
    while i < iEspacios.Value do
    begin
        espacios := espacios + #13;
        inc(i);
    end;

    if qryTripulacionPernoctaFuera.State in [dsInsert, dsEdit] then
    begin
      qryTripulacionPernoctaFuera.FieldValues['sContrato']  := global_contrato_barco;
      qryTripulacionPernoctaFuera.FieldValues['sIdTurno']   := QryTurnos.FieldValues['sIdTurno'];
      qryTripulacionPernoctaFuera.FieldValues['dIdFecha']   := tdIdFecha.Date;
      qryTripulacionPernoctaFuera.FieldValues['dCantidad']  := dCantidadPernoctaFuera.Value;
      qryTripulacionPernoctaFuera.FieldValues['dCantidadBordo'] := strtoint(ttripulacion_nacionales.Text) - dCantidadPernoctaFuera.Value;
      qryTripulacionPernoctaFuera.FieldValues['mEspacios']  := espacios;
      qryTripulacionPernoctaFuera.Post;
      qryTripulacionPernoctaFuera.Refresh;
    end;

  except
    on e: exception do begin
      MessageDlg(E.Message, mtError, [mbOk], 0);
//      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Tripulacion Diaria', 'Al salvar registro', 0);
    end;
  end;

end;

procedure TfrmTripulacionPernoctas.qryTripulacionPernoctaFueraAfterScroll(
  DataSet: TDataSet);
begin
(*      qryTripulacionPernoctaFuera.FieldValues['sContrato'] := global_contrato;
      qryTripulacionPernoctaFuera.FieldValues['sIdTurno'] := QryTurnos.FieldValues['sIdTurno'];
      qryTripulacionPernoctaFuera.FieldValues['dIdFecha'] := tdIdFecha.Date;*)
end;

procedure TfrmTripulacionPernoctas.qry_tripulacionAfterInsert(
  DataSet: TDataSet);
begin
  qry_tripulacion.Cancel
end;

procedure TfrmTripulacionPernoctas.qry_TripulacionAfterScroll(DataSet: TDataSet);
begin
  with connection.QryBusca do
  begin
    Active := False;
    SQL.Clear;
    SQL.Add('SELECT sum(dCantidad) as cantidad FROM bitacoradepernocta where sContrato = :contrato and dIdfecha = :fecha ' +
      'And sIdCategoria like :Categoria and sNumeroOrden like :Orden group by sContrato ');
    params.ParamByName('contrato').DataType := ftString;
    params.ParamByName('contrato').Value := global_contrato;
    params.ParamByName('fecha').DataType := ftDate;
    params.ParamByName('fecha').Value := tdIdFecha.Date;
    params.ParamByName('categoria').DataType := ftString;
    if tsIdCategoria.KeyValue = 0 then
       params.ParamByName('categoria').Value := '%'
    else
       params.ParamByName('categoria').Value := tsIdCategoria.KeyValue;
    params.ParamByName('Orden').DataType := ftString;
    if tsIdFolio.KeyValue = '<todos>' then
       params.ParamByName('Orden').Value := '%'
    else
       params.ParamByName('Orden').Value := tsIdFolio.KeyValue;
    Open;
    if qry_Tripulacion.RecordCount > 0 then
    begin
        ttripulacion_nacionales.Text  := FloatTostr(fieldbyname('cantidad').AsFloat);
        ttripulacion_extranjeros.Text := floatToStr(fieldbyName('cantidad').AsFloat);
    end;
  end;
end;

procedure TfrmTripulacionPernoctas.qry_TripulacionBeforeDelete(
  DataSet: TDataSet);
begin
  If ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
  begin
      MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
      abort;
  end;
end;

procedure TfrmTripulacionPernoctas.qry_TripulacionBeforeEdit(DataSet: TDataSet);
begin
  If ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
  begin
      MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
      abort;
  end;
end;

procedure TfrmTripulacionPernoctas.qry_TripulacionBeforeInsert(
  DataSet: TDataSet);
begin
  If ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
  begin
      MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
      abort;
  end;
end;

procedure TfrmTripulacionPernoctas.qry_TripulacionBeforePost(DataSet: TDataSet);
begin
  if qry_Tripulacion.FieldValues['dCantidad'] < 0 then
    qry_Tripulacion.cancel;

end;

procedure TfrmTripulacionPernoctas.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  utgrid.Destroy;
  action := cafree;
end;

procedure TfrmTripulacionPernoctas.tAceptarClick(Sender: TObject);
begin
    ZLookTripulacion.FieldByName('sIdTripulacion').AsString      := tdIdCategoria.Text;
    ZLookTripulacion.FieldByName('sDescripcion').AsString        := tdCategoria.Text;
    ZLookTripulacion.FieldByName('sIdTripulacionGrupo').AsString := tdIdCategoria.Text;
    ZLookTripulacion.FieldByName('sDescripcionGrupo').AsString   := tdCategoria.Text;
    ZLookTripulacion.Post;
    ZLookTripulacion.Refresh;
    tsIdTripulacion.KeyValue := tdIdCategoria.Text;
    tCancelar.Click;
end;

procedure TfrmTripulacionPernoctas.tAgregarClick(Sender: TObject);
begin
   Panel2.Height := 260;
   Panel2.Width  := 337;
   cmdAgregar.Enabled := False;
   cmdSalir.Enabled   := False;
   tsIdTripulacion.Enabled := False;
   txtCantidad.Enabled     := FAlse;
   tdIdCategoria.Text := '';
   tdCategoria.Text  := '';
   ZLookTripulacion.Append;
   ZLookTripulacion.FieldByName('sContrato').AsString        := global_contrato;
   ZLookTripulacion.FieldByName('dFechaVigencia').AsDateTime := fechaAntes;
   ZLookTripulacion.FieldByName('sIdCategoria').AsString     := categorias.FieldByName('sIdCategoria').AsString;
   ZLookTripulacion.FieldByName('iNacionales').AsInteger     := 0;
   ZLookTripulacion.FieldByName('iExtranjeros').AsInteger    := 0;
   ZLookTripulacion.FieldByName('iOrden').AsInteger          :=  ZLookTripulacion.RecordCount + 1;
   tdIdCategoria.Text := categorias.FieldByName('sMascara').AsString + IntToStr(ZLookTripulacion.RecordCount + 1);
   tdIdCategoria.SetFocus;
end;

procedure TfrmTripulacionPernoctas.tCancelarClick(Sender: TObject);
begin
   Panel2.Height := 123;
   Panel2.Width  := 337;
   cmdAgregar.Enabled := True;
   cmdSalir.Enabled   := True;
   tsIdTripulacion.Enabled := True;
   txtCantidad.Enabled     := True;
   ZLookTripulacion.Cancel;
   tsIdTripulacion.SetFocus;
end;

procedure TfrmTripulacionPernoctas.tdCategoriaEnter(Sender: TObject);
begin
   tdCategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionPernoctas.tdCategoriaExit(Sender: TObject);
begin
   tdCategoria.Color := global_color_salida;
end;

procedure TfrmTripulacionPernoctas.tdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tAceptar.SetFocus
end;

procedure TfrmTripulacionPernoctas.tdIdCategoriaEnter(Sender: TObject);
begin
   tdIdCategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionPernoctas.tdIdCategoriaExit(Sender: TObject);
begin
   tdIdCategoria.Color := global_color_salida;
end;

procedure TfrmTripulacionPernoctas.tdIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tdcategoria.SetFocus;
end;

procedure TfrmTripulacionPernoctas.tdIdFechaEnter(Sender: TObject);
begin
  tdIdFecha.Color := global_Color_entrada
end;


procedure TfrmTripulacionPernoctas.qry_tripulacionCalcFields(
  DataSet: TDataSet);
begin
  if qry_tripulacion.RecordCount > 0 then
  begin
        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Add('Select sDescripcion from categorias Where ' +
          ' sIdCategoria = :Categoria and dFechaVigencia =:Fecha ');
        Connection.QryBusca.Params.ParamByName('Categoria').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Categoria').Value := qry_Tripulacion.FieldValues['sIdCategoria'];
        Connection.QryBusca.Params.ParamByName('Fecha').DataType := ftDate;
        Connection.QryBusca.Params.ParamByName('Fecha').Value := fechaAntes;
        Connection.QryBusca.Open;
        if Connection.QryBusca.RecordCount > 0 then
          qry_TripulacionsDescripcion.Text := Connection.QryBusca.FieldValues['sDescripcion'];

        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Add('Select sIdFolio from ordenesdetrabajo Where ' +
          ' sContrato = :Contrato and sNumeroOrden =:Orden ');
        Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Contrato').Value := Global_Contrato;
        Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Orden').Value := qry_Tripulacion.FieldValues['sNumeroOrden'];
        Connection.QryBusca.Open;
        if Connection.QryBusca.RecordCount > 0 then
          qry_TripulacionsDescripcionFolio.Text := Connection.QryBusca.FieldValues['sIdFolio'];
  end;
end;



procedure TfrmTripulacionPernoctas.ripulacinDiariaDiaAnterior1Click(
  Sender: TObject);
var
  dFechaAnterior : tDate;
  lNuevaVigencia : boolean;

begin
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('Select * from bitacoradepernocta Where sContrato =:Contrato And dIdFecha =:Fecha');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value    := Global_Contrato;
  Connection.qryBusca.Params.ParamByName('Fecha').Datatype    := ftDate;
  Connection.qryBusca.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
  Connection.qryBusca.Open;
  if Connection.qryBusca.RecordCount > 0 then
    if MessageDlg('Desea Eliminar todas las Pernoctas Asignadas?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('Delete From bitacoradepernoctas where sContrato = :contrato and dIdFecha = :fecha  ');
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
      connection.zCommand.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
      connection.zCommand.ExecSQL;
    end;

  if MessageDlg('Desea adicionar las pernoctas Existentes en el reporte anterior?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
      //Ahora verificamos si las vigencias son diferentes,
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('select dFechaVigencia from categorias group by dFechaVigencia order by dFechaVigencia ASC');
      connection.zCommand.Open;

      if connection.zCommand.RecordCount > 0 then
      begin
          dFechaAnterior := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
          while not connection.zCommand.Eof do
          begin
              if tdIdFecha.Date  - 1 >= connection.zCommand.FieldByName('dFechaVigencia').AsDateTime then
                 dFechaAnterior := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
              connection.zCommand.Next;
          end;

          if dFechaAnterior <> fechaAntes then
          begin
              messageDLG('Existe una Nueva Vigencia de personal con fecha del '+DateToStr(fechaAntes)+' diefrente al Dia Anterior. Los Datos se insertaran en 0', mtInformation, [mbOk], 0);

              Connection.Auxiliar.SQL.Clear;
              Connection.Auxiliar.SQL.Add('Select * from tripulacion Where dFechaVigencia =:Fecha');
              Connection.Auxiliar.Params.ParamByName('Fecha').Datatype    := ftDate;
              Connection.Auxiliar.Params.ParamByName('Fecha').Value       := FechaAntes;
              Connection.Auxiliar.Open;
          end
          else
          begin
              dFechaAnterior := tdIdFecha.Date - 1;
              Connection.Auxiliar.SQL.Clear;
              Connection.Auxiliar.SQL.Add('Select * from bitacoradepernocta Where sContrato =:Contrato And dIdFecha =:Fecha');
              Connection.Auxiliar.Params.ParamByName('Contrato').DataType := ftString;
              Connection.Auxiliar.Params.ParamByName('Contrato').Value    := Global_Contrato;
              Connection.Auxiliar.Params.ParamByName('Fecha').Datatype    := ftDate;
              Connection.Auxiliar.Params.ParamByName('Fecha').Value       := dFechaAnterior;
              Connection.Auxiliar.Open;
          end;
      end;

      if Connection.Auxiliar.RecordCount > 0 then
      begin
          Connection.zcommand.SQL.Clear;
          Connection.zcommand.SQL.Add('INSERT INTO bitacoradepernocta ( sContrato , didFecha , sIdCategoria, sNumeroOrden, sIdCuenta, dCantidad, iIdDiario ) ' +
            ' VALUES (:contrato , :fecha  , :categoria , :Orden , :cuenta , :Cantidad, 0 )');
          while not Connection.Auxiliar.Eof do
          begin
            with connection do
            begin
                try
                  zcommand.params.ParamByName('contrato').DataType    := ftString;
                  zcommand.params.ParamByName('contrato').value       := Global_Contrato;
                  zcommand.params.ParamByName('fecha').DataType       := ftDate;
                  zcommand.params.ParamByName('fecha').value          := tdIdFecha.Date;
                  zcommand.params.ParamByName('categoria').DataType   := ftString;
                  zcommand.params.ParamByName('categoria').value      := Auxiliar.FieldValues['sIdCategoria'];
                  zcommand.params.ParamByName('orden').DataType       := ftString;
                  zcommand.params.ParamByName('orden').value          := Auxiliar.FieldValues['sNumeroOrden'];
                  zcommand.params.ParamByName('cuenta').DataType      := ftString;
                  zcommand.params.ParamByName('cuenta').value         := Auxiliar.FieldValues['sIdCuenta'];
                  zcommand.params.ParamByName('cantidad').DataType    := ftInteger;
                  zcommand.params.ParamByName('cantidad').value       := Auxiliar.FieldValues['dCantidad'];
                  zcommand.ExecSQL;
                except
                end;
                Auxiliar.Next;
            end;
          end;
      end
      else
         messageDLG('No se encontró personal el día anterior', mtInformation, [mbOk], 0);
  end;
  tripulacion.Refresh;
  tdIdFecha.OnExit(sender);
end;

procedure TfrmTripulacionPernoctas.btnDeleteClick(Sender: TObject);
begin
  if (not qry_Tripulacion.active) or (qry_Tripulacion.RecordCount < 1) then
    exit;
  Connection.zCommand.Active := False;
  Connection.zCommand.SQL.Clear;
  Connection.zCommand.SQL.Add('Delete from bitacoradepernocta where sContrato = :contrato and sIdCategoria = :Categoria ' +
    'And dIdFecha =:Fecha And sIdCuenta =:Cuenta and sNumeroOrden =:Orden ');
  Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
  Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
  Connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
  Connection.zCommand.Params.ParamByName('Orden').Value := qry_Tripulacion.FieldValues['sNumeroOrden'];
  Connection.zCommand.Params.ParamByName('Categoria').DataType := ftString;
  Connection.zCommand.Params.ParamByName('Categoria').Value := qry_Tripulacion.FieldValues['sIdCategoria'];
  Connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
  Connection.zCommand.Params.ParamByName('Fecha').Value := tdIdFecha.date;
  Connection.zCommand.Params.ParamByName('Cuenta').DataType := ftString;
  Connection.zCommand.Params.ParamByName('Cuenta').Value := qry_Tripulacion.FieldValues['sIdCuenta'];
  Connection.zCommand.ExecSQL();
  qry_Tripulacion.Refresh;
end;

procedure TfrmTripulacionPernoctas.btnNuevoClick(Sender: TObject);
begin
  if (tsIdCategoria.KeyValue = null) or (tsIdCategoria.KeyValue = 0) then begin
    ShowMessage('Favor de Seleccionar una Categoria');
    exit;
  end;

  if (tsIdFolio.KeyValue = null) or (tsIdFolio.KeyValue = '<todos>') then begin
    ShowMessage('Favor de Seleccionar un Folio');
    exit;
  end;

  ZLookTripulacion.First;
  ZLookTripulacion.Locate('sIdPadre', categorias.FieldByName('sTdPu').AsString, [loCaseInsensitive]);
  tsIdTripulacion.keyvalue := ZLookTripulacion.FieldValues['sIdCuenta'];
  txtCantidad.Text := '0';
  Panel2.Visible := True;
  tsIdTripulacion.SetFocus;
end;


procedure TfrmTripulacionPernoctas.btnPrinterClick(Sender: TObject);
begin
  try
      if (qry_Tripulacion.Active) and (qry_Tripulacion.RecordCount > 0) then
        procreporteTripulacion(Global_Contrato_barco, QryTurnos.FieldValues['sIdTurno'], tdIdFecha.DateTime, frmTripulacionDiaria, frxTripulacion.OnGetValue, connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, 'reporteBarco'), 'Barco')
      else
        showmessage('No hay datos para imprimir');
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Control de pernoctas', 'Al imprimir', 0);
    end;
  end;
end;

procedure TfrmTripulacionPernoctas.tsIdCategoriaEnter(Sender: TObject);
begin
  tsidcategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionPernoctas.tsIdCategoriaExit(Sender: TObject);
begin
    tsIdcategoria.Color := global_Color_salida;

    qry_tripulacion.Active := False;
    qry_tripulacion.params.ParamByName('contrato').asString  := global_contrato;
    qry_tripulacion.params.ParamByName('fecha').asDate       := tdIdFecha.Date;
    qry_tripulacion.params.ParamByName('categoria').DataType := ftString;
    if tsIdCategoria.KeyValue = 0 then
       qry_tripulacion.params.ParamByName('categoria').Value := '%'
    else
       qry_tripulacion.params.ParamByName('categoria').Value := tsIdCategoria.KeyValue;
    qry_tripulacion.params.ParamByName('folio').DataType := ftString;
    if tsIdFolio.KeyValue = '<todos>' then
       qry_tripulacion.params.ParamByName('folio').Value := '%'
    else
       qry_tripulacion.params.ParamByName('folio').Value := tsIdFolio.KeyValue;
    qry_tripulacion.Open;

    if qry_Tripulacion.RecordCount = 0 then
    begin
      ttripulacion_nacionales.Text := '0';
      ttripulacion_extranjeros.Text := '0';
    end;

end;

 //SOAD -> Busqued de tripulacion por id.

procedure TfrmTripulacionPernoctas.cmbTurnosExit(Sender: TObject);
begin
  tdIdFecha.OnExit(sender);
end;

procedure TfrmTripulacionPernoctas.cmdAgregarClick(Sender: TObject);
begin
  try

    tripulacionDiaria.Open;
    tripulacionDiaria.Append;
    tripulacionDiaria.FieldValues['sContrato'] := global_contrato;
    tripulacionDiaria.FieldValues['dIdFecha'] := tdIdFecha.Date;
    tripulacionDiaria.FieldValues['sIdCategoria'] := tsIdCategoria.KeyValue;
    tripulacionDiaria.FieldValues['sNumeroOrden'] := tsIdFolio.KeyValue;
    tripulacionDiaria.FieldValues['dCantidad']    := txtCantidad.Text;
    tripulacionDiaria.FieldValues['sIdCuenta']    := tsIdTripulacion.KeyValue;
    tripulacionDiaria.FieldValues['iidDiario']    := 0;
    tripulacionDiaria.Post; 

    qry_tripulacion.Refresh;
    txtCantidad.Text := '0';
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Control de pernoctas', 'Al salvar registro', 0);
    end;
  end;

  if ZLookTripulacion.RecordCount > 0 then
  begin
    ZLookTripulacion.First;
    tsIdTripulacion.keyvalue := ZLookTripulacion.FieldValues['sIdCuenta'];
  end;
  tsIdTripulacion.SetFocus;

end;

procedure TfrmTripulacionPernoctas.cmdSalirClick(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
  cmdSalir.Visible := False;    
end;

procedure TfrmTripulacionPernoctas.Panel2Click(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
end;

procedure TfrmTripulacionPernoctas.PopupMenu1Popup(Sender: TObject);
var
    zPersonal : tzReadOnlyQuery;
begin
    zPersonal := tzReadOnlyQuery.Create(self);
    zPersonal.Connection := connection.zConnection;

    zPersonal.Active := False;
    zPersonal.SQL.Clear;
    zPersonal.SQL.Add('select sAnexo from anexos where sTipo = "PERSONAL"');
    zPersonal.Open;

    if zPersonal.RecordCount > 0 then
       ActualizaPersonal.Caption := 'Actualizar Partidas de Anexo '+ zPersonal.FieldValues['sAnexo'];

    zPersonal.Destroy;
end;

end.

