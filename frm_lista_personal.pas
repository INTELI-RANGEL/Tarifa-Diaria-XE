unit frm_lista_personal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms,
  Dialogs, StdCtrls, DBCtrls, ComCtrls, Grids, DBGrids, global, frm_connection, utilerias,
  DB, ADODB, Buttons, ExtCtrls, frxClass, frxDBSet, ZAbstractRODataset, ExcelXP, Excel2000,
  ZDataset, ZAbstractDataset, Controls, Menus, UnitExcepciones, udbgrid, UFunctionsGHH,
  DBDateTimePicker, UnitValidacion, rxToolEdit, OleServer, ComObj, StrUtils,
  Mask;

type
  TfrmLista_personal = class(TForm)
    grid_tripulacion: TDBGrid;
    ds_tripulacion: TDataSource;
    Panel1: TPanel;
    ttripulacion_nacionales: TEdit;
    DBTripulacion: TfrxDBDataset;
    DBTotalesxCategoria: TfrxDBDataset;
    frxTripulacion: TfrxReport;
    ds_categorias: TDataSource;
    categorias: TZReadOnlyQuery;
    tripulaciondiaria: TZQuery;
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
    ds_turnos: TDataSource;
    QryTurnos: TZQuery;
    qry_TripulacionsContrato: TStringField;
    qry_TripulacionsIdTurno: TStringField;
    qry_TripulaciondIdFecha: TDateField;
    qry_TripulacionsIdCategoria: TStringField;
    qry_TripulacionsIdTripulacion: TStringField;
    qry_TripulacionsOrden: TStringField;
    qry_TripulacionsNombre: TStringField;
    qry_TripulacionsIdPersonal: TStringField;
    qry_TripulacionsDescripcion: TStringField;
    qry_TripulacionsNumeroCabina: TStringField;
    qry_TripulacionsNacionalidad: TStringField;
    qry_TripulacioniNacionales: TIntegerField;
    qry_TripulacioniExtranjeros: TIntegerField;
    OpenXLS: TOpenDialog;
    ExcelApplication1: TExcelApplication;
    ExcelWorksheet1: TExcelWorksheet;
    ExcelWorkbook1: TExcelWorkbook;
    ds_ordenes: TDataSource;
    ordenes: TZQuery;
    Label8: TLabel;
    Edit1: TEdit;
    Label9: TLabel;
    zComPersonal: TZQuery;
    strngfld_TripulacionsIdCompania: TStringField;
    pnlFiltro: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label7: TLabel;
    tsIdCategoria: TDBLookupComboBox;
    tdIdFecha: TDBDateTimePicker;
    cmbTurnos: TDBLookupComboBox;
    tsOrdenes: TDBLookupComboBox;
    pnlBotonera: TPanel;
    btnNuevo: TBitBtn;
    btnDelete: TBitBtn;
    btnPrinter: TBitBtn;
    cmdImportar: TBitBtn;
    Label6: TLabel;
    tsArchivo: TEdit;
    btnExportar: TBitBtn;
    strngfld_TripulacionCompania: TStringField;
    strngfld_TripulacionCuenta: TStringField;
    zCuentas: TZQuery;
    strngfld_TripulacionsIdCuenta: TStringField;
    strngfld_TripulacionsRfc: TStringField;
    cmdAgregar: TButton;
    pnlDatos: TPanel;
    dbedtFicha: TDBEdit;
    dbedtNombre: TDBEdit;
    dbedtOrden: TDBEdit;
    DbLCbbCompania: TDBLookupComboBox;
    dbedtCabina: TDBEdit;
    DbLCbbCuenta: TDBLookupComboBox;
    btnGuardar: TButton;
    btnCancelar: TButton;
    lbl1: TLabel;
    lbl2: TLabel;
    lbl3: TLabel;
    lbl5: TLabel;
    lbl6: TLabel;
    lbl7: TLabel;
    lbl8: TLabel;
    zTripulacionDiariaUpt: TZQuery;
    dsTripulacionDiariaUpt: TDataSource;
    zLookup_Compania: TZQuery;
    dsLookup_Compania: TDataSource;
    zLookup_Cuentas: TZQuery;
    dsLookup_Cuentas: TDataSource;
    dbedtsIdPersonal: TDBEdit;
    lbl4: TLabel;
    dbedtRfc: TDBEdit;
    lblRfc: TLabel;
    DbLCbbCategoria: TDBLookupComboBox;
    zLookUp_Categorias: TZQuery;
    zlookup_Tripulacion: TZQuery;
    dsLookup_Tripulacion: TDataSource;
    btnEditar: TBitBtn;
    pnlEditar: TPanel;
    Label15: TLabel;
    Label16: TLabel;
    Label18: TLabel;
    dbCompUpt: TDBLookupComboBox;
    dbCuentaUpt: TDBLookupComboBox;
    btnActualizar: TButton;
    Button2: TButton;
    dbRfcUpt: TDBEdit;
    zq_TripulacionUpt: TZQuery;
    ds_TripulacionUpt: TDataSource;
    btnDeleteAll: TBitBtn;
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdCategoriaKeyPress(Sender: TObject; var Key: Char);
    procedure ttripulacion_extranjerosKeyPress(Sender: TObject;
      var Key: Char);
    procedure tdIdFechaExit(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
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
    procedure grid_tripulacionMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_tripulacionMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_tripulacionTitleClick(Column: TColumn);
    procedure qry_TripulacionBeforePost(DataSet: TDataSet);
    procedure txtCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure txtCantidadChange(Sender: TObject);
    procedure cmbTurnosExit(Sender: TObject);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure cmdImportarClick(Sender: TObject);
    procedure tsOrdenesExit(Sender: TObject);
    procedure tsOrdenesEnter(Sender: TObject);
    procedure tsOrdenesKeyPress(Sender: TObject; var Key: Char);
    procedure btnExportarClick(Sender: TObject);
    procedure tdIdFechaKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnGuardarClick(Sender: TObject);
    procedure tsOrdenesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tsIdCategoriaKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cmbTurnosKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnEditarClick(Sender: TObject);
    procedure btnActualizarClick(Sender: TObject);
    procedure btnDeleteAllClick(Sender: TObject);



  private
    forma: TForm;
    procedure onShowCustomForm(Sender: TObject);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLista_personal: TfrmLista_personal;
  sFirma_PEP, sPuesto_PEP, sFirma_Contratista, sPuesto_Contratista: string;
  fechaAntes: tDate;
  utgrid: ticdbgrid;

  Excel, Libro, Hoja: Variant;
  columnas: array[1..260] of string;
  MensajePartidas, sDatoNivel: string;
  lContratoActual: boolean;
  flcid, Fila: Integer;
  Resp: Integer;

implementation

{$R *.dfm}

procedure TfrmLista_personal.tdIdFechaKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 then
    tdIdFechaExit(nil);
end;

procedure TfrmLista_personal.tdIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
  {if key = #13 then
    tsIdCategoria.SetFocus  }
end;

procedure TfrmLista_personal.tsIdCategoriaKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if key = 13 then
    if tsIdCategoria.KeyValue <> null then
      tsIdCategoriaExit(nil);
end;

procedure TfrmLista_personal.tsIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
 { if key = #13 then
    grid_tripulacion.SetFocus   }
end;

procedure TfrmLista_personal.tsIdTripulacionEnter(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_entrada;
end;

procedure TfrmLista_personal.tsIdTripulacionExit(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_salida;
//  if tsidtripulacion.KeyValue <> null then
//    lblTripulacion.Caption := 'Tripulacion ' + tsIdTripulacion.KeyValue;
end;

procedure TfrmLista_personal.tsOrdenesEnter(Sender: TObject);
begin
  tsOrdenes.Color := global_color_entrada;
end;

procedure TfrmLista_personal.tsOrdenesExit(Sender: TObject);
begin
  tsOrdenes.Color := global_Color_salida;
  with qry_tripulacion do
  begin
    Active := False;
    SQL.Clear;
    SQL.Add('SELECT a.*, c.sDescripcion as compania, cu.sDescripcion as Cuenta FROM tripulaciondiaria_listado as a ' +
      ' left join compersonal as c on(c.sIdCompania = a.sIdCompania) ' +
      ' Left join cuentas as cu on (cu.sIdCuenta = a.sIdCuenta) ' +
      ' where sContrato = :contrato and dIdfecha = :fecha ' +
      ' And sIdCategoria= :Categoria and sOrden =:Ordenes and sIdTurno =:Turno ');
    params.ParamByName('contrato').DataType  := ftString;
    params.ParamByName('contrato').Value     := global_contrato_barco;
    params.ParamByName('fecha').DataType     := ftDate;
    params.ParamByName('fecha').Value        := tdIdFecha.Date;
    params.ParamByName('categoria').DataType := ftString;
    params.ParamByName('categoria').Value    := tsIdCategoria.KeyValue;
    params.ParamByName('ordenes').DataType   := ftString;
    params.ParamByName('ordenes').Value      := tsOrdenes.KeyValue;
    params.ParamByName('turno').DataType     := ftString;
    params.ParamByName('turno').Value        := QryTurnos.FieldValues['sIdTurno'];
    Open;
    if qry_Tripulacion.RecordCount = 0 then
       ttripulacion_nacionales.Text := '0';
  end;

  //Consultamos las categorias de personal que se deben mostrar...
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select sIdPersonal from personal where sContrato =:contrato and sIdTipoPersonal not like "%EXT%" ');
  connection.QryBusca.ParamByName('contrato').AsString := tsOrdenes.KeyValue;
  connection.QryBusca.Open;

  grid_tripulacion.Columns[3].PickList.clear;
  grid_tripulacion.Columns[3].PickList.add('');
  while not connection.QryBusca.Eof do
  begin
      grid_tripulacion.Columns[3].PickList.add(connection.QryBusca.FieldValues['sIdPersonal']);
      connection.QryBusca.Next;
  end;

  //Ahora sumamos las cantidades de qie existen por categorias..
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('SELECT sum(iNacionales) as cantidad FROM tripulaciondiaria_listado where sContrato = :contrato and dIdfecha = :fecha '+
      'And sIdCategoria= :Categoria and sOrden =:Ordenes and sIdTurno =:Turno group by sContrato ');
  connection.QryBusca.params.ParamByName('contrato').DataType  := ftString;
  connection.QryBusca.params.ParamByName('contrato').Value     := global_contrato_barco;
  connection.QryBusca.params.ParamByName('fecha').DataType     := ftDate;
  connection.QryBusca.params.ParamByName('fecha').Value        := tdIdFecha.Date;
  connection.QryBusca.params.ParamByName('categoria').DataType := ftString;
  connection.QryBusca.params.ParamByName('categoria').Value    := tsIdCategoria.KeyValue;
  connection.QryBusca.params.ParamByName('ordenes').DataType   := ftString;
  connection.QryBusca.params.ParamByName('ordenes').Value      := tsOrdenes.KeyValue;
  connection.QryBusca.params.ParamByName('turno').DataType     := ftString;
  connection.QryBusca.params.ParamByName('turno').Value        := QryTurnos.FieldValues['sIdTurno'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
     ttripulacion_nacionales.Text := connection.QryBusca.FieldValues['cantidad']
  else
     ttripulacion_nacionales.Text := '0';
end;

procedure TfrmLista_personal.tsOrdenesKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 13 then
    if tsOrdenes.KeyValue <> null then
      tsOrdenesExit(nil);
end;

procedure TfrmLista_personal.tsOrdenesKeyPress(Sender: TObject; var Key: Char);
begin
    if key =#10 then
       tsOrdenes.KeyValue := Null;
end;

procedure TfrmLista_personal.txtCantidadChange(Sender: TObject);
begin
  tEditChangef(txtCantidad, 'Cantidad');
end;

procedure TfrmLista_personal.txtCantidadEnter(Sender: TObject);
begin
  txtcantidad.Color := global_color_entrada;
end;

procedure TfrmLista_personal.txtCantidadExit(Sender: TObject);
begin
  txtcantidad.Color := global_color_salida
end;

procedure TfrmLista_personal.txtCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTedit(txtCantidad, key) then
    key := #0;
end;

procedure TfrmLista_personal.ttripulacion_extranjerosKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdidFecha.SetFocus
end;

//Inicio LA forma

procedure TfrmLista_personal.FormShow(Sender: TObject);
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

  QryTurnos.Active := False;
  QryTurnos.ParamByName('Contrato').AsString := global_contrato_barco;
  QryTurnos.Open;

  cmbTurnos.KeyValue := 'A';

  Tripulacion.Active := False;
  Tripulacion.SQL.Clear;
  Tripulacion.SQL.Add('select * from tripulacion where sContrato = :Contrato and dFechaVigencia =:Fecha order by sIdTripulacion');
  Tripulacion.params.ParamByName('Contrato').DataType := ftString;
  Tripulacion.params.ParamByName('Contrato').Value := Global_Contrato_barco;
  Tripulacion.params.ParamByName('fecha').DataType := ftDate;
  Tripulacion.params.ParamByName('fecha').Value := fechaAntes;
  Tripulacion.Open;
end;

procedure TfrmLista_personal.frxTripulacionGetValue(const VarName: string;
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

procedure TfrmLista_personal.grid_tripulacionMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
//if grid_tripulacion.datasource.DataSet.IsEmpty=false  then
//if grid_tripulacion.DataSource.DataSet.RecordCount>0 then
//UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmLista_personal.grid_tripulacionMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
//if grid_tripulacion.datasource.DataSet.IsEmpty=false  then
//if grid_tripulacion.DataSource.DataSet.RecordCount>0 then
//UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmLista_personal.grid_tripulacionTitleClick(Column: TColumn);
begin
    UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmLista_personal.onShowCustomForm(Sender: TObject);
begin
  if dbedtFicha.CanFocus then
    dbedtFicha.SetFocus;
end;

procedure TfrmLista_personal.tdIdFechaExit(Sender: TObject);
begin
  tdidfecha.Color := global_color_salida;
  tsOrdenes.KeyValue := Null;
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
  Tripulacion.params.ParamByName('Contrato').Value    := Global_Contrato_barco;
  Tripulacion.params.ParamByName('fecha').DataType    := ftDate;
  Tripulacion.params.ParamByName('fecha').Value       := fechaAntes;
  Tripulacion.Open;

  Categorias.Active := False;
  Categorias.ParamByName('fecha').AsDate := fechaAntes;
  Categorias.Open;

  with qry_tripulacion do
  begin
      Active := False;
      SQL.Clear;
      SQL.Add('SELECT a.*, c.sDescripcion as Compania, cu.sDescripcion as Cuenta FROM tripulaciondiaria_listado as a ' +
      ' Left join compersonal as c on (c.sIdCompania = a.sIdCompania) ' +
      ' Left join cuentas as cu on (cu.sIdCuenta = a.sIdCuenta) ' +
      ' where sContrato = :contrato and dIdfecha = :fecha and sIdTurno =:Turno ');
      params.ParamByName('contrato').DataType := ftString;
      params.ParamByName('contrato').Value := global_contrato_barco;
      params.ParamByName('fecha').DataType := ftDate;
      params.ParamByName('fecha').Value := tdIdFecha.Date;
      params.ParamByName('turno').DataType := ftString;
      params.ParamByName('turno').Value := QryTurnos.FieldValues['sIdTurno'];
      Open;

      if qry_Tripulacion.RecordCount = 0 then
      begin
          ttripulacion_nacionales.Text := '0';
          qry_Tripulacion.Close;
      end;
  end;

  //Ahora sumamos las cantidades de personal que existen por categorias..
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('SELECT sum(iNacionales) as cantidad FROM tripulaciondiaria_listado where sContrato = :contrato and dIdfecha = :fecha '+
      ' and sIdTurno =:Turno group by sContrato ');
  connection.QryBusca.params.ParamByName('contrato').DataType  := ftString;
  connection.QryBusca.params.ParamByName('contrato').Value     := global_contrato_barco;
  connection.QryBusca.params.ParamByName('fecha').DataType     := ftDate;
  connection.QryBusca.params.ParamByName('fecha').Value        := tdIdFecha.Date;
  connection.QryBusca.params.ParamByName('turno').DataType     := ftString;
  connection.QryBusca.params.ParamByName('turno').Value        := QryTurnos.FieldValues['sIdTurno'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
     ttripulacion_nacionales.Text := connection.QryBusca.FieldValues['cantidad']
  else
     ttripulacion_nacionales.Text := '0';

  //Consultamos las ordenes de trabajo vigentes a la fecha.
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select * from reportediario where dIdFecha =:Fecha and sContrato <> :contrato group by sContrato ');
  connection.QryBusca.ParamByName('fecha').AsDate      := tdIdfecha.Date;
  connection.QryBusca.ParamByName('contrato').AsString := connection.contrato.FieldValues['sCodigo'];
  connection.QryBusca.Open;

  grid_tripulacion.Columns[2].PickList.clear;
  grid_tripulacion.Columns[2].PickList.add('');
  while not connection.QryBusca.Eof do
  begin
      grid_tripulacion.Columns[2].PickList.add(connection.QryBusca.FieldValues['sContrato']);
      connection.QryBusca.Next;
  end;

  //Consultamos las categorias de personal que se deben mostrar...
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select sIdPersonal from personal where sContrato =:contrato and sIdTipoPersonal not like "%EXT%" ');
  connection.QryBusca.ParamByName('contrato').AsString := connection.contrato.FieldValues['sCodigo'];
  connection.QryBusca.Open;

  grid_tripulacion.Columns[3].PickList.clear;
  grid_tripulacion.Columns[3].PickList.add('');
  while not connection.QryBusca.Eof do
  begin
      grid_tripulacion.Columns[3].PickList.add(connection.QryBusca.FieldValues['sIdPersonal']);
      connection.QryBusca.Next;
  end;

  ordenes.Active := False;
  ordenes.SQL.Clear;
  ordenes.SQL.Add('select * from reportediario where dIdFecha =:Fecha group by sOrden ');
  ordenes.ParamByName('fecha').AsDate      := tdIdfecha.Date;
  ordenes.Open;

end;

//Consultar la Categoria
procedure TfrmLista_personal.btnEditarClick(Sender: TObject);
var
  Cursor: TCursor;
begin
  if tsIdCategoria.KeyValue = null then
  begin
    ShowMessage('Favor de Seleccionar una Categoria');
    exit;
  end;
  try
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    try
      if Assigned(FindComponent('FrmEdit')) then
        TForm(FindComponent('FrmEdit')).Destroy;

      forma := TForm.Create(Self);
      forma.Name := 'FrmEdit';
      forma.OnShow := onShowCustomForm;
      forma.Caption := 'Editar registro.';
      forma.Height := 170;
      forma.Width := 370;
      forma.BorderStyle := bsDialog;
      pnlEditar.Parent := forma;
      pnlEditar.Align := alClient;

      with zLookup_Compania do
      begin
        active := False;
        Open;
      end;

      with zLookup_Cuentas do
      begin
        active := False;
        Open;
      end;

      with zq_TripulacionUpt do
      begin
        Active := False;
        ParamByName('Contrato').AsString:=qry_Tripulacion.FieldByName('sContrato').AsString;
        ParamByName('Fecha').AsDateTime:=qry_Tripulacion.FieldByName('dIdFecha').AsDateTime;
        ParamByName('Turno').AsString:=qry_Tripulacion.FieldByName('sIdTurno').AsString;
        ParamByName('Tripulacion').AsString:=qry_Tripulacion.FieldByName('sIdTripulacion').AsString;
        Open;
        Edit;
      end;

      dbCompUpt.KeyValue:=qry_Tripulacion.FieldByName('sIdCompania').AsString;
      dbCuentaUpt.KeyValue:=qry_Tripulacion.FieldByName('sIdCuenta').AsString;

      pnlEditar.Visible := True;
      if forma.ShowModal <> mrCancel then
        qry_Tripulacion.Refresh;
    finally
      pnlEditar.Visible := False;
      pnlEditar.Align := alNone;
      pnlEditar.Parent := Self;
      forma.Destroy;
      Screen.Cursor := Cursor;
    end;
  except
    on E: Exception do
    begin
        MessageDlg(e.Message, mtInformation, [mbOk], 0)
    end;
  end;
end;

procedure TfrmLista_personal.btnActualizarClick(Sender: TObject);
var
  cursor: TCursor;
begin
  try
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    try

      if dbCompUpt.KeyValue = Null  then
      begin
        MessageDlg('Debe seleccionar una [Compañía] antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if dbCompUpt.CanFocus then
          dbCompUpt.SetFocus;
        Exit;
      end;

      if dbCuentaUpt.KeyValue = Null  then
      begin
        MessageDlg('Debe seleccionar una [Cuenta] antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if dbCuentaUpt.CanFocus then
          dbCuentaUpt.SetFocus;
        Exit;
      end;

      if zq_TripulacionUpt.State in [dsInsert, dsEdit] then
      begin
        zq_TripulacionUpt.Post;
        zq_TripulacionUpt.ApplyUpdates;
        qry_Tripulacion.Refresh;
      end;
      Forma.Close;
    finally
      Screen.Cursor := Cursor;
    end;
  Except
    on e: Exception do
      MessageDlg('Avisar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TfrmLista_personal.btnCancelClick(Sender: TObject);
begin
  close;
end;

procedure TfrmLista_personal.qry_tripulacionAfterInsert(
  DataSet: TDataSet);
begin
  qry_tripulacion.Cancel
end;

procedure TfrmLista_personal.qry_TripulacionBeforePost(DataSet: TDataSet);
begin
  if qry_Tripulacion.FieldValues['inacionales'] < 0 then
    qry_Tripulacion.cancel;
  if qry_Tripulacion.FieldValues['iExtranjeros'] < 0 then
    qry_Tripulacion.cancel;

  if qry_tripulacionsOrden.Text = '' then
     qry_tripulacionsOrden.text := ''
  else
  begin
      //Consultamos los tipos de personal que se deben mostrar en personal y equipo..
      connection.QryBusca.Active := FalsE;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select * from reportediario where dIdFecha =:Fecha and sContrato <> :contrato and sContrato like :orden group by sContrato ');
      connection.QryBusca.ParamByName('fecha').AsDate      := tdIdfecha.Date;
      connection.QryBusca.ParamByName('contrato').AsString := connection.contrato.FieldValues['sCodigo'];
      connection.QryBusca.ParamByName('orden').AsString    := '%'+qry_tripulacionsOrden.Text + '%';
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
         qry_tripulacionsOrden.text := connection.QryBusca.FieldValues['sContrato'];

      if connection.QryBusca.RecordCount = 0 then
      begin
         messageDLg('No existen esta orden en los reportes diarios!', mtInformation, [mbOk], 0);
         Abort;
         exit;
      end;
  end;

  if qry_tripulacionsIdPersonal.Text = '' then
     qry_tripulacionsIdPersonal.text := ''
  else
  begin
      //Consultamos los tipos de personal que se deben mostrar en personal y equipo..
      connection.QryBusca.Active := FalsE;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select sIdPersonal from personal where sContrato =:contrato and sIdPersonal like :personal group by sIdPersonal ');
      connection.QryBusca.ParamByName('contrato').AsString := qry_tripulacion.FieldValues['sContrato'];
      connection.QryBusca.ParamByName('personal').AsString    := '%'+qry_tripulacionsOrden.Text + '%';
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
         qry_tripulacionsIdPersonal.text := connection.QryBusca.FieldValues['sIdPersonal'];

      if connection.QryBusca.RecordCount = 0 then
      begin
         messageDLg('No existe esta partida de Personal en la Orden de Trabajo!', mtInformation, [mbOk], 0);
         Abort;
         exit;
      end;
  end;   

end;

procedure TfrmLista_personal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  utgrid.Destroy;
  action := cafree;
end;

procedure TfrmLista_personal.tdIdFechaEnter(Sender: TObject);
begin
  tdIdFecha.Color := global_Color_entrada
end;


procedure TfrmLista_personal.qry_tripulacionCalcFields(
  DataSet: TDataSet);
begin

  Connection.QryBusca.SQL.Clear;
  Connection.QryBusca.SQL.Add('Select sDescripcion from tripulacion Where sContrato =:Contrato ' +
    'And sIdCategoria = :Categoria And sIdTripulacion = :CatTripulacion and dFechaVigencia =:Fecha ');
  Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.QryBusca.Params.ParamByName('Contrato').Value := Global_Contrato_barco;
  Connection.QryBusca.Params.ParamByName('Categoria').DataType := ftString;
  Connection.QryBusca.Params.ParamByName('Categoria').Value := qry_Tripulacion.FieldValues['sIdCategoria'];
  Connection.QryBusca.Params.ParamByName('CatTripulacion').DataType := ftString;
  Connection.QryBusca.Params.ParamByName('CatTripulacion').Value := qry_Tripulacion.FieldValues['sIdTripulacion'];
  Connection.QryBusca.Params.ParamByName('Fecha').DataType := ftDate;
  Connection.QryBusca.Params.ParamByName('Fecha').Value := fechaAntes;
  Connection.QryBusca.Open;
  if Connection.QryBusca.RecordCount > 0 then
    qry_TripulacionsDescripcion.Text := Connection.QryBusca.FieldValues['sDescripcion'];


end;



procedure TfrmLista_personal.ripulacinDiariaDiaAnterior1Click(
  Sender: TObject);
var
  dFechaAnterior : tDate;
  lNuevaVigencia : boolean;

begin
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('Select * from tripulaciondiaria_listado Where sContrato =:Contrato And dIdFecha =:Fecha');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value    := Global_Contrato_barco;
  Connection.qryBusca.Params.ParamByName('Fecha').Datatype    := ftDate;
  Connection.qryBusca.Params.ParamByName('Fecha').Value       := tdIdFecha.DateTime;
  Connection.qryBusca.Open;
  if Connection.qryBusca.RecordCount > 0 then
    if MessageDlg('Desea Eliminar todo la Tripulacion Asignado?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('Delete From tripulaciondiaria_listado where sContrato = :contrato and dIdFecha = :fecha and sIdTurno =:Turno ');
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato_barco;
      connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
      connection.zCommand.Params.ParamByName('Turno').Value := QryTurnos.FieldValues['sIdTurno'];
      connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
      connection.zCommand.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
      connection.zCommand.ExecSQL;
    end;

  if MessageDlg('Desea adicionar todo la Tripulación Existente en el reporte anterior?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
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
              Connection.Auxiliar.SQL.Add('Select * from tripulaciondiaria_listado Where sContrato =:Contrato And dIdFecha =:Fecha');
              Connection.Auxiliar.Params.ParamByName('Contrato').DataType := ftString;
              Connection.Auxiliar.Params.ParamByName('Contrato').Value    := Global_Contrato_barco;
              Connection.Auxiliar.Params.ParamByName('Fecha').Datatype    := ftDate;
              Connection.Auxiliar.Params.ParamByName('Fecha').Value       := dFechaAnterior;
              Connection.Auxiliar.Open;
          end;
      end;

      if Connection.Auxiliar.RecordCount > 0 then
      begin
          Connection.zcommand.SQL.Clear;
          Connection.zCommand.SQL.Add('INSERT INTO tripulaciondiaria_listado ( sContrato , sIdTurno, dIdFecha, sIdCategoria, sIdTripulacion, sOrden, sNombre, sIdPersonal, sDescripcion, sCompania, sGenero, '+
                                      'dFechaInicio, dFechaFinal, sFolioLibreta, dFechaVigencia, sNumeroCabina, sNacionalidad, iNacionales, iExtranjeros) ' +
                                      'VALUES (:contrato, :turno, :fecha , :categoria, :tripulacion, :orden, :nombre, :idpersonal, :descripcion, :compañia, :genero, '+
                                      ':Inicio, :Final, :libreta, :vigencia, :cabina, :nacionalidad, :nacionales, :extranjeros)');
          while not Connection.Auxiliar.Eof do
          begin
            with connection do
            begin
                try
                  zcommand.params.ParamByName('contrato').DataType    := ftString;
                  zcommand.params.ParamByName('contrato').value       := Global_Contrato_barco;
                  zcommand.params.ParamByName('Turno').DataType       := ftString;
                  zcommand.params.ParamByName('Turno').value          := QryTurnos.FieldValues['sIdTurno'];
                  zcommand.params.ParamByName('Turno').value          := QryTurnos.FieldValues['sIdTurno'];
                  zcommand.params.ParamByName('fecha').DataType       := ftDate;
                  zcommand.params.ParamByName('fecha').value          := tdIdFecha.Date;
                  zcommand.params.ParamByName('categoria').DataType   := ftString;
                  zcommand.params.ParamByName('categoria').value      := Auxiliar.FieldValues['sIdCategoria'];
                  zcommand.params.ParamByName('tripulacion').DataType := ftString;
                  zcommand.params.ParamByName('tripulacion').value    := Auxiliar.FieldValues['sIdTripulacion'];
                  zCommand.Params.ParamByName('orden').DataType       := ftString;
                  zCommand.Params.ParamByName('orden').value          := Auxiliar.FieldValues['sOrden'];
                  zCommand.Params.ParamByName('Nombre').DataType      := ftString;
                  zCommand.Params.ParamByName('Nombre').value         := Auxiliar.FieldValues['sNombre'];
                  zCommand.Params.ParamByName('IdPersonal').DataType  := ftString;
                  zCommand.Params.ParamByName('IdPersonal').value     := Auxiliar.FieldValues['sIdPersonal'];
                  zCommand.Params.ParamByName('Descripcion').DataType := ftString;
                  zCommand.Params.ParamByName('Descripcion').value    := Auxiliar.FieldValues['sDescripcion'];
                  zCommand.Params.ParamByName('Compañia').DataType := ftString;
                  zCommand.Params.ParamByName('Compañia').value    := Auxiliar.FieldValues['sCompania'];
                  zCommand.Params.ParamByName('Genero').DataType   := ftString;
                  zCommand.Params.ParamByName('Genero').value      := Auxiliar.FieldValues['sGenero'];
                  zCommand.Params.ParamByName('inicio').DataType   := ftDate;
                  zCommand.Params.ParamByName('inicio').value      := Auxiliar.FieldValues['dFechaInicio'];
                  zCommand.Params.ParamByName('final').DataType    := ftDate;
                  zCommand.Params.ParamByName('final').value       := Auxiliar.FieldValues['dFechaFinal'];
                  zCommand.Params.ParamByName('libreta').DataType  := ftString;
                  zCommand.Params.ParamByName('libreta').value     := Auxiliar.FieldValues['sFolioLibreta'];
                  zCommand.Params.ParamByName('vigencia').DataType := ftDate;
                  if Auxiliar.FieldValues['dFechaVigencia'] = Null then
                     zCommand.Params.ParamByName('vigencia').value    := tdIdFecha.Date
                  else
                     zCommand.Params.ParamByName('vigencia').value    := Auxiliar.FieldValues['dFechaVigencia'];
                  zCommand.Params.ParamByName('cabina').DataType   := ftString;
                  zCommand.Params.ParamByName('cabina').value      := Auxiliar.FieldValues['sNumeroCabina'];
                  zCommand.Params.ParamByName('nacionalidad').DataType := ftString;
                  zCommand.Params.ParamByName('nacionalidad').value    := Auxiliar.FieldValues['sNacionalidad'];
                  zcommand.params.ParamByName('Nacionales').DataType  := ftInteger;
                  zcommand.params.ParamByName('Nacionales').value     := Auxiliar.FieldValues['iNacionales'];
                  zcommand.params.ParamByName('Extranjeros').DataType := ftInteger;
                  zcommand.params.ParamByName('Extranjeros').value    := Auxiliar.FieldValues['iExtranjeros'];
                  zcommand.ExecSQL;
                  Auxiliar.Next;
                except
                end;
            end;
          end;
      end
      else
         messageDLG('No se encontró personal el día anterior', mtInformation, [mbOk], 0);
  end;
  tripulacion.Refresh;
  with qry_tripulacion do
  begin
      Active := False;
      SQL.Clear;
      SQL.Add('SELECT * FROM tripulaciondiaria_listado where sContrato = :contrato and dIdfecha = :fecha and sIdTurno =:Turno ');
      params.ParamByName('contrato').DataType := ftString;
      params.ParamByName('contrato').Value := global_contrato_barco;
      params.ParamByName('fecha').DataType := ftDate;
      params.ParamByName('fecha').Value := tdIdFecha.Date;
      params.ParamByName('turno').DataType := ftString;
      params.ParamByName('turno').Value := QryTurnos.FieldValues['sIdTurno'];
      Open;
  end;
end;

procedure TfrmLista_personal.btnDeleteAllClick(Sender: TObject);
begin
  try
    if (not qry_Tripulacion.active) or (qry_Tripulacion.RecordCount < 1) then
      exit;

    if (MessageDlg('¿Está seguro que desea eliminar todos los datos?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
    begin
      Connection.zCommand.Active := False;
      Connection.zCommand.SQL.Clear;
      Connection.zCommand.SQL.Add('Delete from tripulaciondiaria_listado where sContrato = :contrato and sIdCategoria = :Categoria ' +
        'And dIdFecha =:Fecha And sIdTurno =:Turno ');
      Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato_barco;
      Connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Turno').Value := QryTurnos.FieldValues['sIdTurno'];
      Connection.zCommand.Params.ParamByName('Categoria').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Categoria').Value := qry_Tripulacion.FieldValues['sIdCategoria'];
      Connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
      Connection.zCommand.Params.ParamByName('Fecha').Value := tdIdFecha.date;
      Connection.zCommand.ExecSQL();
      qry_Tripulacion.Refresh;
    end;
  except
    on e: Exception do
      MessageDlg('Ha ocurrido el siguiente error, informar al administrador del sistema: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TfrmLista_personal.btnDeleteClick(Sender: TObject);
begin
  try
    if (not qry_Tripulacion.active) or (qry_Tripulacion.RecordCount < 1) then
      exit;

    if (MessageDlg('¿Está seguro que desea eliminar [' +  qry_Tripulacion.FieldByName('sNombre').AsString + ']?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
    begin
      Connection.zCommand.Active := False;
      Connection.zCommand.SQL.Clear;
      Connection.zCommand.SQL.Add('Delete from tripulaciondiaria_listado where sContrato = :contrato and sIdCategoria = :Categoria ' +
        'And dIdFecha =:Fecha And sIdTripulacion =:Tripulacion and sIdTurno =:Turno ');
      Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato_barco;
      Connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Turno').Value := QryTurnos.FieldValues['sIdTurno'];
      Connection.zCommand.Params.ParamByName('Categoria').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Categoria').Value := qry_Tripulacion.FieldValues['sIdCategoria'];
      Connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
      Connection.zCommand.Params.ParamByName('Fecha').Value := tdIdFecha.date;
      Connection.zCommand.Params.ParamByName('Tripulacion').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Tripulacion').Value := qry_Tripulacion.FieldValues['sIdTripulacion'];
      Connection.zCommand.ExecSQL();
      qry_Tripulacion.Refresh;
    end;
  except
    on e: Exception do
      MessageDlg('Ha ocurrido el siguiente error, informar al administrador del sistema: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TfrmLista_personal.btnExportarClick(Sender: TObject);
const
  NombreCols: Array[1..13] of String = ('No. Ficha', 'Nombre', 'Id Categoría', 'Categoria', 'OT', 'Id Personal', 'Descripción', 'Cabina', 'Compania', 'sIdCompania', 'Cuenta', 'sIdCuenta', 'RFC');
  NombreFields: Array[1..13] of string = ('sIdTripulacion', 'sNombre', 'sIdCategoria', '', 'sOrden', 'sIdPersonal', 'sDescripcion', 'sNumeroCabina', 'Compania', 'sIdCompania', 'Cuenta', 'sIdCuenta', 'sRfc');
var
  Excel, workSheet, WorkBook: Variant;
  Cursor: TCursor;
  i, j: Integer;
  ExcepcionesFields: TStringList;

Procedure createComboExcel(Var Hoja: Variant; NombreHoja: String; PosCol: String; ListaDatos:string);
begin
  Hoja.Sheets[NombreHoja].Select;
  Hoja.Range[PosCol].Select;
  hoja.Selection.Validation.Delete;
  hoja.Selection.Validation.add(xlValidateList,xlValidAlertStop,xlBetween,ListaDatos);
  hoja.Selection.Validation.IgnoreBlank := True;
  hoja.Selection.Validation.InCellDropdown := True;
  hoja.Selection.Validation.InputTitle := '';
  hoja.Selection.Validation.ErrorTitle := '';
  hoja.Selection.Validation.ErrorMessage := '';
  hoja.Selection.Validation.ShowInput := True;
  hoja.Selection.Validation.ShowError := True;
end;
Procedure AsignaFormulas(var Hoja: Variant; NombreHoja: string; Celda: String; Formula: String; Rango_AutoFill: string;
                         sLocked: Boolean; sFormulaOculta: Boolean; OcultarColumna: String);
begin
  Hoja.Sheets[NombreHoja].Select;
  Hoja.Range[Celda].Select;
  Hoja.Selection.FormulaR1C1 := Formula;
  if Length(Trim(Rango_AutoFill)) > 0 then
    Hoja.Selection.Autofill(Hoja.range[Rango_AutoFill], xlFillDefault);
  if Length(Trim(OcultarColumna)) > 0 then
  begin
    Hoja.Columns[OcultarColumna].select;
    Hoja.Selection.Locked := sLocked;
    Hoja.Selection.FormulaHidden := sFormulaOculta;
    Hoja.Selection.EntireColumn.hidden := sFormulaOculta;
  end;
end;

begin
  try
    i := 0;
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    if (Not qry_Tripulacion.Active) and (qry_Tripulacion.RecordCount = 0) then
    begin
      MessageDlg('No existen datos a exportar, Filtrar por fecha.',mtWarning,[mbOK], 0);
      Exit;
    end;
    try
      //Columnas descartadas
      ExcepcionesFields := TStringList.Create;
      ExcepcionesFields.Add('9');
      ExcepcionesFields.Add('11');

      with zCompersonal do
      begin
        Active := False;
        SQL.Clear;
        SQL.Text := 'Select * from compersonal';
        Open;
      end;

      with zCuentas do
      begin
        Active := False;
        SQL.Clear;
        SQL.Text := 'Select sIdCuenta, sDescripcion from cuentas';
        Open;
      end;

      Excel := CreateOleObject('Excel.Application');
      Excel.Visible := True;
      Excel.WorkBooks.Add(xlWBATWorksheet);
      Excel.ActiveSheet.Name := leftStr('Compersonal', 31);

      //Columnas Excel LookUpComboExcel Companias
      Excel.Sheets['compersonal'].Select;
      zCompersonal.First;
      while Not zCompersonal.Eof do
      begin
        Excel.Cells[zCompersonal.RecNo,2] := zCompersonal.FieldByName('sIdCompania').AsString;
        Excel.Cells[zCompersonal.RecNo,1] := zCompersonal.FieldByName('sDescripcion').AsString;
        zCompersonal.Next;
      end;

       //Columnas Excel LookUpComboExcel Companias
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Cuentas',31);
      Excel.Sheets['Cuentas'].Select;
      zCuentas.First;
      while Not zCuentas.Eof do
      begin
        Excel.Cells[zCuentas.RecNo,2] := zCuentas.FieldByName('sIdCuenta').AsString;
        Excel.Cells[zCuentas.RecNo,1] := zCuentas.FieldByName('sDescripcion').AsString;
        zCuentas.Next;
      end;

      //Llenar el La plantilla de Excel
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Plantilla',31);
      Excel.Sheets['Plantilla'].Select;

      for i  := 1 to 13 do
        Excel.Cells[1, i] := nombreCols[i];

      if (qry_Tripulacion.Active) and (qry_Tripulacion.RecordCount > 0) then
      begin
        qry_Tripulacion.First;
        while not qry_Tripulacion.Eof do
        begin
          for i := 1 to 13 do
            for j := 0 to ExcepcionesFields.Count - 1 do
              if ExcepcionesFields[j] <> IntToStr(i) then
                if Length(trim(NombreFields[i])) > 0 then
                  Excel.Cells[qry_Tripulacion.RecNo + 1, i] := qry_Tripulacion.FieldByName(NombreFields[i]).AsString
                else
                  Excel.Cells[qry_Tripulacion.RecNo + 1, i] := '';
          createComboExcel(Excel, 'Plantilla', 'I' + IntToStr(qry_Tripulacion.RecNo + 1), '=Compersonal!$A:$A');
          Excel.Cells[qry_Tripulacion.RecNo + 1, 9] := qry_Tripulacion.FieldByName(NombreFields[9]).AsString;
          createComboExcel(Excel, 'Plantilla', 'K' + IntToStr(qry_Tripulacion.RecNo + 1), '=Cuentas!$A:$A');
          Excel.Cells[qry_Tripulacion.RecNo + 1, 11] := qry_Tripulacion.FieldByName(NombreFields[11]).AsString;
          qry_Tripulacion.Next;
        end;
      end;
      AsignaFormulas(Excel, 'Plantilla', 'J2', '=VLOOKUP(RC[-1], Compersonal!C[-9]:C[-8], 2, FALSE)', 'J2:' + 'J' + IntToStr(qry_Tripulacion.RecordCount + 1), false, False, 'J:J');
      AsignaFormulas(Excel, 'Plantilla', 'L2', '=VLOOKUP(RC[-1],Cuentas!C[-11]:C[-10],2,FALSE)', 'L2:' + 'L' + IntToStr(qry_Tripulacion.RecordCount + 1), false, False, 'L:L');

    finally
      Screen.Cursor := Cursor;
    end;
  except
    on e: Exception do
    begin
      Excel.Quit;
      Excel := Null;
      ShowMessage(e.Message);
    end;
  end;
end;

procedure TfrmLista_personal.btnGuardarClick(Sender: TObject);
var
  cursor: TCursor;
begin
  try
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    try
      if Length(Trim(dbedtFicha.Text)) = 0  then
      begin
        MessageDlg('El campo [ficha] debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if dbedtFicha.CanFocus then
          dbedtFicha.SetFocus;
        Exit;
      end;

      if Length(Trim(dbedtNombre.Text)) = 0  then
      begin
        MessageDlg('El campo [Nombre] debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if dbedtNombre.CanFocus then
          dbedtNombre.SetFocus;
        Exit;
      end;

      if Length(Trim(dbedtOrden.Text)) = 0  then
      begin
        MessageDlg('El campo [Orden] debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if dbedtOrden.CanFocus then
          dbedtOrden.SetFocus;
        Exit;
      end;

      if Length(Trim(dbedtNombre.Text)) = 0  then
      begin
        MessageDlg('El campo [Nombre] debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if dbedtNombre.CanFocus then
          dbedtNombre.SetFocus;
        Exit;
      end;

      if DbLCbbCategoria.KeyValue = Null  then
      begin
        MessageDlg('Debe seleccionar una categoría antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if DbLCbbCategoria.CanFocus then
          DbLCbbCategoria.SetFocus;
        Exit;
      end;

      if Length(Trim(dbedtCabina.Text)) = 0  then
      begin
        MessageDlg('El campo [Cabina] debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if dbedtCabina.CanFocus then
          dbedtCabina.SetFocus;
        Exit;
      end;

      if DbLCbbCompania.KeyValue = Null  then
      begin
        MessageDlg('Debe seleccionar una [Compañía] antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if DbLCbbCompania.CanFocus then
          DbLCbbCompania.SetFocus;
        Exit;
      end;

      if DbLCbbCuenta.KeyValue = Null  then
      begin
        MessageDlg('Debe seleccionar una [Cuenta] antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
        if DbLCbbCuenta.CanFocus then
          DbLCbbCuenta.SetFocus;
        Exit;
      end;
      if zTripulacionDiariaUpt.State = dsInsert then
      begin
        zTripulacionDiariaUpt.FieldByName('sIdPersonal').AsString := dbedtsIdPersonal.Text;
        zTripulacionDiariaUpt.FieldByName('sNacionalidad').AsString := '';
        zTripulacionDiariaUpt.FieldByName('INacionales').AsString := '1';
        zTripulacionDiariaUpt.FieldByName('IExtranjeros').AsString := '0';
        zTripulacionDiariaUpt.FieldByName('sContrato').AsString := global_contrato_barco;
        zTripulacionDiariaUpt.FieldByName('dIdFecha').AsDateTime := tdIdFecha.Date;
        zTripulacionDiariaUpt.FieldByName('sIdTurno').AsString := QryTurnos.FieldByName('sIdTurno').AsString;
      end;
      if zTripulacionDiariaUpt.State in [dsInsert, dsEdit] then
      begin
        zTripulacionDiariaUpt.FieldByName('sIdCategoria').AsString := zLookUp_tripulacion.FieldByName('sIdCategoria').AsString;
        zTripulacionDiariaUpt.Post;
        zTripulacionDiariaUpt.ApplyUpdates;
      end;
      Forma.Close;
    finally
      Screen.Cursor := Cursor;
    end;
  Except
    on e: Exception do
      MessageDlg('Avisar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TfrmLista_personal.btnNuevoClick(Sender: TObject);
var
  Cursor: TCursor;
begin
  try
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    try
      if Assigned(FindComponent('FrmAdd')) then
        TForm(FindComponent('FrmAdd')).Destroy;

      forma := TForm.Create(Self);
      forma.Name := 'FrmAdd';
      forma.OnShow := onShowCustomForm;
      forma.Caption := 'Nuevo registro.';
      forma.Height := 330;
      forma.Width := 370;
      forma.BorderStyle := bsDialog;
      pnlDatos.Parent := forma;
      pnlDatos.Align := alClient;

      ///lblTripulacion.Caption := '';
      ///
      with zlookup_Tripulacion do
      begin
        Active := False;
        Params.ParamByName('Contrato').AsString := Global_Contrato_barco;
        Params.ParamByName('Categoria').AsString := tsIdCategoria.KeyValue;
        Params.ParamByName('Fecha').AsDate := fechaAntes;
        Open;
      end;

//      if ZLookTripulacion.RecordCount > 0 then
//      begin
//        //ZLookTripulacion.First;
//        //tsIdTripulacion.keyvalue := ZLookTripulacion.FieldValues['sIdTripulacion'];
//        //lblTripulacion.Caption := 'Tripulacion ' + tsIdTripulacion.keyvalue;
//      end;

      with zLookup_Compania do
      begin
        active := False;
        Open;
      end;

      with zLookup_Cuentas do
      begin
        active := False;
        Open;
      end;

      with zTripulacionDiariaUpt do
      begin
        Active := False;
        Open;
      end;

      pnlDatos.Visible := True;
      if forma.ShowModal <> mrCancel then
        qry_Tripulacion.Refresh;
    finally
      pnlDatos.Visible := False;
      pnlDatos.Align := alNone;
      pnlDatos.Parent := Self;
      forma.Destroy;
      Screen.Cursor := Cursor;
    end;
  except
    on E: Exception do
    begin
        MessageDlg(e.Message, mtInformation, [mbOk], 0)
    end;
  end;
end;


procedure TfrmLista_personal.btnPrinterClick(Sender: TObject);
begin
    try
          if (qry_Tripulacion.Active) and (qry_Tripulacion.RecordCount > 0) then
            procreporteTripulacion(Global_Contrato_barco, QryTurnos.FieldValues['sIdTurno'], tdIdFecha.DateTime, frmLista_personal, frxTripulacion.OnGetValue, connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, 'reporteBarco'), 'Barco')
          else
            showmessage('No hay datos para imprimir');
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Tripulacion Diaria', 'Al imprimir', 0);
        end;
      end;
end;

procedure TfrmLista_personal.tsIdCategoriaEnter(Sender: TObject);
begin
  tsidcategoria.Color := global_color_entrada;
end;

procedure TfrmLista_personal.tsIdCategoriaExit(Sender: TObject);
begin
  tsIdcategoria.Color := global_Color_salida;
  tsOrdenes.KeyValue := Null;
  with qry_tripulacion do
  begin
    Active := False;
    SQL.Clear;
    SQL.Add('SELECT a.*, c.sDescripcion as compania, cu.sDescripcion as Cuenta FROM tripulaciondiaria_listado as a ' +
      ' Left join compersonal as c on (c.sIdCompania = a.sIdCompania) ' +
      ' Left join cuentas as cu on (cu.sIdCuenta = a.sIdCuenta) ' +
      ' where sContrato = :contrato and dIdfecha = :fecha ' +
      ' And sIdCategoria= :Categoria and sIdTurno =:Turno ');
    params.ParamByName('contrato').DataType  := ftString;
    params.ParamByName('contrato').Value     := global_contrato_barco;
    params.ParamByName('fecha').DataType     := ftDate;
    params.ParamByName('fecha').Value        := tdIdFecha.Date;
    params.ParamByName('categoria').DataType := ftString;
    params.ParamByName('categoria').Value    := tsIdCategoria.KeyValue;
    params.ParamByName('turno').DataType     := ftString;
    params.ParamByName('turno').Value        := QryTurnos.FieldValues['sIdTurno'];
    Open;
    if qry_Tripulacion.RecordCount = 0 then
       ttripulacion_nacionales.Text := '0';
  end;

  //Ahora sumamos las cantidades de qie existen por categorias..
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('SELECT sum(iNacionales) as cantidad FROM tripulaciondiaria_listado where sContrato = :contrato and dIdfecha = :fecha '+
      'And sIdCategoria= :Categoria and sIdTurno =:Turno group by sContrato ');
  connection.QryBusca.params.ParamByName('contrato').DataType  := ftString;
  connection.QryBusca.params.ParamByName('contrato').Value     := global_contrato_barco;
  connection.QryBusca.params.ParamByName('fecha').DataType     := ftDate;
  connection.QryBusca.params.ParamByName('fecha').Value        := tdIdFecha.Date;
  connection.QryBusca.params.ParamByName('categoria').DataType := ftString;
  connection.QryBusca.params.ParamByName('categoria').Value    := tsIdCategoria.KeyValue;
  connection.QryBusca.params.ParamByName('turno').DataType     := ftString;
  connection.QryBusca.params.ParamByName('turno').Value        := QryTurnos.FieldValues['sIdTurno'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
     ttripulacion_nacionales.Text := connection.QryBusca.FieldValues['cantidad']
  else
     ttripulacion_nacionales.Text := '0';

end;

 //SOAD -> Busqued de tripulacion por id.

procedure TfrmLista_personal.cmbTurnosExit(Sender: TObject);
begin
  tdIdFecha.OnExit(sender);
  tsOrdenes.KeyValue := Null;
end;

procedure TfrmLista_personal.cmbTurnosKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 13 then
    if cmbTurnos.KeyValue <> Null then
      cmbTurnosExit(nil);
end;

procedure TfrmLista_personal.cmdAgregarClick(Sender: TObject);
begin

  try
    tripulacionDiaria.Open;
    tripulacionDiaria.Append;
    tripulacionDiaria.FieldValues['sContrato'] := global_contrato_barco;
    tripulacionDiaria.FieldValues['sIdTurno'] := QryTurnos.FieldValues['sIdTurno'];
    tripulacionDiaria.FieldValues['dIdFecha'] := tdIdFecha.Date;
    tripulacionDiaria.FieldValues['sIdCategoria'] := tsIdCategoria.KeyValue;
    tripulacionDiaria.FieldValues['sIdTripulacion'] := tsIdTripulacion.KeyValue;
    tripulacionDiaria.FieldValues['iNacionales'] := txtCantidad.Text;
    tripulacionDiaria.FieldValues['iExtranjeros'] := 0;
    //tripulacionDiaria.FieldValues['lAplicaPernocta'] := 'Si';
    tripulacionDiaria.FieldValues['sOrden'] := '';
    tripulacionDiaria.Post;

    txtCantidad.Text := '0';
    qry_tripulacion.Refresh;

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Tripulacion Diaria', 'Al salvar registro', 0);
    end;
  end;

  lblTripulacion.Caption := '';
  if ZLookTripulacion.RecordCount > 0 then begin
    ZLookTripulacion.First;
    tsIdTripulacion.keyvalue := ZLookTripulacion.FieldValues['sIdTripulacion'];
    lblTripulacion.Caption := 'Tripulacion ' + tsIdTripulacion.keyvalue;
  end;

end;

procedure TfrmLista_personal.cmdImportarClick(Sender: TObject);
var
  x, y, i: Integer;
  sNombre, sIdCategoria, sCategoria, sIdPersonal, sDescripcion, sCompañia, sGenero, sFolio, sFechaVigencia,
  sCabina, sNacionalidad, sLibreta, sValue, sOrden, sIdCompania, sIdCuenta, sRFC: string;
  dFechaInicio, dFechaFinal, dVigencia : tdate;
begin
  OpenXLS.Title := 'Inserta Archivo de Consulta';
  if OpenXLS.Execute then
  begin
    tsArchivo.Text := OpenXLS.FileName;
    // soad - > Llenado del array..
    for x := 1 to 26 do
      columnas[x] := Chr(64 + x);

    i := 27;
    for x := 1 to 9 do
    begin
      for y := 1 to 26 do
      begin
          columnas[i] := Chr(64 + x) + Chr(64 + y);
          i := i + 1;
      end;
    end;
  end;

  try
    flcid := GetUserDefaultLCID;
    ExcelApplication1.Connect;
    ExcelApplication1.Visible[flcid] := true;
    ExcelApplication1.UserControl := true;

    ExcelWorkbook1.ConnectTo(ExcelApplication1.Workbooks.Open(tsArchivo.Text,
    emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam,
    emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, flcid));

    {Antes de iniciar peguntamos al Usuraio si Tomamos Datos del Contrato Actual o el de Excel..}
    lContratoActual := False;
    Resp := MessageDlg('¿Desea Importar los Datos con el nombre del Contrato Actual "' + global_contrato + '" ?', mtConfirmation, [mbYes, mbCancel], 0);
    case
      Resp
      of
        mrYes: lContratoActual := True;
        mrCancel: raise Exception.Create('Proceso Cancelado por el Usuario.');
    end;

    if lContratoActual then
    begin
      ExcelWorksheet1.ConnectTo(ExcelWorkbook1.Sheets.Item[1] as ExcelWorkSheet);
      Fila := 2;
      sValue := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('DELETE FROM tripulaciondiaria_listado Where sContrato = :contrato And dIdFecha = :fecha ');
      Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Contrato').Value    := Global_Contrato;
      Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
      Connection.zCommand.Params.ParamByName('fecha').Value       := tdIdFecha.Date;
      connection.zCommand.ExecSQL();

      //Procedemos a leer el archivo de Excel..
      sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
      while (sValue <> '') do
      begin
        sNombre := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
        sIdCategoria := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
        sCategoria   := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
        sOrden       := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
        sIdPersonal  := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
        sDescripcion := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
        sCabina      := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
        sIdCompania := ExcelWorkSheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;
        sIdCuenta := ExcelWorkSheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2;
        sRFC := ExcelWorkSheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;

        connection.zCommand.Active := False;
        Connection.zCommand.SQL.Clear;
        Connection.zCommand.SQL.Add('INSERT INTO tripulaciondiaria_listado ( sContrato , sIdTurno, dIdFecha, sIdCategoria, sIdTripulacion, sOrden, sNombre, sIdPersonal, sDescripcion, '+
                                    'sNumeroCabina, sNacionalidad, iNacionales, iExtranjeros, sIdCompania, sIdCuenta, sRfc) ' +
                                    'VALUES (:contrato, :turno, :fecha , :categoria, :tripulacion, :orden, :nombre, :idpersonal, :descripcion, '+
                                    ':cabina, :nacionalidad, :nacionales, :extranjeros, :IdCompania, :IdCuenta, :RFC)');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := global_contrato;
        Connection.zCommand.Params.ParamByName('turno').DataType    := ftString;
        Connection.zCommand.Params.ParamByName('turno').value       := QryTurnos.FieldByName('sIdTurno').AsString;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('categoria').DataType:= ftString;
        if trim(sIdCategoria) = '' then
           Connection.zCommand.Params.ParamByName('categoria').value   := categorias.FieldByName('sIdCategoria').AsString
        else
           Connection.zCommand.Params.ParamByName('categoria').value   := sIdCategoria;
        Connection.zCommand.Params.ParamByName('tripulacion').DataType := ftString;
        Connection.zCommand.Params.ParamByName('tripulacion').value    := sValue;
        Connection.zCommand.Params.ParamByName('orden').DataType       := ftString;
        Connection.zCommand.Params.ParamByName('orden').value          := sOrden;
        Connection.zCommand.Params.ParamByName('Nombre').DataType      := ftString;
        Connection.zCommand.Params.ParamByName('Nombre').value         := Trim(sNombre);
        Connection.zCommand.Params.ParamByName('IdPersonal').DataType  := ftString;
        Connection.zCommand.Params.ParamByName('IdPersonal').value     := sIdPersonal;
        Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString;
        Connection.zCommand.Params.ParamByName('Descripcion').value    := sDescripcion;
        Connection.zCommand.Params.ParamByName('cabina').DataType      := ftString;
        Connection.zCommand.Params.ParamByName('cabina').value         := sCabina;
        Connection.zCommand.Params.ParamByName('nacionalidad').DataType := ftString;
        Connection.zCommand.Params.ParamByName('nacionalidad').value    := sNacionalidad;
        Connection.zCommand.Params.ParamByName('nacionales').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('nacionales').value      := 1;
        Connection.zCommand.Params.ParamByName('extranjeros').DataType  := ftInteger;
        Connection.zCommand.Params.ParamByName('extranjeros').value     := 0;
        Connection.zCommand.Params.ParamByName('IdCompania').DataType := FtString;
        Connection.zCommand.Params.ParamByName('IdCompania').Value := sIdCompania;
        Connection.zCommand.Params.ParamByName('IdCuenta').DataType := FtString;
        Connection.zCommand.Params.ParamByName('IdCuenta').Value := sIdCuenta;
        Connection.zCommand.Params.ParamByName('RFC').DataType := FtString;
        Connection.zCommand.Params.ParamByName('RFC').Value := sRFC;
        connection.zCommand.ExecSQL;

        Fila := Fila + 1;
        sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
      end;
      tdIdFecha.OnExit(sender);
      MessageDlg('Proceso Terminado con exito...', mtInformation, [mbOk], 0);
    end;
  except
    on E: Exception do
    begin
        MessageDlg(e.Message, mtInformation, [mbOk], 0)
    end;
  end;
  //Desconectamos la aplicacion..
  ExcelApplication1.Disconnect;
end;

procedure TfrmLista_personal.cmdSalirClick(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
  cmdSalir.Visible := False;

end;

procedure TfrmLista_personal.Panel2Click(Sender: TObject);
begin
  //Panel2.Visible := False;
  //cmdAgregar.Visible := False;
end;

procedure TfrmLista_personal.PopupMenu1Popup(Sender: TObject);
var
    zPersonal : tzReadOnlyQuery;
begin
    zPersonal := tzReadOnlyQuery.Create(self);
    zPersonal.Connection := connection.zConnection;

    zPersonal.Active := False;
    zPersonal.SQL.Clear;
    zPersonal.SQL.Add('select sAnexo from anexos where sTipo = "PERSONAL"');
    zPersonal.Open;

    zPersonal.Destroy;
end;

end.

