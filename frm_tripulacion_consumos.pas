unit frm_tripulacion_consumos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms,
  Dialogs, StdCtrls, Mask, DBCtrls, ComCtrls, Grids, DBGrids, global, frm_connection, utilerias,
  DB, ADODB, Buttons, ExtCtrls, frxClass, frxDBSet, ZAbstractRODataset,
  ZDataset, ZAbstractDataset, Controls, Menus, UnitExcepciones, udbgrid, UFunctionsGHH,
  DBDateTimePicker, UnitValidacion, rxToolEdit, rxCurrEdit, RXDBCtrl,
  cxGraphics, cxLookAndFeels, cxLookAndFeelPainters, dxSkinsCore,
  dxSkinDevExpressStyle, dxSkinFoggy, cxButtons, cxControls, cxStyles,
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit,
  cxNavigator, cxDBData, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxGridCustomView, cxGrid, cxDBLookupComboBox,
  ZSqlUpdate;

type
  TfrmTripulacionConsumos = class(TForm)
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
    qry_TripulaciondCantidad: TFloatField;
    qry_TripulacionsDescripcion: TStringField;
    pernoctas: TZReadOnlyQuery;
    ds_pernoctas: TDataSource;
    grid_tripulacion: TcxGrid;
    CxGridMoePersonal: TcxGridDBTableView;
    CxOrdenaPersonal: TcxGridDBColumn;
    CxColumnCxGridMoePersonalColumn1: TcxGridDBColumn;
    CxColumnCxGridMoePersonalColumn2: TcxGridDBColumn;
    CxColumnCxGridMoePersonalColumn3: TcxGridDBColumn;
    CxLevel1: TcxGridLevel;
    qry_TripulacionsIdEquipo: TStringField;
    Grid_Bitacora: TcxGrid;
    BView_Actividades: TcxGridDBTableView;
    mDescripcion: TcxGridDBColumn;
    Grid_BitacoraLevel1: TcxGridLevel;
    qry_TripulacionsDescripcionFolio: TStringField;
    ds_foliosGrid: TDataSource;
    zq_foliosGrid: TZReadOnlyQuery;
    zq_foliosGridsNumeroOrden: TStringField;
    zq_foliosGridsIdFolio: TStringField;
    qry_TripulacioniId: TIntegerField;
    USqlTripulacion: TZUpdateSQL;
    tsEliminar: TcxButton;
    ds_notas: TDataSource;
    zqNotas: TZQuery;
    cxOrdenar: TcxGridDBColumn;
    qry_TripulacioniOrden: TIntegerField;
    VerColumnaOrdenamiento1: TMenuItem;
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
    procedure BView_ActividadesDblClick(Sender: TObject);
    procedure BView_ActividadesKeyPress(Sender: TObject; var Key: Char);
    procedure CxGridMoePersonalDblClick(Sender: TObject);
    procedure tsEliminarClick(Sender: TObject);
    procedure CreaNota;
    procedure VerColumnaOrdenamiento1Click(Sender: TObject);



  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTripulacionConsumos: TfrmTripulacionConsumos;
  sFirma_PEP, sPuesto_PEP, sFirma_Contratista, sPuesto_Contratista: string;
  fechaAntes: tDate;
  utgrid: ticdbgrid;
implementation

uses frm_tripulacion_diaria;

{$R *.dfm}

procedure TfrmTripulacionConsumos.tdIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
  {if key = #13 then
    tsIdCategoria.SetFocus  }
end;

procedure TfrmTripulacionConsumos.tsIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
   if key = #13 then
      grid_tripulacion.SetFocus   
end;

procedure TfrmTripulacionConsumos.tsIdFolioEnter(Sender: TObject);
begin
    tsIdFolio.Color := global_color_entrada;
end;

procedure TfrmTripulacionConsumos.tsIdFolioExit(Sender: TObject);
begin
    tsIdFolio.Color := global_color_salida;
end;

procedure TfrmTripulacionConsumos.tsIdFolioKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key =#13 then
       grid_tripulacion.SetFocus;

    qry_tripulacion.Active := False;
    qry_tripulacion.params.ParamByName('contrato').AsString := global_contrato;
    qry_tripulacion.params.ParamByName('fecha').AsDate   := tdIdFecha.Date;
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

procedure TfrmTripulacionConsumos.tsIdTripulacionEnter(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_entrada;
end;

procedure TfrmTripulacionConsumos.tsIdTripulacionExit(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_salida;

end;

procedure TfrmTripulacionConsumos.tsIdTripulacionKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       txtCantidad.SetFocus;
end;

procedure TfrmTripulacionConsumos.ttripulacion_nacionalesKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    ttripulacion_extranjeros.SetFocus
end;

procedure TfrmTripulacionConsumos.txtCantidadChange(Sender: TObject);
begin
  tEditChangef(txtCantidad, 'Cantidad');
end;

procedure TfrmTripulacionConsumos.txtCantidadEnter(Sender: TObject);
begin
  txtcantidad.Color := global_color_entrada;
end;

procedure TfrmTripulacionConsumos.txtCantidadExit(Sender: TObject);
begin
  txtcantidad.Color := global_color_salida
end;

procedure TfrmTripulacionConsumos.txtCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTedit(txtCantidad, key) then
    key := #0
  else
    if key =#13 then
       cmdAgregar.SetFocus;
  

end;

procedure TfrmTripulacionConsumos.VerColumnaOrdenamiento1Click(Sender: TObject);
begin
    if cxOrdenar.Visible = False then    
       cxOrdenar.Visible := True
    else
       cxOrdenar.Visible := False;
end;

procedure TfrmTripulacionConsumos.ttripulacion_extranjerosKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdidFecha.SetFocus
end;

//Inicio LA forma

procedure TfrmTripulacionConsumos.FormShow(Sender: TObject);
begin

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

  zq_FoliosGrid.Active := False;
  zq_FoliosGrid.ParamByName('contrato').AsString := global_contrato;
  zq_FoliosGrid.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  zq_FoliosGrid.Open;

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
  ZLookTripulacion.ParamByName('Contrato').AsString := global_contrato_barco;
  ZLookTripulacion.Open;

  if global_fecha_rd = 0 then
     tdIdFecha.Date := now
  else
     tdIdFecha.Date := global_fecha_rd;
  tdIdFecha.OnExit(sender);

  CreaNota;

  grid_tripulacion.SetFocus;
end;

procedure TfrmTripulacionConsumos.frxTripulacionGetValue(const VarName: string;
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

procedure TfrmTripulacionConsumos.tdIdFechaExit(Sender: TObject);
begin
  tdidfecha.Color := global_color_salida;

  Folios.Active := False;
  Folios.ParamByName('contrato').AsString := global_contrato;
  Folios.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  Folios.Open;

  zq_FoliosGrid.Active := False;
  zq_FoliosGrid.ParamByName('contrato').AsString := global_contrato;
  zq_FoliosGrid.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  zq_FoliosGrid.Open;

  tsIdFolio.KeyValue := '<todos>';

  qry_tripulacion.Active := False;
  qry_tripulacion.params.ParamByName('contrato').asString := global_contrato;
  qry_tripulacion.params.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  qry_tripulacion.params.ParamByName('folio').DataType    := ftString;
  if tsIdFolio.KeyValue = '<todos>' then
     qry_tripulacion.params.ParamByName('folio').Value := '%'
  else
     qry_tripulacion.params.ParamByName('folio').Value := tsIdFolio.KeyValue;
  qry_tripulacion.Open;


  qryTripulacionPernoctaFuera.Active := false;
  qryTripulacionPernoctaFuera.params.ParamByName('contrato').DataType := ftString;
  qryTripulacionPernoctaFuera.params.ParamByName('contrato').Value := global_contrato_barco;
  qryTripulacionPernoctaFuera.params.ParamByName('fecha').DataType := ftDate;
  qryTripulacionPernoctaFuera.params.ParamByName('fecha').Value := tdIdFecha.Date;
  qryTripulacionPernoctaFuera.params.ParamByName('turno').DataType := ftString;
  qryTripulacionPernoctaFuera.params.ParamByName('turno').Value := QryTurnos.FieldValues['sIdTurno'];
  qryTripulacionPernoctaFuera.Open;

  CreaNota;
  cxOrdenar.Visible := False;
end;


//Consultar la Categoria

procedure TfrmTripulacionConsumos.ActualizaPersonalClick(Sender: TObject);
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

procedure TfrmTripulacionConsumos.btnCancelClick(Sender: TObject);
begin
  close;
end;

procedure TfrmTripulacionConsumos.btnUpdateClick(Sender: TObject);
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

procedure TfrmTripulacionConsumos.BView_ActividadesDblClick(Sender: TObject);
begin
    if BView_Actividades.OptionsView.CellAutoHeight then
       BView_Actividades.OptionsView.CellAutoHeight := False
    else
       BView_Actividades.OptionsView.CellAutoHeight := True;
end;

procedure TfrmTripulacionConsumos.BView_ActividadesKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       txtCantidad.SetFocus;
end;

procedure TfrmTripulacionConsumos.tGuardarClick(Sender: TObject);
begin
    zqNotas.Edit;
    zqNotas.FieldByName('sNotaGeneral').AsString := dbMemo1.Text;
    zqNotas.Post;
    messagedLG('Las notas se guardaron correctamente!', mtInformation, [mbOk], 0);
end;

procedure TfrmTripulacionConsumos.qryTripulacionPernoctaFueraAfterScroll(
  DataSet: TDataSet);
begin
(*      qryTripulacionPernoctaFuera.FieldValues['sContrato'] := global_contrato;
      qryTripulacionPernoctaFuera.FieldValues['sIdTurno'] := QryTurnos.FieldValues['sIdTurno'];
      qryTripulacionPernoctaFuera.FieldValues['dIdFecha'] := tdIdFecha.Date;*)
end;

procedure TfrmTripulacionConsumos.qry_tripulacionAfterInsert(
  DataSet: TDataSet);
begin
  qry_tripulacion.Cancel
end;

procedure TfrmTripulacionConsumos.qry_TripulacionAfterScroll(DataSet: TDataSet);
begin
  with connection.QryBusca do
  begin
    Active := False;
    SQL.Clear;
    SQL.Add('SELECT sum(dCantidad) as cantidad FROM consumosdecombustibleporequipo where sContrato = :contrato and dIdfecha = :fecha ' +
      'And sNumeroOrden like :Orden group by sContrato ');
    params.ParamByName('contrato').DataType := ftString;
    params.ParamByName('contrato').Value := global_contrato;
    params.ParamByName('fecha').DataType := ftDate;
    params.ParamByName('fecha').Value := tdIdFecha.Date;
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

procedure TfrmTripulacionConsumos.qry_TripulacionBeforePost(DataSet: TDataSet);
begin
  zq_foliosGrid.Locate('sIdFolio', qry_Tripulacion.FieldByName('sDescripcionFolio').AsString, [loCaseInsensitive]);
  qry_Tripulacion.FieldValues['sNumeroOrden'] := zq_foliosGrid.FieldByName('sNumeroOrden').AsString;
  if qry_Tripulacion.FieldValues['dCantidad'] < 0 then
     qry_Tripulacion.cancel;
end;

procedure TfrmTripulacionConsumos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  action := cafree;
end;

procedure TfrmTripulacionConsumos.tAceptarClick(Sender: TObject);
begin
    //Antes de guardar el registro comparamos si la descripcion ya existe..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select sDescripcion from equipos where sContrato =:Contrato and sDescripcion =:Descripcion ');
    connection.zCommand.ParamByName('Contrato').AsString    := global_contrato_barco;
    connection.zCommand.ParamByName('Descripcion').AsString := tdCategoria.Text;
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
    begin
        messageDLG('El equipo ya existe favor de verificar!', mtInformation, [mbOk],0);
        ZLookTripulacion.Cancel;
        tCancelar.Click;
    end;

    ZLookTripulacion.FieldValues['sContrato']     := global_contrato_barco;
    ZLookTripulacion.FieldValues['sIdEquipo']     := tdIdCategoria.Text;
    ZLookTripulacion.FieldValues['sDescripcion']  := tdCategoria.Text;
    ZLookTripulacion.FieldValues['sIdTipoEquipo'] := 'EQ-C';
    ZLookTripulacion.FieldValues['sMedida']       := 'LITROS';
    ZLookTripulacion.FieldValues['dCantidad']     := 0;
    ZLookTripulacion.FieldValues['lImprime']      := 'Si';
    ZLookTripulacion.FieldValues['lAplicaDiesel'] := 'Si';
    ZLookTripulacion.FieldValues [ 'dCostoMN' ] := 0 ;
    ZLookTripulacion.FieldValues [ 'dCostoDLL' ] := 0 ;
    ZLookTripulacion.FieldValues [ 'dVentaMN' ] := 0 ;
    ZLookTripulacion.FieldValues [ 'dVentaDLL' ] := 0 ;
    ZLookTripulacion.FieldValues [ 'lCobro' ] := 'No' ;
    ZLookTripulacion.FieldValues [ 'lProrrateo' ] := 'No' ;
    ZLookTripulacion.FieldValues [ 'iJornada' ] := 0 ;
    ZLookTripulacion.FieldValues [ 'lDistribuye' ] := 'No' ;
    ZLookTripulacion.FieldValues [ 'lCuadraEquipo' ]   := 'No' ;
    ZLookTripulacion.FieldValues [ 'dFechaInicio' ] := now; ;
    ZLookTripulacion.FieldValues [ 'dFechaFinal' ]  := now ;
    ZLookTripulacion.Post;
    ZLookTripulacion.Refresh;

    tCancelar.Click;
end;

procedure TfrmTripulacionConsumos.tAgregarClick(Sender: TObject);
var
   Maximo : integer;
begin
   Panel2.Height := 345;
   cmdAgregar.Enabled := False;
   cmdSalir.Enabled   := False;
   grid_bitacora.Enabled := False;
   txtCantidad.Enabled     := FAlse;

    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select max(iItemOrden) as maximo from equipos where sContrato =:Contrato group by sContrato ');
    connection.zCommand.ParamByName('Contrato').AsString := global_Contrato_barco;
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
       Maximo := connection.zCommand.FieldByName('maximo').AsInteger +1
    else
       Maximo := 1;

   ZLookTripulacion.Append;
   ZLookTripulacion.FieldValues['iItemOrden']    := Maximo;
   tdIdCategoria.Text := 'EQ-00'+ IntToStr(Maximo);
   tdCategoria.SetFocus;
end;

procedure TfrmTripulacionConsumos.tCancelarClick(Sender: TObject);
begin
   Panel2.Height := 245;
   cmdAgregar.Enabled := True;
   cmdSalir.Enabled   := True;
   grid_bitacora.Enabled := True;
   txtCantidad.Enabled   := True;
   ZLookTripulacion.Cancel;
   grid_bitacora.SetFocus;
end;

procedure TfrmTripulacionConsumos.tdCategoriaEnter(Sender: TObject);
begin
   tdCategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionConsumos.tdCategoriaExit(Sender: TObject);
begin
   tdCategoria.Color := global_color_salida;
end;

procedure TfrmTripulacionConsumos.tdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tAceptar.SetFocus
end;

procedure TfrmTripulacionConsumos.tdIdCategoriaEnter(Sender: TObject);
begin
   tdIdCategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionConsumos.tdIdCategoriaExit(Sender: TObject);
begin
   tdIdCategoria.Color := global_color_salida;
end;

procedure TfrmTripulacionConsumos.tdIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tdcategoria.SetFocus;
end;

procedure TfrmTripulacionConsumos.tdIdFechaEnter(Sender: TObject);
begin
  tdIdFecha.Color := global_Color_entrada
end;


procedure TfrmTripulacionConsumos.qry_tripulacionCalcFields(
  DataSet: TDataSet);
begin
  if qry_tripulacion.RecordCount > 0 then
  begin
        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Add('Select sDescripcion from equipos Where ' +
          ' sContrato =:Contrato and sIdEquipo = :Equipo ');
        Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Contrato').Value    := global_contrato_barco;
        Connection.QryBusca.Params.ParamByName('Equipo').DataType   := ftString;
        Connection.QryBusca.Params.ParamByName('Equipo').Value      := qry_tripulacion.FieldByName('sIdEquipo').AsString;
        Connection.QryBusca.Open;

        if Connection.QryBusca.RecordCount > 0 then
          qry_TripulacionsDescripcion.Text := Connection.QryBusca.FieldValues['sDescripcion'];

  end;
end;



procedure TfrmTripulacionConsumos.ripulacinDiariaDiaAnterior1Click(
  Sender: TObject);
var
  dFechaAnterior : tDate;
  lNuevaVigencia : boolean;

begin
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('Select * from consumosdecombustibleporequipo Where sContrato =:Contrato And dIdFecha =:Fecha');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value    := Global_Contrato;
  Connection.qryBusca.Params.ParamByName('Fecha').Datatype    := ftDate;
  Connection.qryBusca.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
  Connection.qryBusca.Open;
  if Connection.qryBusca.RecordCount > 0 then
    if MessageDlg('Desea Eliminar todas los Equipos Existentes?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('Delete From consumosdecombustibleporequipo where sContrato = :contrato and dIdFecha = :fecha  ');
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
      connection.zCommand.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
      connection.zCommand.ExecSQL;
    end;

    if MessageDlg('Desea adicionar los equipos Existentes en el reporte anterior?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
        Connection.Auxiliar.SQL.Clear;
        Connection.Auxiliar.SQL.Add('Select * from consumosdecombustibleporequipo Where sContrato =:Contrato And dIdFecha =:Fecha');
        Connection.Auxiliar.Params.ParamByName('Contrato').DataType := ftString;
        Connection.Auxiliar.Params.ParamByName('Contrato').Value    := Global_Contrato;
        Connection.Auxiliar.Params.ParamByName('Fecha').Datatype    := ftDate;
        Connection.Auxiliar.Params.ParamByName('Fecha').Value       := tdIdFecha.Date - 1;
        Connection.Auxiliar.Open;


      if Connection.Auxiliar.RecordCount > 0 then
      begin
          Connection.zcommand.SQL.Clear;
          Connection.zcommand.SQL.Add('INSERT INTO consumosdecombustibleporequipo ( sContrato , didFecha , sIdEquipo, sNumeroOrden, dCantidad, iOrden) ' +
            ' VALUES (:contrato , :fecha  , :equipo , :Orden , :Cantidad, :ordenar )');
          while not Connection.Auxiliar.Eof do
          begin
            with connection do
            begin
                try
                  zcommand.params.ParamByName('contrato').DataType    := ftString;
                  zcommand.params.ParamByName('contrato').value       := Global_Contrato;
                  zcommand.params.ParamByName('fecha').DataType       := ftDate;
                  zcommand.params.ParamByName('fecha').value          := tdIdFecha.Date;
                  zcommand.params.ParamByName('equipo').DataType      := ftString;
                  zcommand.params.ParamByName('equipo').value         := Auxiliar.FieldValues['sIdEquipo'];
                  zcommand.params.ParamByName('orden').DataType       := ftString;
                  zcommand.params.ParamByName('orden').value          := Auxiliar.FieldValues['sNumeroOrden'];
                  zcommand.params.ParamByName('cantidad').DataType    := ftInteger;
                  zcommand.params.ParamByName('cantidad').value       := Auxiliar.FieldValues['dCantidad'];
                  zcommand.params.ParamByName('ordenar').DataType     := ftInteger;
                  zcommand.params.ParamByName('ordenar').value        := Auxiliar.FieldValues['iOrden'];
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

procedure TfrmTripulacionConsumos.btnDeleteClick(Sender: TObject);
begin
  if (not qry_Tripulacion.active) or (qry_Tripulacion.RecordCount < 1) then
    exit;
  Connection.zCommand.Active := False;
  Connection.zCommand.SQL.Clear;
  Connection.zCommand.SQL.Add('Delete from consumosdecombustibleporequipo where sContrato = :contrato and sIdEquipo = :Categoria ' +
    'And dIdFecha =:Fecha and sNumeroOrden =:Orden ');
  Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
  Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
  Connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
  Connection.zCommand.Params.ParamByName('Orden').Value := qry_Tripulacion.FieldValues['sNumeroOrden'];
  Connection.zCommand.Params.ParamByName('Categoria').DataType := ftString;
  Connection.zCommand.Params.ParamByName('Categoria').Value := qry_Tripulacion.FieldValues['sIdEquipo'];
  Connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
  Connection.zCommand.Params.ParamByName('Fecha').Value := tdIdFecha.date;
  Connection.zCommand.ExecSQL();
  qry_Tripulacion.Refresh;
end;

procedure TfrmTripulacionConsumos.btnNuevoClick(Sender: TObject);
begin
  if (tsIdFolio.KeyValue = null) or (tsIdFolio.KeyValue = '<todos>') then begin
    ShowMessage('Favor de Seleccionar un Folio');
    exit;
  end;

  ZLookTripulacion.First;
  txtCantidad.Text := '0';
  Panel2.Visible := True;
  tsIdTripulacion.SetFocus;
end;


procedure TfrmTripulacionConsumos.btnPrinterClick(Sender: TObject);
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

procedure TfrmTripulacionConsumos.tsEliminarClick(Sender: TObject);
begin
    //Antes de eliminar verificarmos que no esté reportado..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select sIdEquipo from consumosdecombustibleporequipo where sContrato =:Contrato and sIdequipo =:Equipo ');
    connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
    connection.zCommand.ParamByName('Equipo').AsString   := ZLookTripulacion.FieldByName('sIdEquipo').AsString;
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
    begin
        messageDLG('No se puede eliminar, el equipo ya está reportado!', mtInformation, [mbOk], 0);
        exit;
    end
    else
        ZLookTripulacion.Delete;

end;

procedure TfrmTripulacionConsumos.tsIdCategoriaEnter(Sender: TObject);
begin
  tsidcategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionConsumos.tsIdCategoriaExit(Sender: TObject);
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

procedure TfrmTripulacionConsumos.cmbTurnosExit(Sender: TObject);
begin
  tdIdFecha.OnExit(sender);
end;

procedure TfrmTripulacionConsumos.cmdAgregarClick(Sender: TObject);
begin
  try
    tripulacionDiaria.ParamByName('Contrato').AsString := global_contrato;
    tripulacionDiaria.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
    tripulacionDiaria.Open;
    tripulacionDiaria.Append;
    tripulacionDiaria.FieldValues['sContrato'] := global_contrato;
    tripulacionDiaria.FieldValues['dIdFecha']  := tdIdFecha.Date;
    tripulacionDiaria.FieldValues['sIdEquipo'] := zLookTripulacion.FieldByName('sIdEquipo').AsString;
    tripulacionDiaria.FieldValues['sNumeroOrden'] := tsIdFolio.KeyValue;
    tripulacionDiaria.FieldValues['dCantidad']    := txtCantidad.Text;
    tripulacionDiaria.FieldValues['iOrden']       := zLookTripulacion.FieldByName('iItemOrden').AsInteger;
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
  end;
  grid_bitacora.SetFocus;

end;

procedure TfrmTripulacionConsumos.cmdSalirClick(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
  cmdSalir.Visible := False;    
end;

procedure TfrmTripulacionConsumos.CxGridMoePersonalDblClick(Sender: TObject);
begin
    if CxGridMoePersonal.OptionsView.CellAutoHeight then
       CxGridMoePersonal.OptionsView.CellAutoHeight := False
    else
       CxGridMoePersonal.OptionsView.CellAutoHeight := True;
end;

procedure TfrmTripulacionConsumos.Panel2Click(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
end;

procedure TfrmTripulacionConsumos.PopupMenu1Popup(Sender: TObject);
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

procedure TfrmTripulacionConsumos.CreaNota;
begin
    zqNotas.Active := False;
    zqNotas.ParamByName('Contrato').AsString := global_contrato;
    zqNotas.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
    zqNotas.Open;

    if zqNotas.RecordCount = 0 then
    begin
        zqNotas.Append;
        zqNotas.FieldByName('sContrato').AsString    := global_contrato;
        zqNotas.FieldByName('iOrden').AsInteger      := 0;
        zqNotas.FieldByName('sNotaGeneral').AsString := '*';
        zqNotas.FieldByName('dIdFecha').AsDateTime   := tdIdFecha.Date;
        zqNotas.FieldByName('sContrato').AsString    := global_contrato;
        zqNotas.FieldByName('lAplicaLibro').AsString := 'No';
        zqNotas.FieldByName('eAplicaResumenPersonal').AsString := 'No';
        zqNotas.FieldByName('lAplicaConsumos').AsString        := 'Si';
        zqNotas.Post;
    end;
end;

end.

