unit frm_tripulacion_ajustes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms, unitTarifa,
  Dialogs, StdCtrls, Mask, DBCtrls, ComCtrls, Grids, DBGrids, global, frm_connection, utilerias,
  DB, ADODB, Buttons, ExtCtrls, frxClass, frxDBSet, ZAbstractRODataset,
  ZDataset, ZAbstractDataset, Controls, Menus, UnitExcepciones, udbgrid, UFunctionsGHH,
  DBDateTimePicker, UnitValidacion, rxToolEdit, rxCurrEdit, RXDBCtrl,
  cxGraphics, cxLookAndFeels, cxLookAndFeelPainters, dxSkinsCore,
  dxSkinDevExpressStyle, dxSkinFoggy, cxButtons, ZSqlUpdate,
  AdvSmoothProgressBar;

type
  TfrmTripulacionAjustes = class(TForm)
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
    pernoctas: TZReadOnlyQuery;
    ds_pernoctas: TDataSource;
    qry_TripulacionsTipoObra: TStringField;
    qry_TripulacionsDescripcion: TStringField;
    qry_TripulacionAjuste: TFloatField;
    qry_TripulacionsDescripcionFolio: TStringField;
    qry_TripulacionTotal: TFloatField;
    qry_TripulacionsIdRecurso: TStringField;
    qry_TripulacionsDescripcionRecurso: TStringField;
    chkPersonal: TRadioButton;
    chkEquipo: TRadioButton;
    USqlTripulacion: TZUpdateSQL;
    qry_TripulacionsContrato: TStringField;
    qry_TripulaciondIdFecha: TDateField;
    qry_TripulacioniIdDiario: TIntegerField;
    Label13: TLabel;
    tSolicitado: TEdit;
    qry_TripulaciondSolicitado: TFloatField;
    chkAgrupar: TCheckBox;
    qry_TripulacionTotalAcum: TFloatField;
    qry_TripulacionsNumeroOrden: TStringField;
    Label14: TLabel;
    prgFolios: TAdvSmoothProgressBar;
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
    procedure tsIdFolioKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdFolioEnter(Sender: TObject);
    procedure tsIdFolioExit(Sender: TObject);
    procedure ChkPersonalClick(Sender: TObject);
    procedure ChkEquipoClick(Sender: TObject);
    procedure CargaDatos;
    procedure Solicitado;
    procedure TotalRecurso;
    procedure Recursos;
    procedure CargarDatosAgrupados;
    function TotalRecursoAgrupado(sParamRecurso : string): double;

    procedure qry_TripulacionAfterPost(DataSet: TDataSet);
    procedure chkAgruparClick(Sender: TObject);
    procedure qry_TripulacionCalcFields(DataSet: TDataSet);
    procedure qry_TripulacionBeforeEdit(DataSet: TDataSet);
    procedure grid_tripulacionDblClick(Sender: TObject);
    procedure qry_TripulacionBeforeDelete(DataSet: TDataSet);
    procedure qry_TripulacionBeforeInsert(DataSet: TDataSet);


  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTripulacionAjustes: TfrmTripulacionAjustes;
  sFirma_PEP, sPuesto_PEP, sFirma_Contratista, sPuesto_Contratista: string;
  fechaAntes: tDate;
  utgrid: ticdbgrid;
  local_global_pernocta : string;
implementation

uses frm_tripulacion_diaria;

{$R *.dfm}

procedure TfrmTripulacionAjustes.tdIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
  {if key = #13 then
    tsIdCategoria.SetFocus  }
end;

procedure TfrmTripulacionAjustes.tsIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
   if key = #13 then
      grid_tripulacion.SetFocus   
end;

procedure TfrmTripulacionAjustes.tsIdFolioEnter(Sender: TObject);
begin
    tsIdFolio.Color := global_color_entrada;
end;

procedure TfrmTripulacionAjustes.tsIdFolioExit(Sender: TObject);
begin
    tsIdFolio.Color := global_color_salida;
    grid_tripulacion.Columns[3].FieldName := 'Total';
    grid_tripulacion.Columns[4].ReadOnly  := False;
    CargaDatos;
    TotalRecurso;
    chkAgrupar.Checked := False;
end;

procedure TfrmTripulacionAjustes.tsIdFolioKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key =#13 then
       grid_tripulacion.SetFocus;
end;

procedure TfrmTripulacionAjustes.tsIdTripulacionEnter(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_entrada;
end;

procedure TfrmTripulacionAjustes.tsIdTripulacionExit(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_salida;
  if tsidtripulacion.KeyValue <> null then
    lblTripulacion.Caption := 'Tripulacion ' + tsIdTripulacion.KeyValue;
end;

procedure TfrmTripulacionAjustes.tsIdTripulacionKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       txtCantidad.SetFocus;
end;

procedure TfrmTripulacionAjustes.ttripulacion_nacionalesKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    ttripulacion_extranjeros.SetFocus
end;

procedure TfrmTripulacionAjustes.txtCantidadChange(Sender: TObject);
begin
  tEditChangef(txtCantidad, 'Cantidad');
end;

procedure TfrmTripulacionAjustes.txtCantidadEnter(Sender: TObject);
begin
  txtcantidad.Color := global_color_entrada;
end;

procedure TfrmTripulacionAjustes.txtCantidadExit(Sender: TObject);
begin
  txtcantidad.Color := global_color_salida
end;

procedure TfrmTripulacionAjustes.txtCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTedit(txtCantidad, key) then
    key := #0
  else
    if key =#13 then
       cmdAgregar.SetFocus;
  

end;

procedure TfrmTripulacionAjustes.ttripulacion_extranjerosKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdidFecha.SetFocus
end;

//Inicio LA forma

procedure TfrmTripulacionAjustes.FormShow(Sender: TObject);
begin

  UtGrid := TicdbGrid.create(grid_tripulacion);
  tdIdFecha.Date := Date();
  tdIdFecha.SetFocus;

  if connection.configuracion.FieldByName('sIdEmbarcacion').AsString = '*' then
  begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.clear;
      connection.zCommand.sql.Add('select sIdEmbarcacion, sDescripcion, sTipo from embarcaciones ' +
        'Where sTipo="Principal" order by sDescripcion');
      connection.zCommand.Open;

      local_global_pernocta := connection.zCommand.FieldByName('sIdEmbarcacion').AsString;
  end
  else
      local_global_pernocta := connection.configuracion.FieldByName('sIdEmbarcacion').AsString;

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

  Recursos;

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

  Solicitado;
end;

procedure TfrmTripulacionAjustes.frxTripulacionGetValue(const VarName: string;
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

procedure TfrmTripulacionAjustes.grid_tripulacionDblClick(Sender: TObject);
begin
    if chkAgrupar.Checked then
    begin
        tsIdCategoria.KeyValue := qry_tripulacion.FieldByName('sIdRecurso').AsString;
        tsIdCategoria.OnExit(sender);
    end;
end;

procedure TfrmTripulacionAjustes.grid_tripulacionTitleClick(Column: TColumn);
begin
//if grid_tripulacion.datasource.DataSet.IsEmpty=false  then
  if grid_tripulacion.DataSource.DataSet.RecordCount>0 then
   UtGrid.DbGridTitleClick(Column);
end;

//Consultar la Fecha

procedure TfrmTripulacionAjustes.tdIdFechaExit(Sender: TObject);
begin
  tdidfecha.Color := global_color_salida;
  grid_tripulacion.Columns[3].FieldName := 'Total';
  grid_tripulacion.Columns[4].ReadOnly  := False;

  fechaAntes := tdIdFecha.Date;

  Tripulacion.Active := False;
  Tripulacion.SQL.Clear;
  Tripulacion.SQL.Add('select * from tripulacion where sContrato = :Contrato and dFechaVigencia =:Fecha order by sIdTripulacion');
  Tripulacion.params.ParamByName('Contrato').DataType := ftString;
  Tripulacion.params.ParamByName('Contrato').Value    := Global_Contrato;
  Tripulacion.params.ParamByName('fecha').DataType    := ftDate;
  Tripulacion.params.ParamByName('fecha').Value       := fechaAntes;
  Tripulacion.Open;

  Recursos;

  Folios.Active := False;
  Folios.ParamByName('contrato').AsString := global_contrato;
  Folios.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  Folios.Open;

  tsIdCategoria.KeyValue := '<Todas>';
  tsIdFolio.KeyValue := '<todos>';

  CargaDatos;
  TotalRecurso;

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

  Solicitado;
  chkAgrupar.Checked := False;
end;


//Consultar la Categoria

procedure TfrmTripulacionAjustes.ActualizaPersonalClick(Sender: TObject);
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

procedure TfrmTripulacionAjustes.btnCancelClick(Sender: TObject);
begin
  close;
end;

procedure TfrmTripulacionAjustes.btnUpdateClick(Sender: TObject);
begin
 if (tsIdCategoria.KeyValue = null) or (tsIdCategoria.KeyValue = '<Todas>') then begin
    ShowMessage('Favor de Seleccionar una Partida');
    exit;
  end;

  if (tsIdFolio.KeyValue = null) or (tsIdFolio.KeyValue = '<todos>') then begin
    ShowMessage('Favor de Seleccionar un Folio');
    exit;
  end;

  qry_tripulacion.Refresh;
end;

procedure TfrmTripulacionAjustes.ChkEquipoClick(Sender: TObject);
begin
     CargaDatos;
     TotalRecurso;
end;

procedure TfrmTripulacionAjustes.ChkPersonalClick(Sender: TObject);
begin
    chkAgrupar.Checked := False;
    grid_tripulacion.Columns[3].FieldName := 'Total';
    grid_tripulacion.Columns[4].ReadOnly  := False;
    Recursos;
    CargaDatos;
    solicitado;
    TotalRecurso;
end;

procedure TfrmTripulacionAjustes.tGuardarClick(Sender: TObject);
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

procedure TfrmTripulacionAjustes.qryTripulacionPernoctaFueraAfterScroll(
  DataSet: TDataSet);
begin
(*      qryTripulacionPernoctaFuera.FieldValues['sContrato'] := global_contrato;
      qryTripulacionPernoctaFuera.FieldValues['sIdTurno'] := QryTurnos.FieldValues['sIdTurno'];
      qryTripulacionPernoctaFuera.FieldValues['dIdFecha'] := tdIdFecha.Date;*)
end;

procedure TfrmTripulacionAjustes.qry_tripulacionAfterInsert(
  DataSet: TDataSet);
begin
  qry_tripulacion.Cancel
end;

procedure TfrmTripulacionAjustes.qry_TripulacionAfterPost(DataSet: TDataSet);
begin
   TotalRecurso;

   if chkPersonal.Checked then
   begin
       //consultamos los movimeintos de barco que tengan el folio asignado.
       connection.zCommand.Active := False;
       connection.zCommand.SQL.Clear;
       connection.zCommand.SQL.Add('select sContrato, dIdFecha, iIdDiario from movimientosxfolios where sContrato =:Contrato and dIdFecha =:Fecha and sFolio =:folio ');
       connection.zCommand.ParamByName('Contrato').AsString := global_contrato_barco;
       connection.zCommand.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
       connection.zCommand.ParamByName('folio').AsString    := qry_tripulacion.FieldByName('sNumeroOrden').AsString;
       connection.zCommand.Open;

       if connection.zCommand.RecordCount > 0 then
       begin
           prgFolios.Visible := True;
           label14.Visible   := True;
           prgFolios.Maximum := connection.zCommand.RecordCount;
           prgFolios.Position := 0;
           prgFolios.Refresh;
           label14.Refresh;
           while not connection.zCommand.Eof do
           begin
               TdProrrateoFolio(connection.zCommand.FieldByName('sContrato').AsString,tdIdFecha.Date,connection.zCommand.FieldByName('iIdDiario').AsInteger);
               connection.zCommand.Next;
               prgFolios.Position := prgFolios.Position + 1;
               prgFolios.Refresh;
               label14.Refresh;
           end;
           prgFolios.Visible := False;
           label14.Visible   := False;
       end;
   end;
end;

procedure TfrmTripulacionAjustes.qry_TripulacionBeforeDelete(DataSet: TDataSet);
begin
  If ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
  begin
      MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
      abort;
  end;
end;

procedure TfrmTripulacionAjustes.qry_TripulacionBeforeEdit(DataSet: TDataSet);
begin
   If ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
   begin
      MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
      abort;
   end;

   if chkAgrupar.Checked then
   begin
        messageDLG('No se puede editar en modo <Agrupados Todos los Folios> ', mtInformation, [mbOk], 0);
        qry_tripulacion.Cancel;
   end;
end;

procedure TfrmTripulacionAjustes.qry_TripulacionBeforeInsert(DataSet: TDataSet);
begin
  If ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
  begin
      MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
      abort;
  end;
end;

procedure TfrmTripulacionAjustes.qry_TripulacionBeforePost(DataSet: TDataSet);
begin
  if qry_Tripulacion.FieldValues['dCantidd'] < 0 then
    qry_Tripulacion.cancel;

end;

procedure TfrmTripulacionAjustes.qry_TripulacionCalcFields(DataSet: TDataSet);
begin
   if chkAgrupar.Checked then
   begin
       if qry_tripulacion.RecordCount > 0 then
       begin
           qry_tripulacionTotalAcum.Value := TotalRecursoAgrupado(qry_tripulacion.FieldByName('sIdRecurso').AsString);
       end;
   end;
end;

procedure TfrmTripulacionAjustes.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  utgrid.Destroy;
  action := cafree;
end;

procedure TfrmTripulacionAjustes.tAceptarClick(Sender: TObject);
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

procedure TfrmTripulacionAjustes.tAgregarClick(Sender: TObject);
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

procedure TfrmTripulacionAjustes.tCancelarClick(Sender: TObject);
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

procedure TfrmTripulacionAjustes.tdCategoriaEnter(Sender: TObject);
begin
   tdCategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionAjustes.tdCategoriaExit(Sender: TObject);
begin
   tdCategoria.Color := global_color_salida;
end;

procedure TfrmTripulacionAjustes.tdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tAceptar.SetFocus
end;

procedure TfrmTripulacionAjustes.tdIdCategoriaEnter(Sender: TObject);
begin
   tdIdCategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionAjustes.tdIdCategoriaExit(Sender: TObject);
begin
   tdIdCategoria.Color := global_color_salida;
end;

procedure TfrmTripulacionAjustes.tdIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tdcategoria.SetFocus;
end;

procedure TfrmTripulacionAjustes.tdIdFechaEnter(Sender: TObject);
begin
  tdIdFecha.Color := global_Color_entrada
end;


procedure TfrmTripulacionAjustes.ripulacinDiariaDiaAnterior1Click(
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

procedure TfrmTripulacionAjustes.btnDeleteClick(Sender: TObject);
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

procedure TfrmTripulacionAjustes.btnNuevoClick(Sender: TObject);
begin
  if (tsIdCategoria.KeyValue = null) or (tsIdCategoria.KeyValue = '<Todas>') then begin
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


procedure TfrmTripulacionAjustes.btnPrinterClick(Sender: TObject);
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

procedure TfrmTripulacionAjustes.tsIdCategoriaEnter(Sender: TObject);
begin
  tsidcategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionAjustes.tsIdCategoriaExit(Sender: TObject);
begin
    tsIdcategoria.Color := global_Color_salida;
    grid_tripulacion.Columns[3].FieldName := 'Total';
    grid_tripulacion.Columns[4].ReadOnly  := False;
    CargaDatos;
    TotalRecurso;
    chkAgrupar.Checked := False;
end;

 //SOAD -> Busqued de tripulacion por id.

procedure TfrmTripulacionAjustes.cmbTurnosExit(Sender: TObject);
begin
  tdIdFecha.OnExit(sender);
end;

procedure TfrmTripulacionAjustes.cmdAgregarClick(Sender: TObject);
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

procedure TfrmTripulacionAjustes.cmdSalirClick(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
  cmdSalir.Visible := False;    
end;

procedure TfrmTripulacionAjustes.Panel2Click(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
end;

procedure TfrmTripulacionAjustes.PopupMenu1Popup(Sender: TObject);
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

procedure TfrmTripulacionAjustes.CargaDatos;
begin
    qry_tripulacion.Active := False;
    qry_tripulacion.SQL.Clear;
    //>>Personal<<
    if chkPersonal.Checked then
       qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdPersonal as sIdRecurso, bp.sDescripcion, '+
                      'concat(bp.sIdPersonal, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, '+
                      'o.sIdFolio as sDescripcionFolio, '+
                      'xround(sum(bp.dCantHH),2) as Total, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste, mr.dCantidad as dSolicitado '+
                      'FROM bitacoradepersonal bp '+
                      'inner join bitacoradeactividades ba '+
                      'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                      'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                      'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) '+
                      'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                      'inner join moerecursos mr on (mr.sIdRecurso = bp.sIdPersonal and mr.iIdMoe = ( select m.iIdMoe '+
                      'from moe m '+
                      'where m.sContrato = bp.sContrato '+
                      'and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                      '                    from moe m1 '+
                      '                    where m1.sContrato = bp.sContrato '+
                      '                    and m1.dIdFecha <= bp.dIdFecha '+
                      '                  ) '+
                      ')                           '+
                      'and eTipoRecurso = "Personal") '+
                      'inner join personal e on (e.sContrato =:ContratoBarco and e.sIdPersonal = bp.sIdPersonal) '+
                      'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.didfecha =:fecha '+
                      'AND bp.sidPernocta =:pernocta '+
                      'and bp.sIdPersonal like :categoria GROUP BY bp.sNumeroOrden, bp.sIdPersonal, bp.sTipoObra order by bp.sNumeroOrden, e.iItemOrden  ')
    else
    //>>Equipo<<
       qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdEquipo as sIdRecurso, bp.sDescripcion, '+
                      'concat(bp.sIdEquipo, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, '+
                      'o.sIdFolio as sDescripcionFolio, '+
                      'xround(sum(bp.dCantHH),2) as Total, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste, mr.dCantidad as dSolicitado '+
                      'FROM bitacoradeequipos bp '+
                      'inner join bitacoradeactividades ba '+
                      'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                      'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                      'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) '+
                      'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                      'inner join moerecursos mr on (mr.sIdRecurso = bp.sIdEquipo and mr.iIdMoe = ( select m.iIdMoe '+
                      'from moe m '+
                      'where m.sContrato = bp.sContrato '+
                      'and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                      '                    from moe m1 '+
                      '                    where m1.sContrato = bp.sContrato '+
                      '                    and m1.dIdFecha <= bp.dIdFecha '+
                      '                  ) '+
                      ')                           '+
                      'and eTipoRecurso = "Equipo") '+
                      'inner join equipos e on (e.sContrato =:ContratoBarco and e.sIdEquipo = bp.sIdEquipo) '+
                      'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.didfecha =:fecha '+
                      'AND bp.sidPernocta =:pernocta '+
                      'and bp.sIdEquipo like :categoria GROUP BY bp.sNumeroOrden, bp.sIdEquipo, bp.sTipoObra order by bp.sNumeroOrden, e.iItemOrden ');

    qry_tripulacion.params.ParamByName('contratoBarco').asString  := global_contrato_barco;
    qry_tripulacion.params.ParamByName('contrato').asString      := global_contrato;
    qry_tripulacion.params.ParamByName('fecha').AsDate           := tdIdFecha.Date;
    //--Combos----- <<
    qry_tripulacion.params.ParamByName('categoria').DataType := ftString;
    if tsIdCategoria.KeyValue = '<Todas>' then
       qry_tripulacion.params.ParamByName('categoria').Value := '%'
    else
       qry_tripulacion.params.ParamByName('categoria').Value := tsIdCategoria.KeyValue;

    qry_tripulacion.params.ParamByName('folio').DataType := ftString;
    if tsIdFolio.KeyValue = '<todos>' then
       qry_tripulacion.params.ParamByName('folio').Value := '%'
    else
       qry_tripulacion.params.ParamByName('folio').Value := tsIdFolio.KeyValue;
    //--------------- <<
    qry_tripulacion.params.ParamByName('pernocta').asString  := local_global_pernocta;
    qry_tripulacion.Open;

    if qry_Tripulacion.RecordCount = 0 then
    begin
      ttripulacion_nacionales.Text := '0';
      ttripulacion_extranjeros.Text := '0';
      qry_Tripulacion.Close;
    end;

    UsqlTripulacion.ModifySQL.Clear;
    if chkPersonal.Checked then
       UsqlTripulacion.ModifySQL.Add('update bitacoradepersonal set dAjuste =:Ajuste where sContrato = :sContrato and dIdFecha = :dIdFecha ' +
                       'and iIdDiario = :iIdDiario and sIdPersonal =:sIdRecurso and sTipoObra =:sTipoObra ')
    else
       UsqlTripulacion.ModifySQL.Add('update bitacoradeequipos set dAjuste =:Ajuste where sContrato = :sContrato and dIdFecha = :dIdFecha ' +
                       'and iIdDiario = :iIdDiario and sIdEquipo =:sIdRecurso and sTipoObra =:sTipoObra ');

end;

procedure TfrmTripulacionAjustes.chkAgruparClick(Sender: TObject);
begin
    if chkAgrupar.Checked then
    begin
       tsIdCategoria.KeyValue := '<Todas>';
       CargarDatosAgrupados;
       grid_tripulacion.Columns[3].FieldName := 'TotalAcum';
       grid_tripulacion.Columns[4].ReadOnly  := True;
    end
    else
    begin
       CargaDatos;
       grid_tripulacion.Columns[3].FieldName := 'Total';
       grid_tripulacion.Columns[4].ReadOnly  := False;
    end;
end;

procedure TfrmTripulacionAjustes.Solicitado;
begin
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select sum(dCantidad) as Cantidad from moerecursos '+
                        'where iIdMoe = '+
                        '( select m.iIdMoe '+
                        'from moe m          '+
                        'where m.sContrato = :Orden '+
                        'and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                        '                    from moe m1 '+
                        '                    where m1.sContrato = :Orden '+
                        '                    and m1.dIdFecha <=:fecha '+
                        '                  )) ');
     if chkPersonal.Checked then
        connection.zCommand.SQL.Add(' and eTipoRecurso = "Personal"');
     if chkEquipo.Checked then
        connection.zCommand.SQL.Add(' and eTipoRecurso = "Equipo"');
     connection.zCommand.ParamByName('Orden').AsString := global_contrato;
     connection.zCommand.ParamByName('fecha').AsDate   := tdIdFecha.Date;
     connection.zCommand.Open;

     if connection.zCommand.RecordCount > 0 then
        tSolicitado.Text := FloatToStr(connection.zCommand.FieldByName('Cantidad').AsFloat);
end;

procedure TfrmTripulacionAjustes.TotalRecurso;
var
   dTotal : double;
begin
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    if chkEquipo.Checked then
    begin
        connection.zCommand.SQL.Add('SELECT '+
                  'xround(sum(bp.dCantHH),2) as Total, '+
                  'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                  'FROM bitacoradeequipos bp '+
                  'WHERE bp.scontrato =:Orden AND bp.sNumeroOrden like :Folio AND bp.didfecha =:Fecha and bp.sIdEquipo like :Id '+
                  'GROUP BY bp.sNumeroOrden, bp.sIdEquipo, bp.sTipoObra');
    end;
    if chkPersonal.Checked then
    begin
        connection.zCommand.SQL.Add('SELECT '+
                  'xround(sum(bp.dCantHH),2) as Total, '+
                  'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                  'FROM bitacoradepersonal bp '+
                  'WHERE bp.scontrato =:Orden AND bp.sNumeroOrden like :Folio AND bp.didfecha =:Fecha and bp.sIdPersonal like :Id '+
                  'GROUP BY bp.sNumeroOrden, bp.sIdPersonal, bp.sTipoObra');
    end;
    connection.zCommand.ParamByName('Orden').AsString := global_contrato;
    if tsIdCategoria.KeyValue = '<Todas>' then
       connection.zCommand.ParamByName('Id').AsString := '%'
    else
       connection.zCommand.ParamByName('Id').AsString := tsIdCategoria.KeyValue;
    if tsIdFolio.KeyValue = '<todos>' then
       connection.zCommand.ParamByName('Folio').AsString := '%'
    else
       connection.zCommand.ParamByName('Folio').AsString := tsIdFolio.KeyValue;
    connection.zCommand.ParamByName('fecha').AsDate   := tdIdFecha.Date;
    connection.zCommand.Open;

    dTotal := 0;
    while not connection.zCommand.Eof do
    begin
       dTotal := dTotal + (connection.zCommand.FieldByName('Total').AsFloat + connection.zCommand.FieldByName('Ajuste').AsFloat);
       connection.zCommand.Next;
    end;
    ttripulacion_nacionales.Text := FloatToStr(dTotal);
end;

procedure TfrmTripulacionAjustes.Recursos;
begin
  Categorias.Active := False;
  Categorias.SQL.Clear;
  if chkPersonal.Checked then
     Categorias.SQL.Add('select "<Todas>" as sIdRecurso, "" as sTipoObra from bitacoradepersonal '+
              'union '+
              'select bp.sIdPersonal as sIdRecurso, bp.sTipoObra from bitacoradepersonal bp '+
              'inner join personal p on (p.sContrato =:ContratoBarco and p.sIdPersonal  = bp.sIdPersonal) '+
              'where bp.sContrato =:Contrato and bp.dIdFecha =:Fecha '+
              'group by bp.sIdPersonal ')
  else
     Categorias.SQL.Add('select "<Todas>" as sIdRecurso, "" as sTipoObra from bitacoradeequipos '+
              'union '+
              'select bp.sIdEquipo as sIdRecurso, bp.sTipoObra from bitacoradeequipos bp '+
              'inner join equipos p on (p.sContrato =:ContratoBarco and p.sIdEquipo  = bp.sIdEquipo) '+
              'where bp.sContrato =:contrato and bp.dIdFecha =:Fecha '+
              'group by bp.sIdEquipo ');
  Categorias.ParamByName('ContratoBarco').AsString := global_contrato_barco;
  Categorias.ParamByName('Contrato').AsString      := global_contrato;
  Categorias.ParamByName('fecha').AsDate           := fechaAntes;
  Categorias.Open;
end;

procedure TfrmTripulacionAjustes.CargarDatosAgrupados;
begin
    qry_tripulacion.Active := False;
    qry_tripulacion.SQL.Clear;
    //>>Personal<<
    if chkPersonal.Checked then
       qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdPersonal as sIdRecurso, bp.sDescripcion, '+
                      'concat(bp.sIdPersonal, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, '+
                      '"<Agrupados Todos los Folios>" as sDescripcionFolio, '+
                      'round(sum(bp.dCantHH),2) as Total, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste, mr.dCantidad as dSolicitado '+
                      'FROM bitacoradepersonal bp '+
                      'inner join bitacoradeactividades ba '+
                      'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                      'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                      'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) '+
                      'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                      'inner join moerecursos mr on (mr.sIdRecurso = bp.sIdPersonal and mr.iIdMoe = ( select m.iIdMoe '+
                      'from moe m '+
                      'where m.sContrato = bp.sContrato '+
                      'and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                      '                    from moe m1 '+
                      '                    where m1.sContrato = bp.sContrato '+
                      '                    and m1.dIdFecha <= bp.dIdFecha '+
                      '                  ) '+
                      ')                           '+
                      'and eTipoRecurso = "Personal") '+
                      'inner join personal e on (e.sContrato =:ContratoBarco and e.sIdPersonal = bp.sIdPersonal) '+
                      'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.didfecha =:fecha '+
                      'AND bp.sidPernocta =:pernocta '+
                      'and bp.sIdPersonal like :categoria GROUP BY bp.sIdPersonal, bp.sTipoObra order by e.iItemOrden  ')
    else
    //>>Equipo<<
       qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdEquipo as sIdRecurso, bp.sDescripcion, '+
                      'concat(bp.sIdEquipo, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, '+
                      '"<Agrupados Todos los Folios>" as sDescripcionFolio, '+
                      'round(sum(bp.dCantHH),2) as Total, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste, mr.dCantidad as dSolicitado '+
                      'FROM bitacoradeequipos bp '+
                      'inner join bitacoradeactividades ba '+
                      'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                      'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                      'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) '+
                      'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                      'inner join moerecursos mr on (mr.sIdRecurso = bp.sIdEquipo and mr.iIdMoe = ( select m.iIdMoe '+
                      'from moe m '+
                      'where m.sContrato = bp.sContrato '+
                      'and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                      '                    from moe m1 '+
                      '                    where m1.sContrato = bp.sContrato '+
                      '                    and m1.dIdFecha <= bp.dIdFecha '+
                      '                  ) '+
                      ')                           '+
                      'and eTipoRecurso = "Equipo") '+
                      'inner join equipos e on (e.sContrato =:ContratoBarco and e.sIdEquipo = bp.sIdEquipo) '+
                      'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.didfecha =:fecha '+
                      'AND bp.sidPernocta =:pernocta '+
                      'and bp.sIdEquipo like :categoria GROUP BY bp.sIdEquipo, bp.sTipoObra order by e.iItemOrden ');

    qry_tripulacion.params.ParamByName('contratoBarco').asString  := global_contrato_barco;
    qry_tripulacion.params.ParamByName('contrato').asString      := global_contrato;
    qry_tripulacion.params.ParamByName('fecha').AsDate           := tdIdFecha.Date;
    //--Combos----- <<
    qry_tripulacion.params.ParamByName('categoria').DataType := ftString;
    if tsIdCategoria.KeyValue = '<Todas>' then
       qry_tripulacion.params.ParamByName('categoria').Value := '%'
    else
       qry_tripulacion.params.ParamByName('categoria').Value := tsIdCategoria.KeyValue;

    qry_tripulacion.params.ParamByName('folio').DataType := ftString;
    if tsIdFolio.KeyValue = '<todos>' then
       qry_tripulacion.params.ParamByName('folio').Value := '%'
    else
       qry_tripulacion.params.ParamByName('folio').Value := tsIdFolio.KeyValue;
    //--------------- <<
    qry_tripulacion.params.ParamByName('pernocta').asString  := local_global_pernocta;
    qry_tripulacion.Open;

    if qry_Tripulacion.RecordCount = 0 then
    begin
      ttripulacion_nacionales.Text := '0';
      ttripulacion_extranjeros.Text := '0';
      qry_Tripulacion.Close;
    end;

    UsqlTripulacion.ModifySQL.Clear;

end;

function TfrmTripulacionAjustes.TotalRecursoAgrupado(sParamRecurso : string): double;
var
   dTotal : double;
begin
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    if chkEquipo.Checked then
    begin
        connection.zCommand.SQL.Add('SELECT '+
                  'round(sum(bp.dCantHH),2) as Total, '+
                  'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                  'FROM bitacoradeequipos bp '+
                  'WHERE bp.scontrato =:Orden AND bp.didfecha =:Fecha and bp.sIdEquipo = :Id '+
                  'GROUP BY bp.sNumeroOrden, bp.sIdEquipo, bp.sTipoObra ');
    end;
    if chkPersonal.Checked then
    begin
        connection.zCommand.SQL.Add('SELECT '+
                  'round(sum(bp.dCantHH),2) as Total, '+
                  'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                  'FROM bitacoradepersonal bp '+
                  'WHERE bp.scontrato =:Orden AND bp.didfecha =:Fecha and bp.sIdPersonal = :Id '+
                  'GROUP BY bp.sNumeroOrden, bp.sIdPersonal, bp.sTipoObra');
    end;
    connection.zCommand.ParamByName('Orden').AsString := global_contrato;
    connection.zCommand.ParamByName('Id').AsString    := sParamRecurso;
    connection.zCommand.ParamByName('fecha').AsDate   := tdIdFecha.Date;
    connection.zCommand.Open;

    dTotal := 0;
    while not connection.zCommand.Eof do
    begin
       dTotal := dTotal + (connection.zCommand.FieldByName('Total').AsFloat + connection.zCommand.FieldByName('Ajuste').AsFloat);
       connection.zCommand.Next;
    end;
    result := dTotal;
end;

end.

