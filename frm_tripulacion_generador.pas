unit frm_tripulacion_generador;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms,
  Dialogs, StdCtrls, Mask, DBCtrls, ComCtrls, Grids, DBGrids, global, frm_connection,
  DB, ADODB, Buttons, ExtCtrls, frxClass, frxDBSet, ZAbstractRODataset, dateUtils,
  ZDataset, ZAbstractDataset, Controls, Menus, UnitExcepciones, udbgrid, UFunctionsGHH,
  DBDateTimePicker, UnitValidacion, rxToolEdit, rxCurrEdit, RXDBCtrl, UnitTarifa,
  cxGraphics, cxLookAndFeels, cxLookAndFeelPainters, dxSkinsCore,
  dxSkinDevExpressStyle, dxSkinFoggy, cxButtons, ZSqlUpdate;

type
  TfrmTripulacionGenerador = class(TForm)
    grid_tripulacion: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    tsIdCategoria: TDBLookupComboBox;
    ds_tripulacion: TDataSource;
    btnUpdate: TBitBtn;
    Panel1: TPanel;
    ttripulacion_nacionales: TEdit;
    ttripulacion_extranjeros: TEdit;
    frxGenerador: TfrxReport;
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
    qry_TripulacionTotalAcum: TFloatField;
    chkAgrupar: TCheckBox;
    chkAgruparFolios: TCheckBox;
    tdIdFechaTermino: TDBDateTimePicker;
    qry_TripulacionTotalSolicitado: TFloatField;
    qry_TripulacionsNumeroOrden: TStringField;
    chkBarco: TRadioButton;
    chkPernocta: TRadioButton;
    chkAnexoC6: TRadioButton;
    cxImprimir: TcxButton;
    qry_TripulacionsMedida: TStringField;
    chkAnexoC7: TRadioButton;
    chkAnexoC8: TRadioButton;
    chkContinuo: TCheckBox;
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
    procedure frxGeneradorGetValue(const VarName: string; var Value: Variant);
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
    function TotalRecursoAgrupado(sParamRecurso, sParamFolio : string; dParamFechaI, dParamFechaF :tDate): double;
    function TotalSolicitadoAgrupado(sParamRecurso, sParamFolio : string; dParamFechaI, dParamFechaF :tDate): double;
    procedure CargarDatosAgrupadosFolio;
    function TotalRecursoAgrupadoFolio(sParamRecurso, sParamFolio : string; dParamFechaI, dParamFechaF :tDate): double;
    function TotalSolicitadoAgrupadoFolio(sParamRecurso, sParamFolio : string; dParamFechaI, dParamFechaF :tDate): double;

    procedure chkAgruparClick(Sender: TObject);
    procedure qry_TripulacionCalcFields(DataSet: TDataSet);
    procedure qry_TripulacionBeforeEdit(DataSet: TDataSet);
    procedure grid_tripulacionDblClick(Sender: TObject);
    procedure qry_TripulacionAfterOpen(DataSet: TDataSet);
    procedure chkAgruparFoliosClick(Sender: TObject);
    procedure cxImprimirClick(Sender: TObject);


  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTripulacionGenerador: TfrmTripulacionGenerador;
  sFirma_PEP, sPuesto_PEP, sFirma_Contratista, sPuesto_Contratista: string;
  fechaAntes: tDate;
  utgrid: ticdbgrid;
  local_global_pernocta, local_tipo : string;
implementation

uses frm_tripulacion_diaria;

{$R *.dfm}

procedure TfrmTripulacionGenerador.tdIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tdIdFechaTermino.SetFocus
end;

procedure TfrmTripulacionGenerador.tsIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
   if key = #13 then
      grid_tripulacion.SetFocus   
end;

procedure TfrmTripulacionGenerador.tsIdFolioEnter(Sender: TObject);
begin
    tsIdFolio.Color := global_color_entrada;
end;

procedure TfrmTripulacionGenerador.tsIdFolioExit(Sender: TObject);
begin
    tsIdFolio.Color := global_color_salida;
    CargaDatos;
    chkAgrupar.Checked := False;
end;

procedure TfrmTripulacionGenerador.tsIdFolioKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key =#13 then
       grid_tripulacion.SetFocus;
end;

procedure TfrmTripulacionGenerador.tsIdTripulacionEnter(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_entrada;
end;

procedure TfrmTripulacionGenerador.tsIdTripulacionExit(Sender: TObject);
begin
  tsidtripulacion.Color := global_color_salida;
  if tsidtripulacion.KeyValue <> null then
    lblTripulacion.Caption := 'Tripulacion ' + tsIdTripulacion.KeyValue;
end;

procedure TfrmTripulacionGenerador.tsIdTripulacionKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       txtCantidad.SetFocus;
end;

procedure TfrmTripulacionGenerador.ttripulacion_nacionalesKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    ttripulacion_extranjeros.SetFocus
end;

procedure TfrmTripulacionGenerador.txtCantidadChange(Sender: TObject);
begin
  tEditChangef(txtCantidad, 'Cantidad');
end;

procedure TfrmTripulacionGenerador.txtCantidadEnter(Sender: TObject);
begin
  txtcantidad.Color := global_color_entrada;
end;

procedure TfrmTripulacionGenerador.txtCantidadExit(Sender: TObject);
begin
  txtcantidad.Color := global_color_salida
end;

procedure TfrmTripulacionGenerador.txtCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTedit(txtCantidad, key) then
    key := #0
  else
    if key =#13 then
       cmdAgregar.SetFocus;
  

end;

procedure TfrmTripulacionGenerador.ttripulacion_extranjerosKeyPress(
  Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdidFecha.SetFocus
end;

//Inicio LA forma

procedure TfrmTripulacionGenerador.FormShow(Sender: TObject);
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
  tdIdFechaTermino.Date := tdIdFecha.Date;

  Solicitado;
end;

procedure TfrmTripulacionGenerador.frxGeneradorGetValue(const VarName: string;
  var Value: Variant);
begin
  if CompareText(VarName, 'FECHA_INICIO') = 0 then
    Value := tdIdFecha.Date;

  if CompareText(VarName, 'FECHA_FINAL') = 0 then
    Value := tdIdFechaTermino.Date;

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

 // if CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
 //   Value := sPuestoSuperIntendente;

 // if CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
 //   Value := sPuestoSupervisor;

//  if CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
//    Value := sPuestoSupervisorTierra;


  if CompareText( VarName,'PUESTO_SUPERINTENDENTE' ) = 0 then
       if pos('#', sPuestoSuperIntendente) > 0 then
             Value := copy(sPuestoSuperIntendente,0, pos('#', sPuestoSuperIntendente)-1) +#13+ copy(sPuestoSuperIntendente,pos('#', sPuestoSuperIntendente)+1, length(sPuestoSuperIntendente))
          else
             Value := sPuestoSuperIntendente;

  //  Value := sPuestoSuperIntendente ;

  if CompareText( VarName,'PUESTO_SUPERVISOR' ) = 0 then
    if pos('#', sPuestoSupervisorGenerador) > 0 then
             Value := copy(sPuestoSupervisorGenerador,0, pos('#', sPuestoSupervisorGenerador)-1) +#13+ copy(sPuestoSupervisorGenerador,pos('#', sPuestoSupervisorGenerador)+1, length(sPuestoSupervisorGenerador))
          else
             Value := sPuestoSupervisorGenerador;


  //  Value := sPuestoSupervisorGenerador  ;
  if CompareText( VarName,'PUESTO_SUPERVISOR_TIERRA' ) = 0 then
    if pos('#', sPuestoSupervisorTierra) > 0 then
             Value := copy(sPuestoSupervisorTierra,0, pos('#', sPuestoSupervisorTierra)-1) +#13+ copy(sPuestoSupervisorTierra,pos('#', sPuestoSupervisorTierra)+1, length(sPuestoSupervisorTierra))
          else
             Value := sPuestoSupervisorTierra;
 //   Value := sPuestoSupervisorTierra  ;

end;

procedure TfrmTripulacionGenerador.grid_tripulacionDblClick(Sender: TObject);
begin
    if chkAgrupar.Checked then
    begin
        tsIdCategoria.KeyValue := qry_tripulacion.FieldByName('sIdRecurso').AsString;
        tsIdCategoria.OnExit(sender);
    end;
end;

procedure TfrmTripulacionGenerador.grid_tripulacionTitleClick(Column: TColumn);
begin
//if grid_tripulacion.datasource.DataSet.IsEmpty=false  then
  if grid_tripulacion.DataSource.DataSet.RecordCount>0 then
   UtGrid.DbGridTitleClick(Column);
end;

//Consultar la Fecha

procedure TfrmTripulacionGenerador.tdIdFechaExit(Sender: TObject);
begin
  tdidfecha.Color := global_color_salida;

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
  grid_tripulacion.Columns[0].FieldName := 'dIdFecha';
end;


//Consultar la Categoria

procedure TfrmTripulacionGenerador.ActualizaPersonalClick(Sender: TObject);
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

procedure TfrmTripulacionGenerador.btnCancelClick(Sender: TObject);
begin
  close;
end;

procedure TfrmTripulacionGenerador.btnUpdateClick(Sender: TObject);
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

procedure TfrmTripulacionGenerador.ChkEquipoClick(Sender: TObject);
begin
     CargaDatos;
end;

procedure TfrmTripulacionGenerador.ChkPersonalClick(Sender: TObject);
begin
    //label13.Visible     := True;
    //tsolicitado.Visible := True;
    chkAgrupar.Checked := False;

    Recursos;
    CargaDatos;
    if chkBarco.Checked = False then
       solicitado ;

    //if (chkBarco.Checked ) or (chkPernocta.Checked) or (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
    //begin
    //    label13.Visible := False;
    //    tsolicitado.Visible := False;
    //end;

end;

procedure TfrmTripulacionGenerador.tGuardarClick(Sender: TObject);
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

procedure TfrmTripulacionGenerador.qryTripulacionPernoctaFueraAfterScroll(
  DataSet: TDataSet);
begin
(*      qryTripulacionPernoctaFuera.FieldValues['sContrato'] := global_contrato;
      qryTripulacionPernoctaFuera.FieldValues['sIdTurno'] := QryTurnos.FieldValues['sIdTurno'];
      qryTripulacionPernoctaFuera.FieldValues['dIdFecha'] := tdIdFecha.Date;*)
end;

procedure TfrmTripulacionGenerador.qry_tripulacionAfterInsert(
  DataSet: TDataSet);
begin
  qry_tripulacion.Cancel
end;

procedure TfrmTripulacionGenerador.qry_TripulacionAfterOpen(DataSet: TDataSet);
begin
    TotalRecurso ;
end;

procedure TfrmTripulacionGenerador.qry_TripulacionBeforeEdit(DataSet: TDataSet);
begin
    if chkAgrupar.Checked then
    begin
        messageDLG('No se puede editar en modo <Agrupados Todos los Folios> ', mtInformation, [mbOk], 0);
        qry_tripulacion.Cancel;
    end;
end;

procedure TfrmTripulacionGenerador.qry_TripulacionBeforePost(DataSet: TDataSet);
begin
  if qry_Tripulacion.FieldValues['dCantidd'] < 0 then
    qry_Tripulacion.cancel;

end;

procedure TfrmTripulacionGenerador.qry_TripulacionCalcFields(DataSet: TDataSet);
begin
    if qry_tripulacion.RecordCount > 0 then
    begin
         if chkAgrupar.Checked then
         begin
             qry_tripulacionTotalAcum.Value       := TotalRecursoAgrupado(qry_tripulacion.FieldByName('sIdRecurso').AsString, '%', tdIdFecha.Date, tdIdFechaTermino.Date );
             qry_tripulacionTotalSolicitado.Value := TotalSolicitadoAgrupado(qry_tripulacion.FieldByName('sIdRecurso').AsString, '%', tdIdFecha.Date, tdIdFechaTermino.Date );
         end;

         if chkAgruparFolios.Checked then
         begin
             qry_tripulacionTotalAcum.Value       := TotalRecursoAgrupadoFolio(qry_tripulacion.FieldByName('sIdRecurso').AsString, qry_tripulacion.FieldByName('sNumeroOrden').AsString, tdIdFecha.Date, tdIdFechaTermino.Date );
             qry_tripulacionTotalSolicitado.Value := TotalSolicitadoAgrupadoFolio(qry_tripulacion.FieldByName('sIdRecurso').AsString, qry_tripulacion.FieldByName('sNumeroOrden').AsString, tdIdFecha.Date, tdIdFechaTermino.Date );
         end;
    end;
end;

procedure TfrmTripulacionGenerador.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  utgrid.Destroy;
  action := cafree;
end;

procedure TfrmTripulacionGenerador.tAceptarClick(Sender: TObject);
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

procedure TfrmTripulacionGenerador.tAgregarClick(Sender: TObject);
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

procedure TfrmTripulacionGenerador.tCancelarClick(Sender: TObject);
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

procedure TfrmTripulacionGenerador.tdCategoriaEnter(Sender: TObject);
begin
   tdCategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionGenerador.tdCategoriaExit(Sender: TObject);
begin
   tdCategoria.Color := global_color_salida;
end;

procedure TfrmTripulacionGenerador.tdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tAceptar.SetFocus
end;

procedure TfrmTripulacionGenerador.tdIdCategoriaEnter(Sender: TObject);
begin
   tdIdCategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionGenerador.tdIdCategoriaExit(Sender: TObject);
begin
   tdIdCategoria.Color := global_color_salida;
end;

procedure TfrmTripulacionGenerador.tdIdCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tdcategoria.SetFocus;
end;

procedure TfrmTripulacionGenerador.tdIdFechaEnter(Sender: TObject);
begin
  tdIdFecha.Color := global_Color_entrada
end;


procedure TfrmTripulacionGenerador.ripulacinDiariaDiaAnterior1Click(
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

procedure TfrmTripulacionGenerador.btnDeleteClick(Sender: TObject);
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

procedure TfrmTripulacionGenerador.btnNuevoClick(Sender: TObject);
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


procedure TfrmTripulacionGenerador.btnPrinterClick(Sender: TObject);
begin
//  try
//      if (qry_Tripulacion.Active) and (qry_Tripulacion.RecordCount > 0) then
//        procreporteTripulacion(Global_Contrato_barco, QryTurnos.FieldValues['sIdTurno'], tdIdFecha.DateTime, frmTripulacionDiaria, frxTripulacion.OnGetValue, connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, 'reporteBarco'), 'Barco')
//      else
//        showmessage('No hay datos para imprimir');
//  except
//    on e: exception do begin
//      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Control de pernoctas', 'Al imprimir', 0);
//    end;
//  end;
end;

procedure TfrmTripulacionGenerador.tsIdCategoriaEnter(Sender: TObject);
begin
  tsidcategoria.Color := global_color_entrada;
end;

procedure TfrmTripulacionGenerador.tsIdCategoriaExit(Sender: TObject);
begin
    tsIdcategoria.Color := global_Color_salida;
    CargaDatos;
    chkAgrupar.Checked := False;
end;

 //SOAD -> Busqued de tripulacion por id.

procedure TfrmTripulacionGenerador.cmbTurnosExit(Sender: TObject);
begin
  tdIdFecha.OnExit(sender);
end;

procedure TfrmTripulacionGenerador.cmdAgregarClick(Sender: TObject);
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

procedure TfrmTripulacionGenerador.cmdSalirClick(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
  cmdSalir.Visible := False;
end;

procedure TfrmTripulacionGenerador.cxImprimirClick(Sender: TObject);
var
   mesI, mesF : string;
begin
    mesI := copy(DateToStr(tdIdFecha.Date),4,2);
    mesF := copy(DateToStr(tdIdFechaTermino.Date),4,2);

    if StrToInt(mesI)<> StrToInt(mesF) then
    begin
        messageDLG('No se puede imprimir un generador que comprenda dos o más meses!', mtInformation, [mbOk], 0);
        exit;
    end;

    if chkBarco.Checked then
       local_tipo := 'Barco';

    if chkPersonal.Checked then
       local_tipo := 'Personal';

    if chkEquipo.Checked then
       local_tipo := 'Equipo';

    if chkPernocta.Checked then
       local_tipo := 'Pernocta';

    if chkAnexoC6.Checked then
       local_tipo := 'C6';

    if chkAnexoC7.Checked then
       local_tipo := 'C7';

    if chkAnexoC8.Checked then
       local_tipo := 'C8';


     try
      if (qry_Tripulacion.Active) and (qry_Tripulacion.RecordCount > 0) then
        procReporteGenerador(Global_Contrato, local_tipo, tsIdCategoria.KeyValue, chkContinuo.Checked, tdIdFecha.Date, tdIdFechaTermino.Date, frmTripulacionGenerador, frxGenerador.OnGetValue, connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, 'reporteBarco'))
      else
        showmessage('No hay datos para imprimir');
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Control de pernoctas', 'Al imprimir', 0);
    end;
  end;
end;

procedure TfrmTripulacionGenerador.Panel2Click(Sender: TObject);
begin
  Panel2.Visible := False;
  cmdAgregar.Visible := False;
end;

procedure TfrmTripulacionGenerador.PopupMenu1Popup(Sender: TObject);
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

procedure TfrmTripulacionGenerador.CargaDatos;
begin
    qry_tripulacion.Active := False;
    qry_tripulacion.SQL.Clear;
    if (chkPersonal.Checked) or (chkequipo.Checked) then
    begin
        //>>Personal<<
        if chkPersonal.Checked then
           qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdPersonal as sIdRecurso, bp.sDescripcion, '+
                          'concat(bp.sIdPersonal, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, e.sMedida, '+
                          'o.sIdFolio as sDescripcionFolio, '+
                          'xround(sum(bp.dCantHH),2) as Total, '+
                          'ifnull(SUM(bp.dAjuste),0) AS Ajuste, mr.dCantidad as dSolicitado '+
                          'FROM bitacoradepersonal bp '+
                         (*'from '               +
                          '( SELECT xround(sum((bp3.dCanthh )),2) as dCanthh,sum(bp3.dAjuste) AS dAjuste,bp3.sContrato, ' + #13#10 +
                             'bp3.dIdFecha,bp3.iIdDiario,bp3.sTipoObra,bp3.sIdPersonal,bp3.sDescripcion,bp3.sNumeroOrden,bp3.iIdTarea,' + #13#10 +
                             'bp3.iIdActividad,bp3.sIdPernocta ' +
                              'FROM bitacoradepersonal bp3 ' + #13#10 +
                              'WHERE bp3.scontrato =:Contrato AND bp3.didfecha >=:FechaI and bp3.dIdFecha <=:FechaF ' + #13#10 + 
                              'and bp3.stipopernocta like :Categoria and bp3.sNumeroOrden like :Folio' + #13#10 +
                              'group by bp3.dIdFecha, bp3.sNumeroOrden, bp3.stipopernocta,bp3.sIdPersonal' + #13#10 + 
                              ') bp ' + #13#10 + *)
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
                          'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.dIdFecha >=:fechaI and bp.dIdFecha <=:fechaF '+
                          'AND bp.sidPernocta =:pernocta '+
                          'and bp.sIdPersonal like :categoria GROUP BY bp.dIdFecha, bp.sNumeroOrden, bp.sIdPersonal, bp.sTipoObra order by bp.dIdFecha, e.iItemOrden, bp.sNumeroOrden ')
        else
        //>>Equipo<<
           qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdEquipo as sIdRecurso, bp.sDescripcion, '+
                          'concat(bp.sIdEquipo, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, e.sMedida, '+
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
                          'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.didfecha >=:fechaI and bp.dIdFecha <=:fechaF '+
                          'AND bp.sidPernocta =:pernocta '+
                          'and bp.sIdEquipo like :categoria GROUP BY bp.dIdFecha, bp.sNumeroOrden, bp.sIdEquipo, bp.sTipoObra order by bp.dIdFecha, e.iItemOrden, bp.sNumeroOrden ');

        qry_tripulacion.params.ParamByName('pernocta').asString  := local_global_pernocta;
    end;

    if chkBarco.Checked then
    begin
        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, tp.sIdTipoMovimiento as sIdRecurso, tp.sDescripcion, '+
                        'concat(tp.sIdTipoMovimiento, " ", tp.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, "" as sMedida, '+
                        'o.sIdFolio as sDescripcionFolio, '+
                        'xround(sum(bp.sFactor),6) as Total, '+
                        'ifnull(SUM(bp.dCantHH),0) AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM movimientosxfolios bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sNumeroOrden and o.sNumeroOrden = bp.sFolio ) '+
                        'inner join movimientosdeembarcacion me on (me.sContrato =bp.sContrato and me.dIdFecha = bp.dIdFecha and me.iIdDiario = bp.iIddiario ) '+
                        'inner join tiposdemovimiento tp on (tp.sContrato = bp.sContrato and tp.sIdTipoMovimiento = me.sClasificacion and tp.sClasificacion = "Movimiento de barco") '+
                        'WHERE bp.scontrato =:ContratoBarco AND bp.sNumeroOrden =:Contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and me.sClasificacion like :categoria and bp.sfolio like :folio '+
                        'group by bp.sFolio, bp.iIddiario order by me.dIdFecha, tp.sIdTipoMovimiento, o.iOrden');

    end;

    if chkPernocta.Checked then
    begin
        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, bp.sIdCuenta as sIdRecurso, c.sDescripcion, '+
                        'concat(bp.sIdCuenta, " ", c.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, c.sMedida, '+
                        'o.sIdFolio as sDescripcionFolio, '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.00000 AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM bitacoradepernocta bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                        'inner join cuentas c on (c.sIdCuenta = bp.sIdCuenta) '+
                        'WHERE bp.scontrato =:Contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and bp.sIdCuenta like :Categoria and bp.sNumeroOrden like :Folio and :ContratoBarco = :ContratoBarco '+
                        'group by bp.dIdFecha, bp.sNumeroOrden, bp.sIdCuenta order by bp.dIdFecha, c.sIdCuenta');

    end;

    if (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
    begin
        if chkAnexoC6.Checked then
           local_tipo := 'C6';

        if chkAnexoC7.Checked then
           local_tipo := 'C7';

        if chkAnexoC8.Checked then
           local_tipo := 'C8';

        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, bp.sIdMaterial as sIdRecurso, substr(ax.mDescripcion,1,250) as sDescripcion, '+
                        'concat(bp.sIdMaterial, " ", bp.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, ax.sMedida, ax.mDescripcion, '+
                        'o.sIdFolio as sDescripcionFolio, '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.00000 AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM bitacorademateriales bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                        'inner join actividadesxanexo ax on (ax.sContrato =:ContratoBarco and ax.sNumeroActividad = bp.sIdMaterial and ax.sTipoActividad = "Actividad") '+
                        'WHERE bp.scontrato =:contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and bp.sIdMaterial like :Categoria and bp.sNumeroOrden like :Folio and ax.sAnexo =:Anexo '+
                        'group by bp.dIdFecha, bp.sNumeroOrden, bp.sIdMaterial order by bp.dIdFecha, ax.iItemOrden');
         qry_tripulacion.params.ParamByName('Anexo').asString  := local_tipo;
    end;


    qry_tripulacion.params.ParamByName('contratoBarco').asString  := global_contrato_barco;
    qry_tripulacion.params.ParamByName('contrato').asString       := global_contrato;
    qry_tripulacion.params.ParamByName('fechaI').AsDate           := tdIdFecha.Date;
    qry_tripulacion.params.ParamByName('fechaF').AsDate           := tdIdFechaTermino.Date;

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

    qry_tripulacion.Open;

    if qry_Tripulacion.RecordCount = 0 then
    begin
      ttripulacion_nacionales.Text := '0';
      ttripulacion_extranjeros.Text := '0';
      qry_Tripulacion.Close;
    end;

    grid_tripulacion.Columns[0].FieldName := 'dIdFecha';

end;

procedure TfrmTripulacionGenerador.chkAgruparClick(Sender: TObject);
begin
    if chkAgrupar.Checked then
    begin
       tsIdCategoria.KeyValue := '<Todas>';
       CargarDatosAgrupados;
       grid_tripulacion.Columns[0].FieldName := 'dFecha';
       grid_tripulacion.Columns[4].FieldName := 'TotalAcum';
       grid_tripulacion.Columns[6].FieldName := 'TotalSolicitado';
    end
    else
    begin
       CargaDatos;
       grid_tripulacion.Columns[0].FieldName := 'dIdFecha';
       grid_tripulacion.Columns[4].FieldName := 'Total';
       grid_tripulacion.Columns[6].FieldName := 'dSolicitado';
    end;
end;

procedure TfrmTripulacionGenerador.chkAgruparFoliosClick(Sender: TObject);
begin
   if chkAgruparFolios.Checked then
    begin
       tsIdCategoria.KeyValue := '<Todas>';
       CargarDatosAgrupadosFolio;
       grid_tripulacion.Columns[0].FieldName := 'dFecha';
       grid_tripulacion.Columns[4].FieldName := 'TotalAcum';
       grid_tripulacion.Columns[6].FieldName := 'TotalSolicitado';
    end
    else
    begin
       CargaDatos;
       grid_tripulacion.Columns[0].FieldName := 'dIdFecha';
       grid_tripulacion.Columns[4].FieldName := 'Total';
       grid_tripulacion.Columns[6].FieldName := 'dSolicitado';
    end;
end;

procedure TfrmTripulacionGenerador.Solicitado;
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
     connection.zCommand.ParamByName('fecha').AsDate   := tdIdFechaTermino.Date;
     connection.zCommand.Open;

     if connection.zCommand.RecordCount > 0 then
        tSolicitado.Text := FloatToStr(connection.zCommand.FieldByName('Cantidad').AsFloat);
end;

procedure TfrmTripulacionGenerador.TotalRecurso;
var
   dTotal, dAjuste : double;
   dFechaActual : tDate;
begin
    dTotal := 0;
    dFechaActual := tdIdFecha.Date;

    while dFechaActual <= tdIdFechaTermino.Date  do
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        if (chkEquipo.Checked) or (chkPersonal.Checked) then
        begin
            if (chkEquipo.Checked) then
            begin
                connection.zCommand.SQL.Add('SELECT '+
                          'round(sum(bp.dCantHH),2) as Total, '+
                          'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                          'FROM bitacoradeequipos bp '+
                          'WHERE bp.scontrato =:Orden AND bp.sNumeroOrden like :Folio AND bp.didfecha =:Fecha and bp.sIdEquipo like :Id '+
                          'GROUP BY bp.sNumeroOrden, bp.sIdEquipo');
            end;
            if chkPersonal.Checked then
            begin
                connection.zCommand.SQL.Add('SELECT '+
                          'round(sum(bp.dCantHH),2) as Total, '+
                          'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                          'FROM bitacoradepersonal bp '+
                          'WHERE bp.scontrato =:Orden AND bp.sNumeroOrden like :Folio AND bp.didfecha =:Fecha and bp.sIdPersonal like :Id '+
                          'GROUP BY bp.sNumeroOrden, bp.sIdPersonal');
            end;
        end;

        if chkBarco.Checked then
        begin
             connection.zCommand.SQL.Add('SELECT '+
                        'round(sum(bp.sFactor),6) as Total, '+
                        'sum(bp.sFactorBarco) AS Ajuste '+
                        'FROM movimientosxfolios bp '+
                        'inner join movimientosdeembarcacion me on (me.sContrato =bp.sContrato and me.dIdFecha = bp.dIdFecha and me.iIdDiario = bp.iIddiario ) '+
                        'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden = :Orden AND bp.didfecha =:Fecha '+
                        'and me.sClasificacion like :Id and bp.sfolio like :Folio '+
                        'group by bp.sContrato, bp.sFolio');
             connection.zCommand.ParamByName('Contrato').AsString := global_contrato_barco;
        end;

        if chkPernocta.Checked then
        begin
            connection.zCommand.SQL.Add('SELECT   '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.0000 AS Ajuste '+
                        'FROM bitacoradepernocta bp '+
                        'WHERE  bp.sContrato = :Orden AND bp.didfecha =:fecha '+
                        'and bp.sIdCuenta like :Id and bp.sNumeroOrden like :folio '+
                        'group by bp.sNumeroOrden');
        end;

        if (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
        begin
            connection.zCommand.SQL.Add('SELECT '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.0000 AS Ajuste '+
                        'FROM bitacorademateriales bp '+
                        'WHERE  bp.sContrato =:Orden AND bp.didfecha = :Fecha '+
                        'and bp.sIdMaterial like :Id and bp.sNumeroOrden like :Folio '+
                        'group by bp.sNumeroOrden');
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
        connection.zCommand.ParamByName('fecha').AsDate   := dFechaActual;
        connection.zCommand.Open;

        while not connection.zCommand.Eof do
        begin
           dTotal  := dTotal  + connection.zCommand.FieldByName('Total').AsFloat;
           dAjuste := dAjuste + connection.zCommand.FieldByName('Ajuste').AsFloat;
           connection.zCommand.Next;
        end;
        ttripulacion_nacionales.Text := FloatToStr(dTotal + dAjuste);
       dFechaActual := dFechaActual + 1;
    end;
end;

procedure TfrmTripulacionGenerador.Recursos;
begin
  Categorias.Active := False;
  Categorias.SQL.Clear;
  if (chkPersonal.Checked) or (chkEquipo.Checked) then
  begin
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
  end;

  if chkBarco.Checked then
  begin
      Categorias.SQL.Add('select "<Todas>" as sIdRecurso, "" as sTipoObra from bitacoradeequipos '+
                  'union '+
                  'select tm.sIdTipomovimiento as sIdRecurso, tm.sDescripcion as sTipoObra from tiposdemovimiento tm '+
                  'where tm.sContrato =:ContratoBarco and sClasificacion = "Movimiento de barco" ');
      Categorias.ParamByName('ContratoBarco').AsString := global_contrato_barco;
  end;

  if chkPernocta.Checked then
  begin
      Categorias.SQL.Add('select "<Todas>" as sIdRecurso, "" as sTipoObra from cuentas '+
                  'union '+
                  'select c.sIdCuenta as sIdRecurso, c.sDescripcion as sTipoObra from cuentas c  ');
  end;

  if (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
  begin
      if chkAnexoC6.Checked then
         local_tipo := 'C6';

      if chkAnexoC7.Checked then
         local_tipo := 'C7';

      if chkAnexoC8.Checked then
         local_tipo := 'C8';

      Categorias.SQL.Add('select "<Todas>" as sIdRecurso, "" as sTipoObra from actividadesxanexo '+
                  'union '+
                  'select sNumeroActividad as sIdRecurso, substr(mDescripcion,1,250) as sTipoObra from actividadesxanexo '+
                  'where sContrato =:ContratoBarco and sTipoActividad = "Actividad" and sAnexo =:Anexo');
      Categorias.ParamByName('ContratoBarco').AsString := global_contrato_barco;
      Categorias.ParamByName('Anexo').AsString         := local_tipo;
  end;

  Categorias.Open;
end;

procedure TfrmTripulacionGenerador.CargarDatosAgrupados;
begin
    qry_tripulacion.Active := False;
    qry_tripulacion.SQL.Clear;
    if (chkPersonal.Checked) or (chkEquipo.Checked) then
    begin
        //>>Personal<<
        if chkPersonal.Checked then
           qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra, bp.sIdPersonal as sIdRecurso, bp.sDescripcion, '+
                          'concat(bp.sIdPersonal, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, e.sMedida, '+
                          '"<Agrupados Todos los Folios>" as sDescripcionFolio, "<Al corte>" as dFecha, '+
                          'xround(sum(bp.dCantHH),2) as Total, '+
                          'ifnull(SUM(bp.dAjuste),0) AS Ajuste, mr.dCantidad as dSolicitado '+
                         // 'FROM bitacoradepersonal bp '+
                          'from '               +
                          '( SELECT xround(sum((bp3.dCanthh )),2) as dCanthh,sum(bp3.dAjuste) AS dAjuste,bp3.sContrato, ' + #13#10 +
                             'bp3.dIdFecha,bp3.iIdDiario,bp3.sTipoObra,bp3.sIdPersonal,bp3.sDescripcion,bp3.sNumeroOrden,bp3.iIdTarea,' + #13#10 +
                             'bp3.iIdActividad,bp3.sIdPernocta ' +
                              'FROM bitacoradepersonal bp3 ' + #13#10 +
                              'WHERE bp3.scontrato =:Contrato AND bp3.didfecha >=:FechaI and bp3.dIdFecha <=:FechaF ' + #13#10 + 
                              'and bp3.stipopernocta like :Categoria and bp3.sNumeroOrden like :Folio' + #13#10 +
                              'group by bp3.dIdFecha, bp3.sNumeroOrden, bp3.stipopernocta,bp3.sIdPersonal' + #13#10 + 
                              ') bp ' + #13#10 + 
                       //   'inner join bitacoradeactividades ba '+
                      //    'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                      //    'bp.dIdFecha = ba.didfecha) ' +
                         // 'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                         // 'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) '+
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
                          'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.dIdFecha >=:fechaI and bp.dIdFecha <=:FechaF '+
                          'AND bp.sidPernocta =:pernocta '+
                          'and bp.sIdPersonal like :categoria GROUP BY bp.sIdPersonal, bp.sTipoObra order by e.iItemOrden  ')
        else
        //>>Equipo<<
           qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdEquipo as sIdRecurso, bp.sDescripcion, '+
                          'concat(bp.sIdEquipo, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, e.sMedida, '+
                          '"<Agrupados Todos los Folios>" as sDescripcionFolio, "<Al corte>" as dFecha, '+
                          'xround(sum(bp.dCantHH),2) as Total, '+
                          'ifnull(SUM(bp.dAjuste),0) AS Ajuste, mr.dCantidad as dSolicitado '+
                          //'FROM bitacoradeequipos bp '+
                          'from '               +
                          '( SELECT xround(sum((bp3.dCanthh )),2) as dCanthh,sum(bp3.dAjuste) AS dAjuste,bp3.sContrato, ' + #13#10 +
                             'bp3.dIdFecha,bp3.iIdDiario,bp3.sTipoObra,bp3.sIdEquipo,bp3.sDescripcion,bp3.sNumeroOrden,bp3.iIdTarea,' + #13#10 +
                             'bp3.iIdActividad,bp3.sIdPernocta ' +
                              'FROM bitacoradeequipos bp3 ' + #13#10 +
                              'WHERE bp3.scontrato =:Contrato AND bp3.didfecha >=:FechaI and bp3.dIdFecha <=:FechaF ' + #13#10 + 
                              'and bp3.sNumeroOrden like :Folio' + #13#10 +
                              'group by bp3.dIdFecha, bp3.sNumeroOrden,bp3.sIdEquipo' + #13#10 +
                              ') bp ' + #13#10 +
                   //       'inner join bitacoradeactividades ba '+
                    //      'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                    //       'bp.dIdFecha = ba.didfecha) ' +
                         // 'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                         // 'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) '+
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
                          'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.dIdFecha >=:fechaI and bp.dIdFecha <=:FechaF '+
                          'AND bp.sidPernocta =:pernocta '+
                          'and bp.sIdEquipo like :categoria GROUP BY bp.sIdEquipo, bp.sTipoObra order by e.iItemOrden ');
        qry_tripulacion.params.ParamByName('pernocta').asString  := local_global_pernocta;
    end;

    if chkBarco.Checked then
    begin
        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, tp.sIdTipoMovimiento as sIdRecurso, tp.sDescripcion, '+
                        'concat(tp.sIdTipoMovimiento, " ", tp.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, " " as sMedida, '+
                        '"<Agrupados Todos los Folios>" as sDescripcionFolio, '+
                        'xround(sum(bp.sFactor),6) as Total, '+
                        'ifnull(SUM(bp.dCantHH),0) AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM movimientosxfolios bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sNumeroOrden and o.sNumeroOrden = bp.sFolio ) '+
                        'inner join movimientosdeembarcacion me on (me.sContrato =bp.sContrato and me.dIdFecha = bp.dIdFecha and me.iIdDiario = bp.iIddiario ) '+
                        'inner join tiposdemovimiento tp on (tp.sContrato = bp.sContrato and tp.sIdTipoMovimiento = me.sClasificacion and tp.sClasificacion = "Movimiento de barco") '+
                        'WHERE bp.scontrato =:ContratoBarco AND bp.sNumeroOrden =:Contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and me.sClasificacion like :categoria and bp.sfolio like :folio '+
                        'group by tp.sIdTipoMovimiento order by tp.sIdTipoMovimiento');

    end;

    if chkPernocta.Checked then
    begin
        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, bp.sIdCuenta as sIdRecurso, c.sDescripcion, '+
                        'concat(bp.sIdCuenta, " ", c.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, c.sMedida, '+
                        '"<Agrupados Todos los Folios>" as sDescripcionFolio, '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.00000 AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM bitacoradepernocta bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                        'inner join cuentas c on (c.sIdCuenta = bp.sIdCuenta) '+
                        'WHERE bp.scontrato =:Contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and bp.sIdCuenta like :Categoria and bp.sNumeroOrden like :Folio and :ContratoBarco = :ContratoBarco '+
                        'group by bp.sIdCuenta order by c.sIdCuenta');

    end;

    if (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
    begin
        if chkAnexoC6.Checked then
           local_tipo := 'C6';

        if chkAnexoC7.Checked then
           local_tipo := 'C7';

        if chkAnexoC8.Checked then
           local_tipo := 'C8';

        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, bp.sIdMaterial as sIdRecurso, substr(ax.mDescripcion,1,250) as sDescripcion, '+
                        'concat(bp.sIdMaterial, " ", bp.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, ax.sMedida, ax.mDescripcion, '+
                        '"<Agrupados Todos los Folios>" as sDescripcionFolio, '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.00000 AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM bitacorademateriales bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                        'inner join actividadesxanexo ax on (ax.sContrato =:ContratoBarco and ax.sNumeroActividad = bp.sIdMaterial and ax.sTipoActividad = "Actividad") '+
                        'WHERE bp.scontrato =:contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and bp.sIdMaterial like :Categoria and bp.sNumeroOrden like :Folio and ax.sAnexo =:Anexo '+
                        'group by bp.sIdMaterial order by ax.iItemOrden');
        qry_tripulacion.params.ParamByName('Anexo').asString  := local_tipo;
    end;

    qry_tripulacion.params.ParamByName('contratoBarco').asString  := global_contrato_barco;
    qry_tripulacion.params.ParamByName('contrato').asString       := global_contrato;
    qry_tripulacion.params.ParamByName('fechaI').AsDate           := tdIdFecha.Date;
    qry_tripulacion.params.ParamByName('fechaF').AsDate           := tdIdFechaTermino.Date;
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
    qry_tripulacion.Open;

    if qry_Tripulacion.RecordCount = 0 then
    begin
      ttripulacion_nacionales.Text := '0';
      ttripulacion_extranjeros.Text := '0';
      qry_Tripulacion.Close;
    end;

end;

function TfrmTripulacionGenerador.TotalRecursoAgrupado(sParamRecurso, sParamFolio : string; dParamFechaI, dParamFechaF :tDate): double;
var
   dTotal : double;
   dFechaActual : tDate;
begin
    dTotal := 0;
    dFechaActual := tdIdFecha.Date;

    while dFechaActual <= tdIdFechaTermino.Date do
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        if chkEquipo.Checked then
        begin
            connection.zCommand.SQL.Add('SELECT '+
                      'xround(sum(bp.dCantHH),2) as Total, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                      'FROM bitacoradeequipos bp '+
                      'WHERE bp.scontrato =:Orden AND bp.sNumeroOrden like :Folio AND bp.didfecha =:Fecha and bp.sIdEquipo = :Id '+
                      'GROUP BY bp.sNumeroOrden, bp.sIdEquipo,bp.dIdFecha');
        end;
        if chkPersonal.Checked then
        begin
            connection.zCommand.SQL.Add('SELECT '+
                      'xround(sum(bp.dCantHH),2) as Total, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                      'FROM bitacoradepersonal bp '+
                      'WHERE bp.scontrato =:Orden AND bp.sNumeroOrden like:Folio and bp.didfecha =:Fecha and bp.sIdPersonal = :Id '+
                      'GROUP BY bp.sNumeroOrden, bp.sIdPersonal,bp.dIdFecha');
        end;

        if chkBarco.Checked then
        begin
            connection.zCommand.SQL.Add('SELECT '+
                        'round(sum(bp.sFactor),6) as Total, '+
                        'sum(bp.sFactorBarco) AS Ajuste '+
                        'FROM movimientosxfolios bp '+
                        'inner join movimientosdeembarcacion me on (me.sContrato =bp.sContrato and me.dIdFecha = bp.dIdFecha and me.iIdDiario = bp.iIddiario ) '+
                        'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden = :Orden AND bp.didfecha =:Fecha '+
                        'and me.sClasificacion like :Id and bp.sfolio like :Folio '+
                        'group by bp.sContrato, bp.sFolio');
           connection.zCommand.ParamByName('Contrato').AsString := global_contrato_barco;
        end;

        if chkPernocta.Checked then
        begin
            connection.zCommand.SQL.Add('SELECT   '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.0000 AS Ajuste '+
                        'FROM bitacoradepernocta bp '+
                        'WHERE  bp.sContrato = :Orden AND bp.didfecha =:fecha '+
                        'and bp.sIdCuenta like :Id and bp.sNumeroOrden like :folio '+
                        'group by bp.sNumeroOrden');
        end;

        if (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
        begin
            connection.zCommand.SQL.Add('SELECT '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.0000 AS Ajuste '+
                        'FROM bitacorademateriales bp '+
                        'WHERE  bp.sContrato =:Orden AND bp.didfecha = :Fecha '+
                        'and bp.sIdMaterial like :Id and bp.sNumeroOrden like :Folio '+
                        'group by bp.sNumeroOrden');
        end;

        connection.zCommand.ParamByName('Orden').AsString := global_contrato;
        connection.zCommand.ParamByName('folio').AsString := sParamFolio;
        connection.zCommand.ParamByName('Id').AsString    := sParamRecurso;
        connection.zCommand.ParamByName('fecha').AsDate   := dFechaActual;
        connection.zCommand.Open;


        while not connection.zCommand.Eof do
        begin
           dTotal := dTotal + (connection.zCommand.FieldByName('Total').AsFloat + connection.zCommand.FieldByName('Ajuste').AsFloat);
           connection.zCommand.Next;
        end;

        dFechaActual := dFechaActual + 1;
    end;
    result := dTotal;
end;

function TfrmTripulacionGenerador.TotalSolicitadoAgrupado(sParamRecurso, sParamFolio : string; dParamFechaI, dParamFechaF :tDate): double;
var
   dTotal : double;
   dFechaActual : tDate;
begin
    dTotal := 0;
    dFechaActual := tdIdFecha.Date;

    if (chkBarco.Checked) or (chkPernocta.Checked) or (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
    begin
        result := 0;
        exit;
    end;

    while dFechaActual <= tdIdFechaTermino.Date do
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        if chkEquipo.Checked then
        begin
            connection.zCommand.SQL.Add('select dCantidad as Total from moerecursos '+
                       ' where iIdMoe = '+

                       ' ( select m.iIdMoe '+
                       ' from moe m '+
                       ' where m.sContrato =:Orden '+
                       ' and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                       '                     from moe m1 '+
                       '                     where m1.sContrato =:Orden '+
                       '                     and m1.dIdFecha <= :Fecha '+
                       '                   ) '+
                       ' ) '+
                       ' and eTipoRecurso = "Equipo" and sIdRecurso =:Id '+
                       ' group by sIdrecurso');
        end;
        if chkPersonal.Checked then
        begin
            connection.zCommand.SQL.Add('select dCantidad as Total from moerecursos '+
                       ' where iIdMoe = '+

                       ' ( select m.iIdMoe '+
                       ' from moe m '+
                       ' where m.sContrato =:Orden '+
                       ' and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                       '                     from moe m1 '+
                       '                     where m1.sContrato =:Orden '+
                       '                     and m1.dIdFecha <= :Fecha '+
                       '                   ) '+
                       ' ) '+
                       ' and eTipoRecurso = "Personal" and sIdRecurso =:Id '+
                       ' group by sIdrecurso');
        end;
        connection.zCommand.ParamByName('Orden').AsString := global_contrato;
        connection.zCommand.ParamByName('Id').AsString    := sParamRecurso;
        connection.zCommand.ParamByName('fecha').AsDate   := dFechaActual;
        connection.zCommand.Open;

        while not connection.zCommand.Eof do
        begin
           dTotal := dTotal + (connection.zCommand.FieldByName('Total').AsFloat) ;
           connection.zCommand.Next;
        end;

        dFechaActual := dFechaActual + 1;
    end;
    result := dTotal;
end;

procedure TfrmTripulacionGenerador.CargarDatosAgrupadosFolio;
begin
    qry_tripulacion.Active := False;
    qry_tripulacion.SQL.Clear;

    if (chkPersonal.Checked) or (chkEquipo.Checked) then
    begin
        //>>Personal<<
        if chkPersonal.Checked then
           qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra, "<Todas>" as sIdRecurso, bp.sDescripcion, '+
                          '"<Todas las Partidas>" as sDescripcionRecurso, bp.sNumeroOrden, e.sMedida, '+
                          'o.sIdFolio as sDescripcionFolio, "<Al corte>" as dFecha, '+
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
                          'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.dIdFecha >=:fechaI and bp.dIdFecha <=:FechaF '+
                          'AND bp.sidPernocta =:pernocta '+
                          'and bp.sIdPersonal like :categoria GROUP BY bp.sNumeroOrden order by o.iOrden ')
        else
        //>>Equipo<<
           qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra, "<Todas>" as sIdRecurso, bp.sDescripcion, '+
                          '"<Todas las Partidas>" as sDescripcionRecurso, bp.sNumeroOrden, e.sMedida, '+
                          'o.sIdFolio as sDescripcionFolio, "<Al corte>" as dFecha, '+
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
                          'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.dIdFecha >=:fechaI and bp.dIdFecha <=:FechaF '+
                          'AND bp.sidPernocta =:pernocta '+
                          'and bp.sIdEquipo like :categoria GROUP BY bp.sNumeroOrden order by o.iOrden ');
        qry_tripulacion.params.ParamByName('pernocta').asString  := local_global_pernocta;
    end;

    if chkBarco.Checked then
    begin
        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, "<Todas>" as sIdRecurso, tp.sDescripcion, '+
                        '"<Todas las Partidas>" as sDescripcionRecurso, o.sNumeroOrden, " " as sMedida, '+
                        'o.sIdFolio as sDescripcionFolio, '+
                        'xround(sum(bp.sFactor),6) as Total, '+
                        'ifnull(SUM(bp.dCantHH),0) AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM movimientosxfolios bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sNumeroOrden and o.sNumeroOrden = bp.sFolio ) '+
                        'inner join movimientosdeembarcacion me on (me.sContrato =bp.sContrato and me.dIdFecha = bp.dIdFecha and me.iIdDiario = bp.iIddiario ) '+
                        'inner join tiposdemovimiento tp on (tp.sContrato = bp.sContrato and tp.sIdTipoMovimiento = me.sClasificacion and tp.sClasificacion = "Movimiento de barco") '+
                        'WHERE bp.scontrato =:ContratoBarco AND bp.sNumeroOrden =:Contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and me.sClasificacion like :categoria and bp.sfolio like :folio '+
                        'group by bp.sFolio order by o.iOrden ');
    end;

    if chkPernocta.Checked then
    begin
        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra,  "<Todas>" as sIdRecurso, c.sDescripcion, '+
                        '"<Agrupados Todos los Folios>"  as sDescripcionRecurso, o.sNumeroOrden, c.sMedida, '+
                        'o.sIdFolio as sDescripcionFolio, '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.00000 AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM bitacoradepernocta bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                        'inner join cuentas c on (c.sIdCuenta = bp.sIdCuenta) '+
                        'WHERE bp.scontrato =:Contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and bp.sIdCuenta like :Categoria and bp.sNumeroOrden like :Folio and :ContratoBarco = :ContratoBarco '+
                        'group by bp.sNumeroOrden order by o.iOrden');
    end;

    if (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
    begin
        if chkAnexoC6.Checked then
           local_tipo := 'C6';

        if chkAnexoC7.Checked then
           local_tipo := 'C7';

        if chkAnexoC8.Checked then
           local_tipo := 'C8';

        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, "<Todas>" as sIdRecurso, substr(ax.mDescripcion,1,250) as sDescripcion, '+
                        '"<Agrupados Todos los Folios>" as sDescripcionRecurso, o.sNumeroOrden, ax.sMedida, ax.mDescripcion, '+
                        'o.sIdFolio as sDescripcionFolio, '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.00000 AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM bitacorademateriales bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                        'inner join actividadesxanexo ax on (ax.sContrato =:ContratoBarco and ax.sNumeroActividad = bp.sIdMaterial and ax.sTipoActividad = "Actividad") '+
                        'WHERE bp.scontrato =:contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and bp.sIdMaterial like :Categoria and bp.sNumeroOrden like :Folio '+
                        'group by bp.sNumeroOrden order by o.iOrden');
        qry_tripulacion.params.ParamByName('Anexo').asString  := local_tipo;
    end;

    qry_tripulacion.params.ParamByName('contratoBarco').asString  := global_contrato_barco;
    qry_tripulacion.params.ParamByName('contrato').asString       := global_contrato;
    qry_tripulacion.params.ParamByName('fechaI').AsDate           := tdIdFecha.Date;
    qry_tripulacion.params.ParamByName('fechaF').AsDate           := tdIdFechaTermino.Date;
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
    qry_tripulacion.Open;

    if qry_Tripulacion.RecordCount = 0 then
    begin
      ttripulacion_nacionales.Text := '0';
      ttripulacion_extranjeros.Text := '0';
      qry_Tripulacion.Close;
    end;

end;

function TfrmTripulacionGenerador.TotalRecursoAgrupadoFolio(sParamRecurso, sParamFolio : string; dParamFechaI, dParamFechaF :tDate): double;
var
   dTotal : double;
   dFechaActual : tDate;
begin
    dTotal := 0;
    dFechaActual := tdIdFecha.Date;

    while dFechaActual <= tdIdFechaTermino.Date do
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        if (chkPersonal.Checked) or (chkEquipo.Checked) then
        begin
            if chkEquipo.Checked then
            begin
                connection.zCommand.SQL.Add('SELECT '+
                          'xround(sum(bp.dCantHH),2) as Total, '+
                          'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                          'FROM bitacoradeequipos bp '+
                          'WHERE bp.scontrato =:Orden AND bp.sNumeroOrden like :Folio AND bp.didfecha =:Fecha  '+
                          'GROUP BY bp.sNumeroOrden,bp.didfecha,bp.sIdEquipo');
            end;
            if chkPersonal.Checked then
            begin
                connection.zCommand.SQL.Add('SELECT '+
                          'xround(sum(bp.dCantHH),2) as Total, '+
                          'ifnull(SUM(bp.dAjuste),0) AS Ajuste '+
                          'FROM bitacoradepersonal bp '+
                          'WHERE bp.scontrato =:Orden AND bp.sNumeroOrden like:Folio and bp.didfecha =:Fecha  '+
                          'GROUP BY bp.sNumeroOrden,bp.didfecha,bp.sIdPersonal');
            end;
        end;

        if chkBarco.Checked then
        begin
            connection.zCommand.SQL.Add('SELECT '+
                        'xround(sum(bp.sFactor),6) as Total, '+
                        'sum(bp.sFactorBarco) AS Ajuste '+
                        'FROM movimientosxfolios bp '+
                        'inner join movimientosdeembarcacion me on (me.sContrato =bp.sContrato and me.dIdFecha = bp.dIdFecha and me.iIdDiario = bp.iIddiario ) '+
                        'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden = :Orden AND bp.didfecha =:Fecha '+
                        'and bp.sfolio like :Folio '+
                        'group by bp.sContrato, bp.sFolio');
           connection.zCommand.ParamByName('Contrato').AsString := global_contrato_barco;
        end;

        if chkPernocta.Checked then
        begin
            connection.zCommand.SQL.Add('SELECT   '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.0000 AS Ajuste '+
                        'FROM bitacoradepernocta bp '+
                        'WHERE  bp.sContrato = :Orden AND bp.didfecha =:fecha '+
                        'and bp.sNumeroOrden like :folio '+
                        'group by bp.sNumeroOrden');
        end;

        if (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
        begin
            connection.zCommand.SQL.Add('SELECT '+
                        'sum(bp.dCantidad) as Total, '+
                        '0.0000 AS Ajuste '+
                        'FROM bitacorademateriales bp '+
                        'WHERE  bp.sContrato =:Orden AND bp.didfecha = :Fecha '+
                        'and bp.sNumeroOrden like :Folio '+
                        'group by bp.sNumeroOrden');
        end;

        connection.zCommand.ParamByName('Orden').AsString := global_contrato;
        connection.zCommand.ParamByName('folio').AsString := sParamFolio;
        connection.zCommand.ParamByName('fecha').AsDate   := dFechaActual;
        connection.zCommand.Open;


        while not connection.zCommand.Eof do
        begin
           dTotal := dTotal + (connection.zCommand.FieldByName('Total').AsFloat + connection.zCommand.FieldByName('Ajuste').AsFloat);
           connection.zCommand.Next;
        end;

        dFechaActual := dFechaActual + 1;
    end;
    result := dTotal;
end;

function TfrmTripulacionGenerador.TotalSolicitadoAgrupadoFolio(sParamRecurso, sParamFolio : string; dParamFechaI, dParamFechaF :tDate): double;
var
   dTotal : double;
   dFechaActual : tDate;
begin
    dTotal := 0;
    dFechaActual := tdIdFecha.Date;

    if (chkBarco.Checked) or (chkPernocta.Checked) or (chkAnexoC6.Checked) or (chkAnexoC7.Checked) or (chkAnexoC8.Checked) then
    begin
        result := 0;
        exit;
    end;

    while dFechaActual <= tdIdFechaTermino.Date do
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        if chkEquipo.Checked then
        begin
            connection.zCommand.SQL.Add('select sum(dCantidad) as Total from moerecursos '+
                       ' where iIdMoe = '+

                       ' ( select m.iIdMoe '+
                       ' from moe m '+
                       ' where m.sContrato =:Orden '+
                       ' and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                       '                     from moe m1 '+
                       '                     where m1.sContrato =:Orden '+
                       '                     and m1.dIdFecha <= :Fecha '+
                       '                   ) '+
                       ' ) '+
                       ' and eTipoRecurso = "Equipo" '+
                       ' group by iIdMoe');
        end;
        if chkPersonal.Checked then
        begin
            connection.zCommand.SQL.Add('select sum(dCantidad) as Total from moerecursos '+
                       ' where iIdMoe = '+

                       ' ( select m.iIdMoe '+
                       ' from moe m '+
                       ' where m.sContrato =:Orden '+
                       ' and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                       '                     from moe m1 '+
                       '                     where m1.sContrato =:Orden '+
                       '                     and m1.dIdFecha <= :Fecha '+
                       '                   ) '+
                       ' ) '+
                       ' and eTipoRecurso = "Personal" '+
                       ' group by iIdMoe');
        end;
        connection.zCommand.ParamByName('Orden').AsString := global_contrato;
        connection.zCommand.ParamByName('fecha').AsDate   := dFechaActual;
        connection.zCommand.Open;

        while not connection.zCommand.Eof do
        begin
           dTotal := dTotal + (connection.zCommand.FieldByName('Total').AsFloat) ;
           connection.zCommand.Next;
        end;

        dFechaActual := dFechaActual + 1;
    end;
    result := dTotal;
end;


end.

