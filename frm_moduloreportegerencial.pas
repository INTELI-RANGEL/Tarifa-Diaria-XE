unit frm_moduloreportegerencial;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, OleCtrls, 
  DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Global, 
  
  NxClasses, DateUtils, NxPageControl, Grids,
  DBGrids, RXDBCtrl, Newpanel;

type
  TfrmModuloReporteGerencial = class(TForm)
    pBarraAvance: TProgressBar;
    QryFolios: TZQuery;
    dsFolios: TDataSource;
    QryFoliossContrato: TStringField;
    QryFoliossIdFolio: TStringField;
    QryFoliossNumeroOrden: TStringField;
    QryFoliossDescripcionCorta: TStringField;
    QryFoliossOficioAutorizacion: TStringField;
    QryFoliosmDescripcion: TMemoField;
    QryFoliossIdTipoOrden: TStringField;
    QryFoliossApoyo: TStringField;
    QryFoliossIdPlataforma: TStringField;
    QryFoliossIdPernocta: TStringField;
    QryFoliosdFiProgramado: TDateField;
    QryFoliosdFfProgramado: TDateField;
    QryFolioscIdStatus: TStringField;
    QryFoliosmComentarios: TMemoField;
    QryFoliossFormato: TStringField;
    QryFoliosiConsecutivo: TIntegerField;
    QryFoliosiConsecutivoTierra: TIntegerField;
    QryFoliosiJornada: TIntegerField;
    QryFolioslGeneraAnexo: TStringField;
    QryFolioslGeneraConsumibles: TStringField;
    QryFolioslGeneraPersonal: TStringField;
    QryFolioslGeneraEquipo: TStringField;
    QryFoliossDepsolicitante: TStringField;
    QryFoliosdFechaInicioT: TDateField;
    QryFoliosdFechaSitioM: TDateField;
    QryFoliossEquipo: TStringField;
    QryFoliossPozo: TStringField;
    QryFoliosdFechaElaboracion: TDateField;
    QryFoliossPuestoPep: TStringField;
    QryFoliossFirmantePep: TStringField;
    QryFoliossPuestocia: TStringField;
    QryFoliossFirmantecia: TStringField;
    QryFolioslMostrarAvanceProgramado: TStringField;
    QryFoliossTipoOrden: TStringField;
    QryFoliosbAvanceFrente: TStringField;
    QryFoliosbAvanceContrato: TStringField;
    QryFoliosbComentarios: TStringField;
    QryFoliosbPermisos: TStringField;
    QryFoliosbTipoAdmon: TStringField;
    QryFoliosbCostaFuera: TStringField;
    QryFoliossTipoPrograma: TStringField;
    QryFoliossTipoImpresionActividad: TStringField;
    QryFoliossTipoAvanceAdmon: TStringField;
    QryFoliosiDecimales: TIntegerField;
    QryFoliosiNiveles: TIntegerField;
    QryFolioslImprimeProgramado: TStringField;
    QryFolioslImprimeFisico: TStringField;
    QryFolioslImprimePlaticas: TStringField;
    QryFolioslImprimePersonalTM: TStringField;
    QryFolioslPersonalxPartida: TStringField;
    QryFolioslImprimeFases: TStringField;
    QryFolioslMostrarPartidasReportes: TStringField;
    QryFolioslMostrarPartidasGeneradores: TStringField;
    QryFoliosdFechaIniPReportes: TDateField;
    QryFoliosdFechaFinPReportes: TDateField;
    QryFoliosdFechaIniPGeneradores: TDateField;
    QryFoliosdFechaFinPGeneradores: TDateField;
    QryFolioslEstado: TStringField;
    Panel1: TPanel;
    Label1: TLabel;
    dIdFecha: TDateTimePicker;
    Panel2: TPanel;
    Panel3: TPanel;
    NxPageControl1: TNxPageControl;
    NxTabSheet1: TNxTabSheet;
    Grid_Reportes: TDBGrid;
    zq_personalabordo: TZQuery;
    ds_personalabordo: TDataSource;
    zq_personalabordoiId: TIntegerField;
    zq_personalabordodIdFecha: TDateTimeField;
    zq_personalabordosContrato: TStringField;
    zq_personalabordosPartida: TStringField;
    zq_personalabordosDescripcion: TStringField;
    zq_personalabordodCantidad: TFloatField;
    zq_personalabordodCantidadaBordo: TFloatField;
    NxTabSheet2: TNxTabSheet;
    DBGridPersonalFaltante: TDBGrid;
    Panel: tNewGroupBox;
    ListaObjeto: TRxDBGrid;
    zq_PersonalFaltante: TZQuery;
    ds_PersonalFaltante: TDataSource;
    zq_PersonalFaltanteiId: TIntegerField;
    zq_PersonalFaltantedIdFecha: TDateField;
    zq_PersonalFaltantesContrato: TStringField;
    zq_PersonalFaltantesIdRecurso: TStringField;
    zq_PersonalFaltantedCantidad: TFloatField;
    zq_PersonalFaltantesTipoRecurso: TStringField;
    zq_PersonalFaltantesTipoFaltante: TStringField;
    zq_Personal: TZQuery;
    ds_Personal: TDataSource;
    zq_PersonalsContrato: TStringField;
    zq_PersonalsIdPersonal: TStringField;
    zq_PersonaliItemOrden: TIntegerField;
    zq_PersonalsDescripcion: TStringField;
    zq_PersonalsIdTipoPersonal: TStringField;
    zq_PersonalsMedida: TStringField;
    zq_PersonaldCantidad: TFloatField;
    zq_PersonaldCostoMN: TFloatField;
    zq_PersonaldCostoDLL: TFloatField;
    zq_PersonaldVentaMN: TFloatField;
    zq_PersonaldVentaDLL: TFloatField;
    zq_PersonaldFechaInicio: TDateField;
    zq_PersonaldFechaFinal: TDateField;
    zq_PersonallProrrateo: TStringField;
    zq_PersonallCobro: TStringField;
    zq_PersonallImprime: TStringField;
    zq_PersonallAplicaTM: TStringField;
    zq_PersonaliJornada: TIntegerField;
    zq_PersonallDistribuye: TStringField;
    zq_PersonallPernocta: TStringField;
    zq_PersonalsAgrupaPersonal: TStringField;
    zq_PersonallTotalizarPernocta: TStringField;
    zq_PersonallAplicaGerencial: TStringField;
    zq_PersonaliId_AgrupadorPersonal: TIntegerField;
    zq_PersonallSumaSolicitado: TStringField;
    zq_PersonalFaltantesPersonal: TStringField;
    BuscaObjeto: TZReadOnlyQuery;
    ds_buscaobjeto: TDataSource;
    NxTabSheet3: TNxTabSheet;
    DBGridPersonalPendiente: TDBGrid;
    zq_PersonalPendiente: TZQuery;
    ds_PersonalPendiente: TDataSource;
    zq_PersonalPendienteiId: TIntegerField;
    zq_PersonalPendientedIdFecha: TDateField;
    zq_PersonalPendientesContrato: TStringField;
    zq_PersonalPendientesIdRecurso: TStringField;
    zq_PersonalPendientedCantidad: TFloatField;
    zq_PersonalPendientesTipoRecurso: TStringField;
    zq_PersonalPendientesTipoFaltante: TStringField;
    zq_PersonalPendientesPersonal: TStringField;
    NxTabSheet4: TNxTabSheet;
    NxTabSheet5: TNxTabSheet;
    DBGridEquipoDescuento: TDBGrid;
    zq_EquipoDescuento: TZQuery;
    ds_EquipoDescuento: TDataSource;
    zq_EquipoDescuentoiId: TIntegerField;
    zq_EquipoDescuentodIdFecha: TDateField;
    zq_EquipoDescuentosContrato: TStringField;
    zq_EquipoDescuentosIdRecurso: TStringField;
    zq_EquipoDescuentodCantidad: TFloatField;
    zq_EquipoDescuentosTipoRecurso: TStringField;
    zq_EquipoDescuentosTipoFaltante: TStringField;
    zq_Equipos: TZQuery;
    ds_Equipos: TDataSource;
    zq_EquipossContrato: TStringField;
    zq_EquipossIdEquipo: TStringField;
    zq_EquiposiItemOrden: TIntegerField;
    zq_EquipossDescripcion: TStringField;
    zq_EquipossIdTipoEquipo: TStringField;
    zq_EquipossMedida: TStringField;
    zq_EquiposdCantidad: TFloatField;
    zq_EquiposdCostoMN: TFloatField;
    zq_EquiposdCostoDLL: TFloatField;
    zq_EquiposdVentaMN: TFloatField;
    zq_EquiposdVentaDLL: TFloatField;
    zq_EquiposdFechaInicio: TDateField;
    zq_EquiposdFechaFinal: TDateField;
    zq_EquiposlProrrateo: TStringField;
    zq_EquiposlCobro: TStringField;
    zq_EquiposlImprime: TStringField;
    zq_EquiposiJornada: TIntegerField;
    zq_EquiposlDistribuye: TStringField;
    zq_EquiposlCuadraEquipo: TStringField;
    zq_EquiposlAplicaDiesel: TStringField;
    zq_EquipossDescripcionDiesel: TStringField;
    zq_EquiposlSumaSolicitado: TStringField;
    zq_EquipoDescuentosEquipo: TStringField;
    zq_EquiposFueradeOperacion: TZQuery;
    ds_EquiposFueradeOperacion: TDataSource;
    DBGridFO: TDBGrid;
    zq_EquiposFueradeOperacioniId: TIntegerField;
    zq_EquiposFueradeOperaciondIdFecha: TDateField;
    zq_EquiposFueradeOperacionsContrato: TStringField;
    zq_EquiposFueradeOperacionsIdRecurso: TStringField;
    zq_EquiposFueradeOperaciondCantidad: TFloatField;
    zq_EquiposFueradeOperacionsTipoRecurso: TStringField;
    zq_EquiposFueradeOperacionsTipoFaltante: TStringField;
    zq_EquiposFueradeOperacionsEquipo: TStringField;
    function IsIn(Valor: string; Lista: TStringList): Boolean;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure zq_personalabordoAfterInsert(DataSet: TDataSet);
    procedure dIdFechaChange(Sender: TObject);
    procedure zq_PersonalFaltantesIdRecursoChange(Sender: TField);
    procedure zq_PersonalFaltanteAfterInsert(DataSet: TDataSet);
    procedure ListaObjetoExit(Sender: TObject);
    procedure ListaObjetoKeyPress(Sender: TObject; var Key: Char);
    procedure zq_PersonalPendientesIdRecursoChange(Sender: TField);
    procedure zq_PersonalPendienteAfterInsert(DataSet: TDataSet);
    procedure zq_EquipoDescuentoAfterInsert(DataSet: TDataSet);
    procedure zq_EquipoDescuentosIdRecursoChange(Sender: TField);
    procedure zq_EquiposFueradeOperacionsIdRecursoChange(Sender: TField);
    procedure zq_EquiposFueradeOperacionAfterInsert(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmModuloReporteGerencial: TfrmModuloReporteGerencial;
  PrimerFolio: String;

procedure Split(Delimiter: Char; Str: string; ListOfStrings: TStrings);

implementation

uses frm_connection;

{$R *.dfm}

procedure Split(Delimiter: Char; Str: string; ListOfStrings: TStrings);
begin
   ListOfStrings.Clear;
   ListOfStrings.Delimiter     := Delimiter;
   ListOfStrings.DelimitedText := Str;
end;

procedure TfrmModuloReporteGerencial.dIdFechaChange(Sender: TObject);
begin
  zq_Personalabordo.Active := False;
  zq_Personalabordo.ParamByName('Contrato').AsString := Global_Contrato;
  zq_Personalabordo.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_Personalabordo.Open;

  zq_PersonalFaltante.Active := False;
  zq_PersonalFaltante.ParamByName('Contrato').AsString := Global_Contrato;
  zq_PersonalFaltante.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_PersonalFaltante.Open;

  zq_PersonalPendiente.Active := False;
  zq_PersonalPendiente.ParamByName('Contrato').AsString := Global_Contrato;
  zq_PersonalPendiente.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_PersonalPendiente.Open;

  zq_EquipoDescuento.Active := False;
  zq_EquipoDescuento.ParamByName('Contrato').AsString := Global_Contrato;
  zq_EquipoDescuento.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_EquipoDescuento.Open;

  zq_EquiposFueradeOperacion.Active := False;
  zq_EquiposFueradeOperacion.ParamByName('Contrato').AsString := Global_Contrato;
  zq_EquiposFueradeOperacion.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_EquiposFueradeOperacion.Open;
end;

procedure TfrmModuloReporteGerencial.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := CaFree;
end;

procedure TfrmModuloReporteGerencial.FormShow(Sender: TObject);
begin
  dIdFecha.Date := IncDay(Now, -1);

  zq_Personal.Active := False;
  zq_Personal.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
  zq_Personal.Open;

  zq_Equipos.Active := False;
  zq_Equipos.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
  zq_Equipos.Open;

  zq_Personalabordo.Active := False;
  zq_Personalabordo.ParamByName('Contrato').AsString := Global_Contrato;
  zq_Personalabordo.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_Personalabordo.Open;

  zq_PersonalFaltante.Active := False;
  zq_PersonalFaltante.ParamByName('Contrato').AsString := Global_Contrato;
  zq_PersonalFaltante.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_PersonalFaltante.Open;

  zq_PersonalPendiente.Active := False;
  zq_PersonalPendiente.ParamByName('Contrato').AsString := Global_Contrato;
  zq_PersonalPendiente.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_PersonalPendiente.Open;

  zq_EquipoDescuento.Active := False;
  zq_EquipoDescuento.ParamByName('Contrato').AsString := Global_Contrato;
  zq_EquipoDescuento.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_EquipoDescuento.Open;

  zq_EquiposFueradeOperacion.Active := False;
  zq_EquiposFueradeOperacion.ParamByName('Contrato').AsString := Global_Contrato;
  zq_EquiposFueradeOperacion.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
  zq_EquiposFueradeOperacion.Open;
end;

function TfrmModuloReporteGerencial.IsIn(Valor: string; Lista: TStringList): Boolean;
var
   nIdx: Integer;
begin
   Result := False;
   for nIdx := 0 to Lista.Count - 1 do begin
      if Lista[nIdx] = Valor then begin
         Result := True;
         Break;
      end;
   end;
end;

procedure TfrmModuloReporteGerencial.ListaObjetoExit(Sender: TObject);
begin
  if Panel.Visible = True then
  begin
    if BuscaObjeto.RecordCount > 0 then //Aqui debo aplicar correccion
    begin
      if NxPageControl1.ActivePageIndex = 1 then
      begin
        zq_PersonalFaltante.FieldValues['sIdRecurso'] := BuscaObjeto.FieldValues['sIdPersonal'];
        DBGridPersonalFaltante.SetFocus;
      end;
      if NxPageControl1.ActivePageIndex = 2 then
      begin
        zq_PersonalPendiente.FieldValues['sIdRecurso'] := BuscaObjeto.FieldValues['sIdPersonal'];
        DBGridPersonalPendiente.SetFocus;
      end;
      if NxPageControl1.ActivePageIndex = 3 then
      begin
        zq_EquipoDescuento.FieldValues['sIdRecurso'] := BuscaObjeto.FieldValues['sIdEquipo'];
        DBGridEquipoDescuento.SetFocus;
      end;
      if NxPageControl1.ActivePageIndex = 4 then
      begin
        zq_EquiposFueradeOperacion.FieldValues['sIdRecurso'] := BuscaObjeto.FieldValues['sIdEquipo'];
        DBGridFO.SetFocus;
      end;
    end;
    Panel.Visible := False;
  end
end;

procedure TfrmModuloReporteGerencial.ListaObjetoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
  begin
    if NxPageControl1.ActivePageIndex = 1 then begin
      DBGridPersonalFaltante.SetFocus;
    end;
    if NxPageControl1.ActivePageIndex = 2 then begin
      DBGridPersonalPendiente.SetFocus;
    end;
    if NxPageControl1.ActivePageIndex = 3 then begin
      DBGridEquipoDescuento.SetFocus;
    end;
    if NxPageControl1.ActivePageIndex = 4 then begin
      DBGridFO.SetFocus;
    end;
  end;
end;

procedure TfrmModuloReporteGerencial.zq_EquipoDescuentoAfterInsert(
  DataSet: TDataSet);
begin
  TZQuery(DataSet).FieldByName('dIdFecha').AsDateTime := dIdFecha.Date;
  TZQuery(DataSet).FieldByName('sContrato').AsString := Global_Contrato;
  TZQuery(DataSet).FieldByName('sTipoRecurso').AsString := 'Equipo';
  TZQuery(DataSet).FieldByName('sTipoFaltante').AsString := 'Descuento';
end;

procedure TfrmModuloReporteGerencial.zq_EquipoDescuentosIdRecursoChange(
  Sender: TField);
Var
  sDescripcion: String;
begin
  if not zq_EquipoDescuento.FieldByName('sIdRecurso').IsNull then
  begin
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('select sIdEquipo, sDescripcion from equipos ' +
      'WHERE sContrato = :ContratoBarco And sIdEquipo = :Equipo');
    Connection.QryBusca.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
    Connection.QryBusca.ParamByName('Equipo').AsString := zq_EquipoDescuentosIdRecurso.Text;
    Connection.QryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
//      zq_EquipoDescuento.FieldValues['sIdRecurso'] := Connection.QryBusca.FieldValues['sIdPersonal'];
      zq_EquipoDescuento.FieldValues['sEquipo'] := Connection.QryBusca.FieldValues['sDescripcion'];
    end
    else
      if not zq_EquipoDescuento.FieldByName('sIdRecurso').IsNull then
        if Trim(zq_EquipoDescuento.FieldByName('sIdRecurso').AsString) <> '' then
        begin
          sDescripcion := '%' + Trim(zq_EquipoDescuento.FieldValues['sIdRecurso']) + '%';
          BuscaObjeto.Active := False;
          ListaObjeto.Columns.Clear;

          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[0].FieldName := 'sIdEquipo';
          ListaObjeto.Columns[0].Width := 80;
          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[1].FieldName := 'sDescripcion';
          ListaObjeto.Columns[1].Width := 500;

          BuscaObjeto.SQL.Clear;
          BuscaObjeto.SQL.Add('Select sIdEquipo, sDescripcion  from equipos WHERE ' +
            'sContrato = :ContratoBarco And sDescripcion Like :Descripcion');
          BuscaObjeto.Params.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          BuscaObjeto.Params.ParamByName('Descripcion').AsString := sDescripcion;
          BuscaObjeto.Open;
               // yavienedeRegreso := 'No' ;
          Panel.Visible := True;
          Panel.Height  := 200;
          Panel.Width   := 400;
          ListaObjeto.SetFocus
        end
  end;

end;

procedure TfrmModuloReporteGerencial.zq_personalabordoAfterInsert(
  DataSet: TDataSet);
begin
  zq_personalabordo.FieldByName('dIdFecha').AsDateTime := dIdFecha.Date;
  zq_Personalabordo.FieldByName('sContrato').AsString := Global_Contrato;
end;

procedure TfrmModuloReporteGerencial.zq_PersonalFaltanteAfterInsert(
  DataSet: TDataSet);
begin
  TZQuery(DataSet).FieldByName('dIdFecha').AsDateTime := dIdFecha.Date;
  TZQuery(DataSet).FieldByName('sContrato').AsString := Global_Contrato;
  TZQuery(DataSet).FieldByName('sTipoRecurso').AsString := 'Personal';
  TZQuery(DataSet).FieldByName('sTipoFaltante').AsString := 'Faltante';
end;

procedure TfrmModuloReporteGerencial.zq_PersonalFaltantesIdRecursoChange(
  Sender: TField);
Var
  sDescripcion: String;
begin
  if not zq_PersonalFaltante.FieldByName('sIdRecurso').IsNull then
  begin
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('select sIdPersonal, sDescripcion from personal ' +
      'WHERE sContrato = :ContratoBarco And sIdPersonal = :Personal');
    Connection.QryBusca.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
    Connection.QryBusca.ParamByName('Personal').AsString := zq_PersonalFaltantesIdRecurso.Text;
    Connection.QryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
//      zq_PersonalFaltante.FieldValues['sIdRecurso'] := Connection.QryBusca.FieldValues['sIdPersonal'];
      zq_PersonalFaltante.FieldValues['sPersonal'] := Connection.QryBusca.FieldValues['sDescripcion'];
    end
    else
      if not zq_PersonalFaltante.FieldByName('sIdRecurso').IsNull then
        if Trim(zq_PersonalFaltante.FieldByName('sIdRecurso').AsString) <> '' then
        begin
          sDescripcion := '%' + Trim(zq_PersonalFaltante.FieldValues['sIdRecurso']) + '%';
          BuscaObjeto.Active := False;
          ListaObjeto.Columns.Clear;

          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[0].FieldName := 'sIdPersonal';
          ListaObjeto.Columns[0].Width := 80;
          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[1].FieldName := 'sDescripcion';
          ListaObjeto.Columns[1].Width := 500;

          BuscaObjeto.SQL.Clear;
          BuscaObjeto.SQL.Add('Select sIdPersonal, sDescripcion  from personal WHERE ' +
            'sContrato = :ContratoBarco And sDescripcion Like :Descripcion');
          BuscaObjeto.Params.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          BuscaObjeto.Params.ParamByName('Descripcion').AsString := sDescripcion;
          BuscaObjeto.Open;
               // yavienedeRegreso := 'No' ;
          Panel.Visible := True;
          Panel.Height  := 200;
          Panel.Width   := 400;
          ListaObjeto.SetFocus
        end
  end;

end;

procedure TfrmModuloReporteGerencial.zq_PersonalPendienteAfterInsert(
  DataSet: TDataSet);
begin
  TZQuery(DataSet).FieldByName('dIdFecha').AsDateTime := dIdFecha.Date;
  TZQuery(DataSet).FieldByName('sContrato').AsString := Global_Contrato;
  TZQuery(DataSet).FieldByName('sTipoRecurso').AsString := 'Personal';
  TZQuery(DataSet).FieldByName('sTipoFaltante').AsString := 'Pendiente';
end;

procedure TfrmModuloReporteGerencial.zq_PersonalPendientesIdRecursoChange(
  Sender: TField);
Var
  sDescripcion: String;
begin
  if not zq_PersonalPendiente.FieldByName('sIdRecurso').IsNull then
  begin
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('select sIdPersonal, sDescripcion from personal ' +
      'WHERE sContrato = :ContratoBarco And sIdPersonal = :Personal');
    Connection.QryBusca.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
    Connection.QryBusca.ParamByName('Personal').AsString := zq_PersonalPendientesIdRecurso.Text;
    Connection.QryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
//      zq_PersonalPendiente.FieldValues['sIdRecurso'] := Connection.QryBusca.FieldValues['sIdPersonal'];
      zq_PersonalPendiente.FieldValues['sPersonal'] := Connection.QryBusca.FieldValues['sDescripcion'];
    end
    else
      if not zq_PersonalPendiente.FieldByName('sIdRecurso').IsNull then
        if Trim(zq_PersonalPendiente.FieldByName('sIdRecurso').AsString) <> '' then
        begin
          sDescripcion := '%' + Trim(zq_PersonalPendiente.FieldValues['sIdRecurso']) + '%';
          BuscaObjeto.Active := False;
          ListaObjeto.Columns.Clear;

          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[0].FieldName := 'sIdPersonal';
          ListaObjeto.Columns[0].Width := 80;
          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[1].FieldName := 'sDescripcion';
          ListaObjeto.Columns[1].Width := 500;

          BuscaObjeto.SQL.Clear;
          BuscaObjeto.SQL.Add('Select sIdPersonal, sDescripcion  from personal WHERE ' +
            'sContrato = :ContratoBarco And sDescripcion Like :Descripcion');
          BuscaObjeto.Params.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          BuscaObjeto.Params.ParamByName('Descripcion').AsString := sDescripcion;
          BuscaObjeto.Open;
               // yavienedeRegreso := 'No' ;
          Panel.Visible := True;
          Panel.Height  := 200;
          Panel.Width   := 400;
          ListaObjeto.SetFocus
        end
  end;
end;

procedure TfrmModuloReporteGerencial.zq_EquiposFueradeOperacionAfterInsert(
  DataSet: TDataSet);
begin
  TZQuery(DataSet).FieldByName('dIdFecha').AsDateTime := dIdFecha.Date;
  TZQuery(DataSet).FieldByName('sContrato').AsString := Global_Contrato;
  TZQuery(DataSet).FieldByName('sTipoRecurso').AsString := 'Equipo';
  TZQuery(DataSet).FieldByName('sTipoFaltante').AsString := 'FO';
end;

procedure TfrmModuloReporteGerencial.zq_EquiposFueradeOperacionsIdRecursoChange(
  Sender: TField);
Var
  sDescripcion: String;
begin
  if not zq_EquiposFueraDeOperacion.FieldByName('sIdRecurso').IsNull then
  begin
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('select sIdEquipo, sDescripcion from equipos ' +
      'WHERE sContrato = :ContratoBarco And sIdEquipo = :Equipo');
    Connection.QryBusca.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
    Connection.QryBusca.ParamByName('Equipo').AsString := zq_EquiposFueraDeOperacionsIdRecurso.Text;
    Connection.QryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
//      zq_EquiposFueraDeOperacion.FieldValues['sIdRecurso'] := Connection.QryBusca.FieldValues['sIdPersonal'];
      zq_EquiposFueraDeOperacion.FieldValues['sEquipo'] := Connection.QryBusca.FieldValues['sDescripcion'];
    end
    else
      if not zq_EquiposFueraDeOperacion.FieldByName('sIdRecurso').IsNull then
        if Trim(zq_EquiposFueraDeOperacion.FieldByName('sIdRecurso').AsString) <> '' then
        begin
          sDescripcion := '%' + Trim(zq_EquiposFueraDeOperacion.FieldValues['sIdRecurso']) + '%';
          BuscaObjeto.Active := False;
          ListaObjeto.Columns.Clear;

          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[0].FieldName := 'sIdEquipo';
          ListaObjeto.Columns[0].Width := 80;
          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[1].FieldName := 'sDescripcion';
          ListaObjeto.Columns[1].Width := 500;

          BuscaObjeto.SQL.Clear;
          BuscaObjeto.SQL.Add('Select sIdEquipo, sDescripcion FROM equipos WHERE ' +
            'sContrato = :ContratoBarco And sDescripcion Like :Descripcion');
          BuscaObjeto.Params.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          BuscaObjeto.Params.ParamByName('Descripcion').AsString := sDescripcion;
          BuscaObjeto.Open;
               // yavienedeRegreso := 'No' ;
          Panel.Visible := True;
          Panel.Height  := 200;
          Panel.Width   := 400;
          ListaObjeto.SetFocus
        end
  end;

end;

end.
