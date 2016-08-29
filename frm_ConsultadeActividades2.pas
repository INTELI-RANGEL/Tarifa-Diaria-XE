unit frm_ConsultadeActividades2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, StdCtrls, DBCtrls, ComCtrls, ExtCtrls, DB,
  Mask, Grids, DBGrids, global, Buttons, frxClass, frxDBSet, ToolWin,
  Menus, RXDBCtrl, utilerias, Newpanel, UnitTBotonesPermisos,
  ZAbstractRODataset, ZDataset, RxLookup, rxToolEdit, rxCurrEdit, udbgrid, unitexcepciones, UnitValidaTexto, UFunctionsGHH,
  cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxContainer,
  cxEdit, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel,
  dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver,
  dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue, cxGroupBox, cxStyles, dxSkinscxPCPainter, cxCustomData,
  cxFilter, cxData, cxDataStorage, cxNavigator, cxDBData, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGridCustomView, cxClasses, cxGridLevel,
  cxGrid, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit,
  cxDBLookupComboBox, cxLabel;

type
  TfrmConsultaActividad2 = class(TForm)
    ds_bitacora: TDataSource;
    ds_actividadesxorden: TDataSource;
    sbPaquete: TStatusBar;
    btnSalir: TBitBtn;
    ds_Resumen: TDataSource;
    PopupPrincipal: TPopupMenu;
    ComentariosAdicionales: TMenuItem;
    N4: TMenuItem;
    Cut1: TMenuItem;
    N3: TMenuItem;
    Salir1: TMenuItem;
    ds_Historico: TDataSource;
    PanelHistorico: TGroupBox;
    Grid_Historico: TRxDBGrid;
    HistorialdeSuministros1: TMenuItem;
    mnFichaTecnica: TMenuItem;
    imgNotas: TImage;
    ActividadesxOrden: TZReadOnlyQuery;
    AvGeneral: TZReadOnlyQuery;
    ResumendeAlcances: TZReadOnlyQuery;
    Bitacora: TZReadOnlyQuery;
    Historico: TZReadOnlyQuery;
    rDiario: TfrxReport;
    ds_Partidas: TDataSource;
    QryPartidas: TZReadOnlyQuery;
    PopupMenu2: TPopupMenu;
    VerocultarHistorialdeSuministros1: TMenuItem;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label1: TLabel;
    Label10: TLabel;
    Label15: TLabel;
    Label11: TLabel;
    tdPonderado: TCurrencyEdit;
    tdVentaMN: TCurrencyEdit;
    tdCantidadAnexo: TCurrencyEdit;
    tdInstalado: TCurrencyEdit;
    tdPendiente: TCurrencyEdit;
    tdExcedente: TCurrencyEdit;
    tsWbs: TEdit;
    tsMedida: TEdit;
    qryPartidasDelAnexo: TZReadOnlyQuery;
    dsPartidasDelAnexo: TDataSource;
    cxGroupBox1: TcxGroupBox;
    cxLabel1: TcxLabel;
    gridActividadesAnexo: TcxGrid;
    grdActLvl: TcxGridLevel;
    grdActividades: TcxGridDBTableView;
    grdActividadessNumeroActividad: TcxGridDBColumn;
    grdActividadesdCantidadAnexo: TcxGridDBColumn;
    grdActividadesmDescripcion: TcxGridDBColumn;
    tsNumeroActividad: TcxLookupComboBox;
    grpDetalle: TcxGroupBox;
    Grid_Resumen: TcxGrid;
    grdResulvl: TcxGridLevel;
    grdResumen: TcxGridDBTableView;
    grdResumensDescripcion: TcxGridDBColumn;
    grdResumendPonderado: TcxGridDBColumn;
    grdResumendCantidad: TcxGridDBColumn;
    grpActividadAnexo: TcxGroupBox;
    cxGroupBox2: TcxGroupBox;
    GridActividades: TcxGrid;
    grd_actlvl: TcxGridLevel;
    Grid_Actividades: TcxGridDBTableView;
    Grid_ActividadessNumeroOrden: TcxGridDBColumn;
    Grid_ActividadessWbs: TcxGridDBColumn;
    Grid_ActividadessDescripcion: TcxGridDBColumn;
    Grid_ActividadesdCantidad: TcxGridDBColumn;
    Grid_ActividadessMedida: TcxGridDBColumn;
    Grid_ActividadesdVentaMN: TcxGridDBColumn;
    Grid_ActividadesdInstaladoTotal: TcxGridDBColumn;
    Grid_ActividadesdPendiente: TcxGridDBColumn;
    Grid_ActividadesdExcedente: TcxGridDBColumn;
    Grid_ActividadesdPonderado: TcxGridDBColumn;
    cxGroupBox3: TcxGroupBox;
    cxGrid1: TcxGrid;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1DBTableView1dIdFecha: TcxGridDBColumn;
    cxGrid1DBTableView1iIdDiario: TcxGridDBColumn;
    cxGrid1DBTableView1sNumeroReporte: TcxGridDBColumn;
    cxGrid1DBTableView1sDescripcionTurno: TcxGridDBColumn;
    cxGrid1DBTableView1sTitulo: TcxGridDBColumn;
    cxGrid1DBTableView1dCantidad: TcxGridDBColumn;
    cxGrid1DBTableView1dAvance: TcxGridDBColumn;
    cxGrid1DBTableView1sIsometrico: TcxGridDBColumn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure tsNumeroActividadKeyPress(Sender: TObject; var Key: Char);
    procedure btnSalirClick(Sender: TObject);
    procedure BitacoraCalcFields(DataSet: TDataSet);
    procedure ComentariosAdicionalesClick(Sender: TObject);
    procedure imgNotasDblClick(Sender: TObject);
    procedure HistorialdeSuministros1Click(Sender: TObject);
    procedure grid_bitacoraDblClick(Sender: TObject);
    procedure rDiarioGetValue(const VarName: string; var Value: Variant);
    procedure mnFichaTecnicaClick(Sender: TObject);
    procedure ActividadesxOrdenCalcFields(DataSet: TDataSet);
    procedure ActividadesxOrdenAfterScroll(DataSet: TDataSet);
    procedure VerocultarHistorialdeSuministros1Click(Sender: TObject);
    procedure qryPartidasDelAnexoAfterScroll(DataSet: TDataSet);
    procedure tsNumeroActividadPropertiesCloseUp(Sender: TObject);
  private
    sMenuP: string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmConsultaActividad2: TfrmConsultaActividad2;
  sOpcion: string;
  BotonPermiso: TBotonesPermisos;
implementation

uses frm_comentariosxanexo;

{$R *.dfm}

procedure TfrmConsultaActividad2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  BotonPermiso.Free;
  action := cafree;
end;

procedure TfrmConsultaActividad2.FormShow(Sender: TObject);
begin
//  gridActividadesAnexo.Visible := false;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cConsulta1', PopupPrincipal);
  BotonPermiso.permisosBotones(nil);
  try
    sMenuP := stMenu;
    connection.configuracion.refresh;
    QryPartidas.Active := False;
    QryPartidas.Params.ParamByName('Contrato').DataType := ftString;
    QryPartidas.Params.ParamByName('Contrato').Value := global_contrato;
    QryPartidas.Params.ParamByName('Convenio').DataType := ftString;
    QryPartidas.Params.ParamByName('Convenio').Value := global_convenio;
    QryPartidas.Open;
    Bitacora.Active := False;
    tsNumeroActividad.SetFocus;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Consulta de Partidas por Anexo', 'Al iniciar el formulario', 0);
    end;
  end;

end;

procedure TfrmConsultaActividad2.tsNumeroActividadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    if GridActividades.Visible then
      gridActividadesAnexo.SetFocus
end;

procedure TfrmConsultaActividad2.tsNumeroActividadPropertiesCloseUp(
  Sender: TObject);
begin
  qryPartidasDelAnexo.Active := False;
  qryPartidasDelAnexo.Params.ParamByName('contrato').DataType := ftString;
  qryPartidasDelAnexo.Params.ParamByName('contrato').Value := global_contrato;
  qryPartidasDelAnexo.Params.ParamByName('convenio').DataType := ftString;
  qryPartidasDelAnexo.Params.ParamByName('convenio').Value := global_convenio;
  qryPartidasDelAnexo.Params.ParamByName('actividad').DataType := ftString;
  qryPartidasDelAnexo.Params.ParamByName('actividad').Value := tsNumeroActividad.EditValue;
  qryPartidasDelAnexo.Open;

  {if qryPartidasDelAnexo.RecordCount > 1 then
    gridActividadesAnexo.Visible := true
  else
    gridActividadesAnexo.Visible := false;}
end;

procedure TfrmConsultaActividad2.VerocultarHistorialdeSuministros1Click(
  Sender: TObject);
begin
  PanelHistorico.Visible := not PanelHistorico.Visible;
end;

procedure TfrmConsultaActividad2.btnSalirClick(Sender: TObject);
begin
  close
end;

procedure TfrmConsultaActividad2.BitacoraCalcFields(DataSet: TDataSet);
begin
  if Bitacora.FieldValues['sIdTipoMovimiento'] = Connection.configuracion.FieldValues['sTipoAlcance'] then
  begin
    Connection.QryBusca.Active := False;
    Connection.QryBusca.SQL.Clear;
    Connection.QryBusca.SQL.Add('Select a.sDescripcion From bitacoradealcances b INNER JOIN alcancesxactividad a ON ' +
      '(b.sContrato = a.sContrato And b.sNumeroActividad = a.sNumeroActividad And b.iFase = a.iFase) ' +
      'Where b.sContrato = :Contrato And b.dIdFecha = :Fecha And b.iIdDiario = :Diario');
    Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.QryBusca.Params.ParamByName('Fecha').DataType := ftDate;
    Connection.QryBusca.Params.ParamByName('Fecha').Value := Bitacora.FieldValues['dIdFecha'];
    Connection.QryBusca.Params.ParamByName('Diario').DataType := ftInteger;
    Connection.QryBusca.Params.ParamByName('Diario').Value := Bitacora.FieldValues['iIdDiario'];
    Connection.QryBusca.Open;
    {if Connection.QryBusca.RecordCount > 0 then
      BitacorasTitulo.Text := Connection.QryBusca.FieldValues['sDescripcion']}
  end
  {else
    BitacorasTitulo.Text := 'VOLUMEN DE OBRA';}
end;


procedure TfrmConsultaActividad2.ComentariosAdicionalesClick(
  Sender: TObject);
begin
  if grid_Actividades.DataController.DataSource.DataSet.IsEmpty = false then
  begin
    if grid_Actividades.DataController.DataSource.DataSet.RecordCount > 0 then
    begin
      global_partida := tsNumeroActividad.Text;
      Application.CreateForm(TfrmComentariosxAnexo, frmComentariosxAnexo);
      frmComentariosxAnexo.show
    end
    else
      showmessage('no existen partidas para aplicar comentarios');
  end
  else
    showmessage('no existen partidas para aplicar comentarios');
end;

procedure TfrmConsultaActividad2.imgNotasDblClick(Sender: TObject);
begin
  ComentariosAdicionales.Click
end;

procedure TfrmConsultaActividad2.HistorialdeSuministros1Click(
  Sender: TObject);
begin
  PanelHistorico.Visible := not PanelHistorico.Visible;
end;

procedure TfrmConsultaActividad2.grid_bitacoraDblClick(Sender: TObject);
begin
  try
    if tsNumeroActividad.Text <> '' then
      if Bitacora.RecordCount > 0 then
        if connection.contrato.FieldValues['sTipoObra'] = 'PROGRAMADA' then
          procReporteDiarioCotemarProg(Bitacora.FieldValues['sContrato'], Bitacora.FieldValues['sNumeroOrden'], Bitacora.FieldValues['sNumeroReporte'], Bitacora.FieldValues['sIdTurno'], Bitacora.FieldValues['sIdConvenio'], Bitacora.FieldValues['dIdFecha'], '', frmConsultaActividad2, rDiario.OnGetValue, nil)
        else
          if connection.contrato.FieldValues['sTipoObra'] = 'OPTATIVA' then
            procReporteDiarioCotemarOpt(Bitacora.FieldValues['sContrato'], Bitacora.FieldValues['sNumeroOrden'], Bitacora.FieldValues['sNumeroReporte'], Bitacora.FieldValues['sIdTurno'], Bitacora.FieldValues['sIdConvenio'], Bitacora.FieldValues['dIdFecha'], '', frmConsultaActividad2, rDiario.OnGetValue)
          else
            if connection.contrato.FieldValues['sTipoObra'] = 'BARCO' then
              procReporteBarco(Bitacora.FieldValues['sContrato'], Bitacora.FieldValues['sNumeroOrden'], Bitacora.FieldValues['sIdTurno'], Bitacora.FieldValues['dIdFecha'], frmConsultaActividad2, rDiario.OnGetValue)
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Consulta x Partida Anexo', 'Al imprimir reporte diario', 0);
    end;
  end;
end;

procedure TfrmConsultaActividad2.rDiarioGetValue(const VarName: string;
  var Value: Variant);
begin
  if CompareText(VarName, 'DESCRIPCION_CORTA') = 0 then
    Value := sDiarioDescripcionCorta;

  if CompareText(VarName, 'IMPRIME_AVANCES') = 0 then
    Value := sDiarioComentario;

  if CompareText(VarName, 'sNewTexto') = 0 then
    Value := sDiarioTitulo;

  if CompareText(VarName, 'PERIODO') = 0 then
    Value := sDiarioPeriodo;


  if CompareText(VarName, 'SUPERINTENDENTE') = 0 then
    if Bitacora.FieldValues['sOrigenTierra'] = 'No' then
      Value := sSuperIntendente
    else
      Value := sSuperIntendentePatio;

  if CompareText(VarName, 'SUPERVISOR') = 0 then
    if Bitacora.FieldValues['sOrigenTierra'] = 'No' then
      Value := sSupervisor
    else
      Value := sSupervisorPatio;

  if CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
    if Bitacora.FieldValues['sOrigenTierra'] = 'No' then
      Value := sSupervisorTierra
    else
      Value := sResidente;

  if CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
    if Bitacora.FieldValues['sOrigenTierra'] = 'No' then
      Value := sPuestoSuperIntendente
    else
      Value := sPuestoSuperIntendentePatio;

  if CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
    if Bitacora.FieldValues['sOrigenTierra'] = 'No' then
      Value := sPuestoSupervisor
    else
      Value := sPuestoSupervisorPatio;

  if CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
    if Bitacora.FieldValues['sOrigenTierra'] = 'No' then
      Value := sPuestoSupervisorTierra
    else
      Value := sPuestoResidente;

  if CompareText(VarName, 'DESCRIPCION_ORDEN') = 0 then
    Value := mDescripcionOrden;
  if CompareText(VarName, 'PLATAFORMA') = 0 then
    Value := sPlataformaOrden;
  if CompareText(VarName, 'JORNADAS_SUSPENDIDAS') = 0 then
    Value := sJornadasSuspendidas;
  if CompareText(VarName, 'TURNO') = 0 then
    Value := sDescripcionTurno;

  if CompareText(VarName, 'REAL_ANTERIOR') = 0 then
    Value := dRealGlobalAnterior;
  if CompareText(VarName, 'REAL_ACTUAL') = 0 then
    Value := dRealGlobalActual;
  if CompareText(VarName, 'REAL_ACUMULADO') = 0 then
    Value := dRealGlobalAcumulado;
  if CompareText(VarName, 'PROGRAMADO_ANTERIOR') = 0 then
    Value := dProgramadoGlobalAnterior;
  if CompareText(VarName, 'PROGRAMADO_ACTUAL') = 0 then
    Value := dProgramadoGlobalActual;
  if CompareText(VarName, 'PROGRAMADO_ACUMULADO') = 0 then
    Value := dProgramadoGlobalAcumulado;


  if CompareText(VarName, 'REAL_ANTERIOR_MULTIPLE') = 0 then
    Value := dRealOrdenAnterior;
  if CompareText(VarName, 'REAL_ACTUAL_MULTIPLE') = 0 then
    Value := dRealOrdenActual;
  if CompareText(VarName, 'REAL_ACUMULADO_MULTIPLE') = 0 then
    Value := dRealOrdenAcumulado;
  if CompareText(VarName, 'PROGRAMADO_ANTERIOR_MULTIPLE') = 0 then
    Value := dProgramadoOrdenAnterior;
  if CompareText(VarName, 'PROGRAMADO_ACTUAL_MULTIPLE') = 0 then
    Value := dProgramadoOrdenActual;
  if CompareText(VarName, 'PROGRAMADO_ACUMULADO_MULTIPLE') = 0 then
    Value := dProgramadoOrdenAcumulado;
end;

procedure TfrmConsultaActividad2.mnFichaTecnicaClick(Sender: TObject);
begin
  try
    if grid_Actividades.DataController.DataSource.Dataset.isempty = false then
    begin
      if grid_Actividades.DataController.DataSource.DataSet.RecordCount > 0 then
      begin
        if tsNumeroActividad.Text <> '' then
       //<ROJAS>
          procFichaTecnica(global_contrato, global_convenio, tsNumeroActividad.Text, frmConsultaActividad2, connection.configuracion.fieldbyname('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
       //
      end
      else
        showmessage('no existen partidas para generar ficha técnica ');
    end
    else
      showmessage('no existen partidas para generar ficha técnica ');
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Consulta de Partidas por Anexo', 'Al imprimir ficha tecnica', 0);
    end;
  end;
end;



procedure TfrmConsultaActividad2.qryPartidasDelAnexoAfterScroll(
  DataSet: TDataSet);
begin

//  connection.QryBusca.Active := False;
//  connection.QryBusca.SQL.Clear;
//  connection.QryBusca.SQL.Add('Select sWbs, sNumeroActividad, mDescripcion, dCantidadAnexo, sMedida, dVentaMN, dPonderado, dInstalado, dExcedente from actividadesxanexo ' +
//    'Where sContrato = :contrato and sIdConvenio = :convenio and sNumeroActividad = :actividad and sTipoActividad = "Actividad"');
//  connection.QryBusca.Params.ParamByName('contrato').DataType := ftString;
//  connection.QryBusca.Params.ParamByName('contrato').Value := global_contrato;
//  connection.QryBusca.Params.ParamByName('convenio').DataType := ftString;
//  connection.QryBusca.Params.ParamByName('convenio').Value := global_convenio;
//  connection.QryBusca.Params.ParamByName('actividad').DataType := ftString;
//  connection.QryBusca.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;

//
//  try
//    connection.QryBusca.Open;
//  except
//    on e: exception do begin
//      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Consulta de Partidas por Anexo', 'Al consultar la tabla ActividadesxAnexo', 0);
//    end;
//  end;

  try
    if qryPartidasDelAnexo.RecordCount > 0 then
    begin
      global_partida := tsNumeroActividad.Text;
      tdCantidadAnexo.Value := qryPartidasDelAnexo.FieldValues['dCantidadAnexo'];
      tdInstalado.Value := qryPartidasDelAnexo.FieldValues['dInstalado'];
      tdPendiente.Value := qryPartidasDelAnexo.FieldValues['dCantidadAnexo'] - qryPartidasDelAnexo.FieldValues['dInstalado'];
      tdExcedente.Value := qryPartidasDelAnexo.FieldValues['dExcedente'];
      tsWbs.Text := qryPartidasDelAnexo.FieldValues['sWbs'];
      tdVentaMN.Value := qryPartidasDelAnexo.FieldValues['dVentaMN'];
      tsMedida.Text := qryPartidasDelAnexo.FieldValues['sMedida'];
      tdPonderado.Value := qryPartidasDelAnexo.FieldValues['dPonderado'];
      frmConsultaActividad2.Hint := qryPartidasDelAnexo.FieldValues['mDescripcion'];

      Historico.Active := False;
      Historico.Params.ParamByName('Contrato').DataType := ftString;
      Historico.Params.ParamByName('Contrato').Value := global_contrato;
      Historico.Params.ParamByName('Actividad').DataType := ftString;
      Historico.Params.ParamByName('Actividad').Value := tsNumeroActividad.Text;
      Historico.Params.ParamByName('wbs').DataType := ftString;
      Historico.Params.ParamByName('wbs').Value := tsWbs.Text;
      Historico.Open;

      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select sContrato From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad ' +
        ' and sWbs = :wbs ');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Params.ParamByName('actividad').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
      Connection.QryBusca.Params.ParamByName('wbs').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('wbs').Value := tsWbs.Text;
      Connection.QryBusca.Open;
      if Connection.QryBusca.RecordCount > 0 then
        imgNotas.Visible := True;

      ActividadesxOrden.Active := False;
      ActividadesxOrden.Params.ParamByName('contrato').DataType := ftString;
      ActividadesxOrden.Params.ParamByName('contrato').Value := global_contrato;
      ActividadesxOrden.Params.ParamByName('convenio').DataType := ftString;
      ActividadesxOrden.Params.ParamByName('convenio').Value := global_convenio;
      ActividadesxOrden.Params.ParamByName('actividad').DataType := ftString;
      ActividadesxOrden.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
      ActividadesxOrden.Params.ParamByName('wbs').DataType := ftString;
      ActividadesxOrden.Params.ParamByName('wbs').Value := tsWbs.Text;
      ActividadesxOrden.Open;

      ResumendeAlcances.Active := False;
      ResumendeAlcances.Params.ParamByName('contrato').DataType := ftString;
      ResumendeAlcances.Params.ParamByName('contrato').Value := global_contrato;
      ResumendeAlcances.Params.ParamByName('Actividad').DataType := ftString;
      ResumendeAlcances.Params.ParamByName('Actividad').Value := tsNumeroActividad.Text;
      ResumendeAlcances.Params.ParamByName('wbs').DataType := ftString;
      ResumendeAlcances.Params.ParamByName('wbs').Value := tsWbs.Text;
      ResumendeAlcances.Open;
    end
    else
    begin
      global_partida := '';
      tdCantidadAnexo.Value := 0;
      tdInstalado.Value := 0;
      tdPendiente.Value := 0;
      tdExcedente.Value := 0;
      tsWbs.Text := '';
      tdVentaMN.Value := 0;
      tsMedida.Text := '';
      tdPonderado.Value := 0;
      frmConsultaActividad2.Hint := '';
      sbPaquete.Panels.Items[1].Text := '0';
      PanelHistorico.Visible := False;

      Historico.Active := False;
      Historico.Params.ParamByName('Contrato').DataType := ftString;
      Historico.Params.ParamByName('Contrato').Value := global_contrato;
      Historico.Params.ParamByName('Actividad').DataType := ftString;
      Historico.Params.ParamByName('Actividad').Value := '';
      Historico.Params.ParamByName('wbs').DataType := ftString;
      Historico.Params.ParamByName('wbs').Value := '';
      Historico.Open;

      imgNotas.Visible := False;

      ActividadesxOrden.Active := False;
      ActividadesxOrden.Params.ParamByName('contrato').DataType := ftString;
      ActividadesxOrden.Params.ParamByName('contrato').Value := global_contrato;
      ActividadesxOrden.Params.ParamByName('convenio').DataType := ftString;
      ActividadesxOrden.Params.ParamByName('convenio').Value := global_convenio;
      ActividadesxOrden.Params.ParamByName('actividad').DataType := ftString;
      ActividadesxOrden.Params.ParamByName('actividad').Value := '';
      ActividadesxOrden.Params.ParamByName('wbs').DataType := ftString;
      ActividadesxOrden.Params.ParamByName('wbs').Value := '';
      ActividadesxOrden.Open;

      ResumendeAlcances.Active := False;
      ResumendeAlcances.Params.ParamByName('contrato').DataType := ftString;
      ResumendeAlcances.Params.ParamByName('contrato').Value := global_contrato;
      ResumendeAlcances.Params.ParamByName('Actividad').DataType := ftString;
      ResumendeAlcances.Params.ParamByName('Actividad').Value := '';
      ResumendeAlcances.Params.ParamByName('wbs').DataType := ftString;
      ResumendeAlcances.Params.ParamByName('wbs').Value := '';
      ResumendeAlcances.Open;

      Bitacora.Active := False;
      Bitacora.Params.ParamByName('contrato').DataType := ftString;
      Bitacora.Params.ParamByName('contrato').Value := global_contrato;
      Bitacora.Params.ParamByName('orden').DataType := ftString;
      Bitacora.Params.ParamByName('orden').Value := ActividadesxOrden.FieldValues['sNumeroOrden'];
      Bitacora.Params.ParamByName('wbs').DataType := ftString;
      Bitacora.Params.ParamByName('wbs').Value := '';
      Bitacora.Params.ParamByName('actividad').DataType := ftString;
      Bitacora.Params.ParamByName('actividad').Value := '';
      Bitacora.Params.ParamByName('wbs').DataType := ftString;
      Bitacora.Params.ParamByName('wbs').Value := '';
      Bitacora.Open;

      sbPaquete.Panels.Items[1].Text := '0';
    end

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Consulta de Partidas por Anexo', 'Al obtener el número de concepto', 0);
    end;
  end;
end;

procedure TfrmConsultaActividad2.ActividadesxOrdenCalcFields(
  DataSet: TDataSet);
begin
  {if ActividadesxOrden.FieldValues['dCantidad'] >= ActividadesxOrden.FieldValues['dInstalado'] then
    ActividadesxOrdendPendiente.Value := ActividadesxOrden.FieldValues['dCantidad'] - ActividadesxOrden.FieldValues['dInstalado']
  else
    ActividadesxOrdendPendiente.Value := 0;
  ActividadesxOrdendInstaladoTotal.Value := ActividadesxOrden.FieldValues['dExcedente'] + ActividadesxOrden.FieldValues['dInstalado'];
}
  connection.QryBusca2.Active := False;
  connection.QryBusca2.SQL.Clear;
  connection.QryBusca2.SQL.Add('select mDescripcion from actividadesxorden Where sContrato = :contrato and sNumeroOrden = :orden and sIdConvenio = :convenio and sWbs = :wbs and sTipoActividad = "Paquete"');
  connection.QryBusca2.Params.ParamByName('contrato').DataType := ftString;
  connection.QryBusca2.Params.ParamByName('contrato').Value := global_contrato;
  connection.QryBusca2.Params.ParamByName('convenio').DataType := ftString;
  connection.QryBusca2.Params.ParamByName('convenio').Value := global_convenio;
  connection.QryBusca2.Params.ParamByName('orden').DataType := ftString;
  connection.QryBusca2.Params.ParamByName('orden').Value := ActividadesxOrden.FieldValues['sNumeroOrden'];
  connection.QryBusca2.Params.ParamByName('wbs').DataType := ftString;
  connection.QryBusca2.Params.ParamByName('wbs').Value := ActividadesxOrden.FieldValues['sWbsAnterior'];
  connection.QryBusca2.Open;
  {if connection.QryBusca2.RecordCount > 0 then
    ActividadesxOrdensDescripcion.Text := connection.QryBusca2.FieldValues['mDescripcion'];
 }
end;

procedure TfrmConsultaActividad2.ActividadesxOrdenAfterScroll(
  DataSet: TDataSet);
begin
  if ActividadesxOrden.RecordCount > 0 then
  begin
    GridActividades.Hint := ActividadesxOrden.FieldValues['mDescripcion'];

    Bitacora.Active := False;
    Bitacora.Params.ParamByName('contrato').DataType   := ftString;
    Bitacora.Params.ParamByName('contrato').Value      := global_contrato;
    Bitacora.Params.ParamByName('orden').DataType      := ftString;
    Bitacora.Params.ParamByName('orden').Value         := ActividadesxOrden.FieldValues['sNumeroOrden'];
    Bitacora.Params.ParamByName('wbs').DataType        := ftString;
    Bitacora.Params.ParamByName('wbs').Value           := ActividadesxOrden.FieldValues['sWbs'];
    Bitacora.Params.ParamByName('actividad').DataType  := ftString;
    Bitacora.Params.ParamByName('actividad').Value     := ActividadesxOrden.FieldValues['sNumeroActividad'];
    Bitacora.Open;

    AvGeneral.Active := False;
    AvGeneral.Params.ParamByName('contrato').DataType := ftString;
    AvGeneral.Params.ParamByName('contrato').Value := global_contrato;
    AvGeneral.Params.ParamByName('Orden').DataType := ftString;
    AvGeneral.Params.ParamByName('Orden').Value := ActividadesxOrden.FieldValues['sNumeroOrden'];
    AvGeneral.Params.ParamByName('Wbs').DataType := ftString;
    AvGeneral.Params.ParamByName('Wbs').Value := ActividadesxOrden.FieldValues['sWbs'];
    AvGeneral.Params.ParamByName('Actividad').DataType := ftString;
    AvGeneral.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad'];
    AvGeneral.Open;
    if AvGeneral.RecordCount > 0 then
      sbPaquete.Panels.Items[1].Text := avGeneral.FieldValues['dAvance']
    else
      sbPaquete.Panels.Items[1].Text := '0';
  end
  else
    GridActividades.Hint := '';
end;

end.

